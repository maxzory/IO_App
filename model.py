import json, io
import numpy as np
import openpyxl

book = None
valid = False

LEVEL_NAMES = {'dae': '통합대분류', 'jung': '통합중분류', 'so': '통합소분류', 'basic': '기본부문'}
EXPORT_COL = '9140'      # 최종수요 항목: 수출
# 최종수요 7개 항목: 민간소비, 정부소비, 민간고정자본형성, 정부고정자본형성, 재고증감, 귀중품 순취득, 수출
# 일부 산업(선철, 펄프 등)은 부산물 처리 때문에 특정 항목의 유발액이 음수이므로,
# 분모에는 양수인 항목만 더해 비중이 0~100% 안에 들도록 함
FD_COLS = ['9111', '9112', '9121', '9122', '9131', '9132', '9140']


def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)


def clean(s):
    return str(s).replace('　', ' ').replace('\n', ' ').strip()


def initialize(raw, level):
    global book, codes, names, L, M, V, sheets, valid, level_name, diag, exports
    valid = False
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)
    level_name = LEVEL_NAMES.get(level, level)

    # 산업 코드 = 국산투입계수 표에서 열 머리글과 행 머리글에 모두 있는 코드
    # (합계 행·열인 9xxx 코드는 한쪽에만 있어서 자동으로 빠짐)
    base = book['국산투입계수(Ad)']
    row_codes = {str(r[0]) for r in base[6:] if r and r[0] is not None}
    codes = [str(c) for c in base[4][2:] if c is not None and str(c) in row_codes]
    if len(codes) < 10 or len(set(codes)) != len(codes):
        raise ValueError('투입산출표 구조를 인식하지 못했습니다.')
    nm = {str(r[0]): clean(r[1]) for r in base[6:] if r and r[0] is not None}
    names = [nm[c] for c in codes]
    n = len(codes)

    def lookup(s):
        rows = book[s]
        ri = {str(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None}
        ci = {str(c): i for i, c in enumerate(rows[4]) if c is not None}
        return rows, ri, ci

    def matrix(s):
        rows, ri, ci = lookup(s)
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a

    def column(s, col):
        rows, ri, ci = lookup(s)
        return np.array([rows[ri[c]][ci[col]] or 0 for c in codes], dtype=float)

    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    A = matrix('총투입계수(A)'); ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows, ri, ci = lookup('총투입계수(A)')
    v = np.array([rows[ri['9690']][ci[c]] for c in codes], dtype=float)

    # 계수 검증: (I-Ad)L = I, Am·L = M, v·L = V
    residuals = {'생산': float(np.max(np.abs((np.eye(n) - ad) @ L - np.eye(n)))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    if max(residuals.values()) >= 1e-8:
        raise ValueError('계수 검증 실패: ' + str(residuals))
    valid = True

    # ---------- 충격 유형 진단 ----------
    # (1) 수입 충격 취약도 = 영향력계수 × 수입 원료 비중
    colsum = L.sum(axis=0)
    influence = colsum / colsum.mean()                  # 영향력계수 (단위 없음, 평균=1)
    a_sum = A.sum(axis=0)
    import_share = np.divide(am.sum(axis=0), a_sum, out=np.zeros(n), where=a_sum > 0)
    import_vuln = influence * import_share
    # (2) 해외 수요 민감도 = 수출이 유발한 생산액 ÷ 최종수요 항목별 생산유발액 중 양수 항목의 합
    fd_export = column('최종수요 항목별 생산유발액', EXPORT_COL)
    fd_total = sum(np.maximum(column('최종수요 항목별 생산유발액', k), 0) for k in FD_COLS)
    export_share = np.divide(fd_export, fd_total, out=np.zeros(n), where=fd_total > 0)
    exports = column('A표_국산거래표(생산자)', EXPORT_COL)  # 수출액 (백만원)

    x_cut = float(np.median(export_share)); y_cut = float(np.median(import_vuln))
    diag = []
    for i, c in enumerate(codes):
        hi_x = export_share[i] > x_cut; hi_y = import_vuln[i] > y_cut
        kind = ('이중 위험형' if hi_x and hi_y else '수입 위험형' if hi_y
                else '수출 위험형' if hi_x else '안정형')
        diag.append({'code': c, 'name': names[i], 'type': kind,
                     'exportShare': float(export_share[i] * 100),
                     'importShare': float(import_share[i] * 100),
                     'influence': float(influence[i]),
                     'importVuln': float(import_vuln[i]),
                     'exportInduced': float(fd_export[i]),
                     'exports': float(exports[i])})

    return emit({'level': level_name, 'sectors': [{'code': c, 'name': x} for c, x in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals,
                 'diagnosis': diag, 'cuts': {'exportShare': x_cut * 100, 'importVuln': y_cut}})


def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': str(r[0]), 'name': clean(r[1])} for r in rows[6:]
               if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if r and str(r[0]) == rowcode), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if r and r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': clean(rows[5][i] or ''), 'value': val})
    unit = rows[3][0] if len(rows) > 3 and rows[3] and rows[3][0] else None
    unit = clean(unit) if unit and '단위' in str(unit) else '단위 없음(계수·비율) — 행·열 명칭을 확인하세요'
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': clean(chosen[1]),
                 'unit': unit, 'cells': cells})


def simulate(payload):
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload); f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a):
            raise ValueError('수요 변화액은 유한한 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not np.any(f != 0):
        raise ValueError('0이 아닌 수요 변화액을 한 개 이상 입력하세요.')
    x = L @ f; m = M @ f; v = V @ f
    result = [{'code': c, 'name': n, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, n, a, b, d in zip(codes, names, x, m, v)]
    result.sort(key=lambda r: abs(r['production']), reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'rows': result})
