"""2023 한국은행 산업연관표: 브라우저(Pyodide)에서 실행하는 분석 엔진.

원본 Excel 파일의 계산된 셀 값(data_only)을 이용한다.
산업 간 거래표 탐색과 국산품 최종수요 변화의 3대 유발효과를 계산한다.
"""
import io
import json

import numpy as np
import openpyxl

book = {}
codes = []
names = []
sheets = []
L = None
M = None
V = None
residuals = {}
A_d_data = None
A_m_data = None
v_data = None


def _emit(obj):
    return json.dumps(obj, ensure_ascii=False, allow_nan=False)


def _code(value):
    """원본의 앞자리 0 보존: Excel에서는 코드가 문자열로 기록되어 있다."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _matrix(sheet_name):
    rows = book[sheet_name]
    by_row = {_code(row[0]): row for row in rows[6:] if len(row) > 1 and row[0] is not None}
    by_col = {_code(value): j for j, value in enumerate(rows[4]) if value is not None}
    try:
        m = np.array([[by_row[r][by_col[c]] for c in codes] for r in codes], dtype=float)
    except (KeyError, IndexError, TypeError, ValueError) as exc:
        raise ValueError(f"{sheet_name}: 산업별 행·열의 코드가 맞지 않습니다 ({exc}).") from exc
    if m.shape != (len(codes), len(codes)) or not np.isfinite(m).all():
        raise ValueError(f"{sheet_name}: 계수행렬에 결측치 또는 비정상 값이 있습니다.")
    return m


def initialize(raw, expected_count):
    """Excel을 로드하며 분류별로 산업 코드를 읽고 행렬 관계를 검증한다."""
    global book, codes, names, sheets, L, M, V, residuals, A_d_data, A_m_data, v_data
    required = (
        '총투입계수(A)', '국산투입계수(Ad)', '수입투입계수(Am)',
        '생산유발계수', '수입유발계수', '부가가치유발계수',
    )
    workbook = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    try:
        new_book = {name: list(workbook[name].values) for name in workbook.sheetnames}
    finally:
        workbook.close()
    missing = [name for name in required if name not in new_book]
    if missing:
        raise ValueError('필수 Excel 시트가 없습니다: ' + ', '.join(missing))
    book = new_book
    sheets = list(book)
    n = int(expected_count)
    base = book['국산투입계수(Ad)']
    codes = [_code(value) for value in base[4][2:2+n]]
    sector_names = {_code(row[0]): str(row[1]).strip() for row in base[6:]
                    if len(row) > 1 and row[0] is not None and row[1] is not None}
    if len(codes) != n or len(set(codes)) != n or not all(c in sector_names for c in codes):
        raise ValueError(f'산업코드가 예상한 {n}개 부문과 일치하지 않습니다.')
    names = [sector_names[c] for c in codes]
    L = _matrix('생산유발계수')
    M = _matrix('수입유발계수')
    V = _matrix('부가가치유발계수')
    A_d = _matrix('국산투입계수(Ad)')
    A_m = _matrix('수입투입계수(Am)')
    total = book['총투입계수(A)']
    value_added = next((row for row in total if row and _code(row[0]) == '9690'), None)
    if value_added is None:
        raise ValueError('부가가치계수 행(9690)을 찾지 못했습니다.')
    cindex = {_code(c): i for i, c in enumerate(total[4]) if c is not None}
    v = np.array([value_added[cindex[c]] for c in codes], dtype=float)
    A_d_data, A_m_data, v_data = A_d, A_m, v
    identity = np.eye(n)
    residuals = {
        '생산 (I − Ad)L = I': float(np.max(np.abs((identity-A_d) @ L-identity))),
        '수입 AmL = M': float(np.max(np.abs(A_m @ L-M))),
        '부가가치 diag(v)L = V': float(np.max(np.abs(v[:, None] * L-V))),
    }
    if not all(np.isfinite(x) and x < 1e-7 for x in residuals.values()):
        raise ValueError('입출력 계수 검증에 실패했습니다: '+str(residuals))
    return _emit({'sectors': [{'code': c, 'name': name} for c, name in zip(codes, names)],
                  'sheets': sheets, 'validation': residuals, 'count': n})


def explore(sheet_name, rowcode=""):
    if sheet_name not in book:
        raise ValueError('해당 Excel 시트가 존재하지 않습니다.')
    rows = book[sheet_name]
    opts = [{'code': _code(row[0]), 'name': str(row[1]).strip()}
            for row in rows[6:] if len(row)>1 and row[0] is not None and row[1] is not None]
    if not opts:
        raise ValueError('조회 가능한 행이 없습니다.')
    selected = next((row for row in rows[6:] if row and _code(row[0]) == rowcode), None)
    if selected is None:
        code = opts[0]['code']
        selected = next(row for row in rows[6:] if row and _code(row[0]) == code)
    columns = []
    header = rows[4]
    titles = rows[5]
    for i in range(2, len(header)):
        if header[i] is None:
            continue
        value = selected[i] if i < len(selected) else None
        if isinstance(value, (int, float)):
            value = float(value)
            if not np.isfinite(value):
                value = None
        elif value is not None:
            value = str(value)
        columns.append({'code': _code(header[i]),
                        'name': str(titles[i] or '') if i < len(titles) else '',
                        'value': value})
    monetary = sheet_name.startswith('A표_') or sheet_name.startswith('최종수요')
    return _emit({'rows': opts, 'selected': _code(selected[0]), 'name': str(selected[1]),
                  'unit': '백만원' if monetary else '계수 (무차원)', 'cells': columns})


def simulate(payload_json):
    if L is None:
        raise ValueError('Excel을 먼저 불러오세요.')
    selected = json.loads(payload_json)
    f = np.zeros(len(codes))
    lookup = {c: i for i, c in enumerate(codes)}
    for item in selected:
        code = str(item['code'])
        amount = float(item['amount'])  # 백만원 단위
        if code not in lookup or not np.isfinite(amount) or amount < 0:
            raise ValueError('유효한 산업과 0 이상의 금액을 입력해야 합니다.')
        f[lookup[code]] += amount
    if f.sum() <= 0:
        raise ValueError('하나 이상의 양(+)의 최종수요 증가액을 입력하세요.')
    x = L @ f
    m = M @ f
    v = V @ f
    output = [{'code': code, 'name': name, 'production': float(px),
               'imports': float(im), 'valueAdded': float(va), 'finalDemand': float(fd)}
              for code, name, px, im, va, fd in zip(codes, names, x, m, v, f)]
    output.sort(key=lambda row: row['production'], reverse=True)
    return _emit({'demand': float(f.sum()), 'production': float(x.sum()),
                  'imports': float(m.sum()), 'valueAdded': float(v.sum()),
                  'productionMultiplier': float(x.sum()/f.sum()),
                  'rows': output})


def _check_ready():
    if L is None or A_d_data is None:
        raise ValueError('산업분류 Excel 데이터를 먼저 불러오세요.')


def _demand_vector(items):
    _check_ready()
    f = np.zeros(len(codes), dtype=float)
    lookup = {c: i for i, c in enumerate(codes)}
    for item in items:
        code = str(item['code'])
        amount = float(item['amount'])
        if code not in lookup or not np.isfinite(amount) or amount < 0 or amount > 1e16:
            raise ValueError('산업코드 또는 수요 증가액이 유효하지 않습니다.')
        f[lookup[code]] += amount
    if f.sum() <= 0:
        raise ValueError('수요 증가액이 0보다 커야 합니다.')
    return f


def rank_industries():
    """각 부문에서 개별적으로 국산품 최종수요가 100억원 늘어나는 동일 조건을 비교한다."""
    _check_ready()
    prod = L.sum(axis=0)
    added = V.sum(axis=0)
    imported = M.sum(axis=0)
    records = []
    for i, (code, name) in enumerate(zip(codes, names)):
        records.append({
            'code': code, 'name': name,
            'production': float(prod[i] * 100),
            'valueAdded': float(added[i] * 100),
            'imports': float(imported[i] * 100),
            'productionMultiplier': float(prod[i]),
            'valueAddedMultiplier': float(added[i]),
            'importMultiplier': float(imported[i])
        })
    return _emit({'demand': 100, 'rows': records, 'count': len(records)})


def structure(target_code, matrix_kind='Ad'):
    """행=투입 공급 산업, 열=사용 산업. 부문 수 33 초과면 대상 열의 주요 24개 공급 산업 부분행렬."""
    _check_ready()
    which = {'Ad': A_d_data, 'Am': A_m_data, 'L': L}
    if matrix_kind not in which:
        raise ValueError('지원하지 않는 행렬입니다.')
    lookup = {c: i for i, c in enumerate(codes)}
    if target_code not in lookup:
        raise ValueError('선택 산업코드를 확인하세요.')
    j = lookup[target_code]
    matrix = which[matrix_kind]
    n = len(codes)
    if n <= 33:
        indices = list(range(n))
    else:
        leading = sorted(range(n), key=lambda i: (float(matrix[i, j]), -i), reverse=True)
        indices = [j] + [i for i in leading if i != j][:24]
    labels = [{'code': codes[i], 'name': names[i]} for i in indices]
    sub = matrix[np.ix_(indices, indices)]
    strongest = sorted(range(n), key=lambda i: float(matrix[i, j]), reverse=True)[:12]
    top = [{'code': codes[i], 'name': names[i], 'coefficient': float(matrix[i, j])} for i in strongest]
    return _emit({'kind': matrix_kind, 'target': {'code': codes[j], 'name': names[j]},
                  'selection': '전체 33개 부문' if n <= 33 else '대상 산업 중심 상위 25개 부문(부분행렬)',
                  'labels': labels, 'cells': sub.tolist(), 'maxValue': float(np.max(sub)),
                  'topInputs': top, 'count': n})


def decompose(payload_json):
    """총산출: 직접 Δf, 1차 간접 Ad Δf, 2차 이상 (L−I−Ad)Δf."""
    f = _demand_vector(json.loads(payload_json))
    direct = f
    first = A_d_data @ f
    total = L @ f
    higher = total - direct - first
    # 공개된 행렬은 반올림 계수이므로 계산 오차 1e-6 백만원 이하만 0으로 보정
    higher[np.abs(higher) < 1e-6] = 0
    results = []
    for code, name, direct_i, first_i, higher_i, total_i in zip(codes, names, direct, first, higher, total):
        results.append({'code': code, 'name': name, 'direct': float(direct_i / 100),
                        'first': float(first_i / 100), 'higher': float(higher_i / 100),
                        'production': float(total_i / 100)})
    results.sort(key=lambda x: x['production'], reverse=True)
    return _emit({'demand': float(f.sum() / 100), 'direct': float(direct.sum() / 100),
                  'first': float(first.sum() / 100), 'higher': float(higher.sum() / 100),
                  'production': float(total.sum() / 100), 'rows': results})


def industry_profile(target_code):
    """100억원 표준 충격에 대한 산업 요약. 단순 정태적 파급효과, 정책 성과 아님."""
    _check_ready()
    lookup = {c: i for i, c in enumerate(codes)}
    if target_code not in lookup:
        raise ValueError('선택 산업코드가 없습니다.')
    j = lookup[target_code]
    output = L[:, j]
    suppliers = sorted(range(len(codes)), key=lambda i: A_d_data[i, j], reverse=True)[:8]
    beneficiaries = sorted(range(len(codes)), key=lambda i: output[i], reverse=True)[:8]
    return _emit({'code': codes[j], 'name': names[j], 'demand': 100,
                  'production': float(output.sum() * 100),
                  'valueAdded': float(V[:, j].sum() * 100),
                  'imports': float(M[:, j].sum() * 100),
                  'productionMultiplier': float(output.sum()),
                  'valueAddedMultiplier': float(V[:, j].sum()),
                  'importMultiplier': float(M[:, j].sum()),
                  'firstRound': float(A_d_data[:, j].sum() * 100),
                  'direct': 100,
                  'indirect': float((output.sum()-1) * 100),
                  'topSuppliers': [{'code':codes[i], 'name':names[i], 'coefficient':float(A_d_data[i,j])}
                                  for i in suppliers],
                  'topBeneficiaries': [{'code':codes[i], 'name':names[i], 'production':float(output[i]*100)}
                                       for i in beneficiaries]})


# V4: 원본 국산투입계수(Ad)의 공급/사용 연결을 방향별로 조회한다.
def industry_network(target_code, limit=6, threshold=0.001):
    """행 산업 → 열 산업의 직접 국산 중간재 투입계수 기반 국소 네트워크.

    상류: Ad[i,j] = 선택 j의 산출 1단위에 필요한 공급 i의 중간재
    하류: Ad[j,k] = 사용자 k의 산출 1단위에 필요한 선택 j의 중간재
    두 계수의 분모는 각각 j와 k의 산출로 다르며 실제 거래액은 아니다.
    """
    _check_ready()
    idx = {c: i for i, c in enumerate(codes)}
    if target_code not in idx:
        raise ValueError('선택 산업코드가 없습니다.')
    try:
        limit = int(limit)
        threshold = float(threshold)
    except (TypeError, ValueError) as exc:
        raise ValueError('연결 산업 수와 최소 계수 조건을 확인하세요.') from exc
    if not 1 <= limit <= 12 or not np.isfinite(threshold) or not 0 <= threshold <= 1:
        raise ValueError('연결 개수는 1~12개, 최소 계수는 0~1 범위여야 합니다.')
    j = idx[target_code]
    matrix = A_d_data
    def eligible(value):
        return np.isfinite(value) and value > 0 and value >= threshold
    suppliers = sorted(((i, float(matrix[i, j])) for i in range(len(codes)) if i != j and eligible(matrix[i,j])), key=lambda item:(-item[1], item[0]))
    users = sorted(((k, float(matrix[j, k])) for k in range(len(codes)) if k != j and eligible(matrix[j,k])), key=lambda item:(-item[1], item[0]))
    def make(items):
        return [{'code':codes[i], 'name':names[i], 'coefficient':w} for i,w in items[:limit]]
    return _emit({
        'target':{'code':codes[j], 'name':names[j]},
        'suppliers':make(suppliers), 'users':make(users),
        'supplierMatches':len(suppliers), 'userMatches':len(users),
        'selfCoefficient':float(matrix[j,j]),
        'domesticInputRatio':float(matrix[:,j].sum()),
        'limit':limit, 'threshold':threshold, 'count':len(codes),
        'basis':'2023년 생산자가격 국산투입계수(Ad) · 행=공급, 열=사용'
    })
