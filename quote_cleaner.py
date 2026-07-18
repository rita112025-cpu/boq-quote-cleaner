"""
quote_cleaner.py — 報價單清洗核心

與 boq_cleaner.py 完全獨立：不 import、不共用任何清洗邏輯。兩者只共用
excel_io.py 這個純 I/O、無清洗邏輯的讀檔小檔案（read_grid() / list_sheets()），
屬於安全的唯讀重用，不構成邏輯耦合，也不需要為此夾帶整份 boq_cleaner.py。

移植來源：合併版 HTML「報價單處理」模組的 JS 清洗邏輯
（HEADER_RULES / guessHeader / scoreHeaderRow / detectHeaderField /
buildMapping / extractData），逐函式對照移植，行為刻意保持一致。

標單（boq_cleaner.py）有多層分類／情境繼承的狀態機，是因為標單本身就是階層式
文件。報價單是相對扁平的品項列表，不需要那套邏輯，因此本檔核心邏輯簡單很多：

    表頭偵測（欄位關鍵字評分，可辨識雙列合併表頭）
        → 逐列擷取（vendor/item/spec/brand/model/unit/qty/price/total/note）
        → 略過空白列／重複表頭列／小計等雜訊列
        → 複價驗算（容差與判定條件沿用與標單相同的慣例，見 validate_total）

【重要】驗證狀態與 boq_cleaner.py 不同：boq_cleaner.py 已用多份真實標單
（原始資料不隨本庫公開）逐列比對回歸驗證。本檔目前只驗證「Python 移植版與原始 JS 版行為
一致」（見 tests/test_quote_cleaner.py），尚未用真實報價單驗證欄位偵測準確度。
建議先用少量真實檔案試跑、核對 batch_log.csv 與複價驗算欄，再大量使用。

批次（非互動）情境下，若表頭信心不足或缺必要欄位，該工作表會被跳過並記錄
原因，不會自動臆測欄位——這點與 boq_cleaner.py 的既有慣例一致。
"""
import re
import unicodedata

# v1.1(F-3)：改從獨立的 excel_io.py 匯入純 I/O 讀檔函式，不再 import 整份
# boq_cleaner.py（474 行，含全部標單清洗邏輯）。報價單批次包現在只需要帶
# excel_io.py（約 30 行），從根本解決兩包各自流通一份 boq_cleaner.py、
# 未來可能版本不同步的風險。
from excel_io import read_grid, list_sheets

# ── 欄位偵測規則（逐項對照 JS HEADER_RULES，權重與別名完全一致） ──
HEADER_RULES = {
    'vendor': {'weight': 20, 'aliases': ['廠商', '廠商名稱', '供應商', '供應商名稱', '報價廠商', 'vendor', 'supplier']},
    'item':   {'weight': 30, 'aliases': ['品名', '品項', '材料名稱', '設備名稱', '項目', '名稱', 'description', 'item']},
    'qty':    {'weight': 15, 'aliases': ['數量', 'qty', 'quantity']},
    'unit':   {'weight': 15, 'aliases': ['單位', 'unit']},
    'price':  {'weight': 15, 'aliases': ['單價', '價格', 'unit price', 'price']},
    'total':  {'weight': 15, 'aliases': ['複價', '合價', '金額', '小計', '總價', 'amount', 'total']},
    'spec':   {'weight': 10, 'aliases': ['規格', '規格說明', '型式', 'spec']},
    'brand':  {'weight': 5, 'aliases': ['品牌', '廠牌', '製造商', 'brand']},
    'model':  {'weight': 5, 'aliases': ['型號', 'model']},
    'note':   {'weight': 3, 'aliases': ['備註', '說明', 'remark', 'note']},
}
REQUIRED_FIELDS = ['item', 'unit', 'qty', 'price']  # vendor 由檔名遞補，不強制對應
OUTPUT_FIELDS = ['vendor', 'item', 'spec', 'brand', 'model', 'unit', 'qty', 'price', 'total', 'note']

_WS_RE = re.compile(r'\s+')
_HEADER_STRIP_RE = re.compile(r'[：:()（）]')
_NEWLINE_RE = re.compile(r'\r?\n')
_CURRENCY_PREFIX_RE = re.compile(r'^[$＄]+')
_CURRENCY_SUFFIX_RE = re.compile(r'元+$')
_NUM_COMMA_WS_RE = re.compile(r'[,\s]')
_NUM_LIKE_RE = re.compile(r'^-?[\d,]+(?:\.\d+)?$')
_NOISE_LEAD_RE = re.compile(r'^(合計|小計|總計|總額|本頁合計|累計)')
_NOISE_ANY_RE = re.compile(r'第\d+頁|承攬廠商簽章|公司章')
_TITLE_HINT_RE = re.compile(r'工程名稱|公司名稱|報價單|客戶名稱|地址')
_SUBTOTAL_HINT_RE = re.compile(r'合計|小計|總計|總額')


# ── 工具函式（對照 JS normalizeHeader/cleanText/toNumber/classifyCell） ──

def normalize_header(v):
    """判斷用正規化：NFKC、trim、小寫、移除空白與部分標點。對照 JS normalizeHeader()。"""
    if v is None:
        return ''
    s = unicodedata.normalize('NFKC', str(v)).strip().lower()
    s = _WS_RE.sub('', s)
    s = _HEADER_STRIP_RE.sub('', s)
    return s


def clean_text(v):
    """輸出用清理：NFKC、換行轉空格、壓縮空白、trim。對照 JS cleanText()。"""
    if v is None:
        return ''
    s = unicodedata.normalize('NFKC', str(v))
    s = _NEWLINE_RE.sub(' ', s)
    s = _WS_RE.sub(' ', s)
    return s.strip()


def to_number(v):
    """對照 JS toNumber()：空值/非數字/NaN/Infinity 一律回 None。

    v1.1 修正（Q-2）：原本用 `[,\\s$＄元]` 移除字串「任何位置」的貨幣符號，
    會把「3元5」這種黏連文字誤解析成 35（靜默產生錯誤數值，比直接解析
    失敗更危險）。改成只移除開頭的 $／＄ 與結尾的「元」（真實報價單常見
    的寫法是「100元」「$100」，貨幣符號夾在數字中間不是合理格式），
    中間仍出現非數字字元時直接視為無法解析，回傳 None。
    """
    if v is None or v == '':
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        if f != f or f in (float('inf'), float('-inf')):
            return None
        return f
    s = unicodedata.normalize('NFKC', str(v)).strip()
    s = _CURRENCY_PREFIX_RE.sub('', s)
    s = _CURRENCY_SUFFIX_RE.sub('', s)
    s = _NUM_COMMA_WS_RE.sub('', s)
    if not s:
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    if n != n or n in (float('inf'), float('-inf')):
        return None
    return n


def classify_cell(v):
    """對照 JS classifyCell()：empty / number / text。"""
    if v is None or v == '':
        return 'empty'
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return 'number'
    if _NUM_LIKE_RE.match(str(v).strip()):
        return 'number'
    return 'text'


def is_blank_row(r):
    return not r or all(v is None or str(v).strip() == '' for v in r)


def is_noise_row(r):
    """小計／頁碼／簽章等雜訊列。對照 JS isNoiseRow()。"""
    parts = [v for v in (r or []) if v is not None]
    t = normalize_header(' '.join(str(p) for p in parts))
    return bool(_NOISE_LEAD_RE.match(t) or _NOISE_ANY_RE.search(t))


def combine_header_rows(rows, indexes):
    """合併多列表頭（雙列合併表頭），逐欄去重後以空白連接。對照 JS combineHeaderRows()。"""
    max_cols = max((len(rows[i]) if i < len(rows) else 0 for i in indexes), default=0)
    out = []
    for c in range(max_cols):
        parts, seen = [], set()
        for i in indexes:
            row = rows[i] if i < len(rows) else []
            v = clean_text(row[c]) if c < len(row) else ''
            nv = normalize_header(v)
            if v and nv not in seen:
                parts.append(v)
                seen.add(nv)
        out.append(' '.join(parts))
    return out


def is_repeated_header(r, headers):
    """整份檔案內每頁重印表頭時，偵測並跳過。對照 JS isRepeatedHeader()。"""
    if not r or not headers:
        return False
    row = combine_header_rows([r], [0])
    matches = sum(
        1 for i, v in enumerate(row)
        if normalize_header(v) and i < len(headers) and normalize_header(v) == normalize_header(headers[i])
    )
    return matches >= 3


def detect_header_field(v):
    """單一儲存格是否像某個欄位名稱：先精確比對，再部分包含比對。對照 JS detectHeaderField()。"""
    n = normalize_header(v)
    if not n:
        return None
    for field, rule in HEADER_RULES.items():
        if any(normalize_header(a) == n for a in rule['aliases']):
            return {'field': field, 'matchType': 'exact', 'score': rule['weight']}
    for field, rule in HEADER_RULES.items():
        for a in rule['aliases']:
            t = normalize_header(a)
            if t and (t in n or n in t):
                return {'field': field, 'matchType': 'partial', 'score': round(rule['weight'] * 0.65)}
    return None


def score_header_row(row):
    """整列表頭候選評分。對照 JS scoreHeaderRow()。"""
    non_empty = [v for v in row if v is not None and str(v).strip() != '']
    if not non_empty:
        return {'score': -100, 'detected': [], 'fields': []}
    detected, used, score = [], set(), 0
    for column_index, v in enumerate(row):
        x = detect_header_field(v)
        if not x:
            continue
        detected.append({'columnIndex': column_index, 'original': v, **x})
        if x['field'] in used:
            score -= 5
        else:
            score += x['score']
            used.add(x['field'])
    numeric_ratio = sum(1 for v in non_empty if classify_cell(v) == 'number') / len(non_empty)
    text = normalize_header(' '.join(str(v) for v in non_empty))
    if numeric_ratio > 0.5:
        score -= 20
    if len(non_empty) < 3:
        score -= 20
    if _TITLE_HINT_RE.search(text):
        score -= 10
    if _SUBTOTAL_HINT_RE.search(text):
        score -= 20
    if len({f for f in ('item', 'qty', 'unit', 'price', 'total') if f in used}) >= 3:
        score += 15
    return {'score': score, 'detected': detected, 'fields': list(used), 'numericRatio': numeric_ratio}


def guess_header(rows, max_scan=50):
    """掃描前 max_scan 列，嘗試單列與相鄰雙列表頭，回傳最佳候選。對照 JS guessHeader()。"""
    out = []
    n = min(len(rows), max_scan)
    for i in range(n):
        single = combine_header_rows(rows, [i])
        scored = score_header_row(single)
        out.append({'headerRows': [i], 'headers': single,
                     'originalHeaders': [rows[i] if i < len(rows) else []], **scored})
        if i + 1 < n:
            dbl = combine_header_rows(rows, [i, i + 1])
            scored2 = score_header_row(dbl)
            scored2['score'] += 5
            out.append({'headerRows': [i, i + 1], 'headers': dbl,
                         'originalHeaders': [rows[i] if i < len(rows) else [],
                                              rows[i + 1] if i + 1 < len(rows) else []], **scored2})
    out.sort(key=lambda x: -x['score'])
    if not out:
        return {'headerRows': [], 'score': 0, 'headers': [], 'detected': []}
    best = out[0]
    return {'headerRows': best['headerRows'], 'score': best['score'],
            'headers': best['headers'], 'detected': best.get('detected', [])}


def build_mapping(detected):
    """偵測結果轉欄位對應表（同一欄位只取第一個命中）。對照 JS buildMapping()。"""
    m = {}
    for x in detected:
        if x['field'] not in m:
            m[x['field']] = {'columnIndex': x['columnIndex'],
                              'confidence': 0.85 if x['matchType'] == 'exact' else 0.65}
    return m


def validate_total(qty, price, total):
    """複價驗算，容差與判定條件沿用與標單相同的慣例。

    v1.1 修正（Q-1）：原本只排除 qty/price/total 任一為 None 的情況，
    但 price=0 時 0×qty-0=0 一定會落在容差內，回報「OK」——這在報價單情境
    容易誤導，因為單價 0 通常代表「未報價／併入他項」，不代表複價真的驗算
    通過。boq_cleaner.py 對單價 0 的處理是 `_up in (None, 0)` 一律視為無法
    驗算，這裡補上同樣的判斷，才是文件宣稱的「與標單相同慣例」。
    """
    if qty is None or price is None or total is None:
        return '無法驗算'
    if price == 0:
        return '無法驗算'
    if abs(qty * price - total) <= abs(total) * 0.01 + 1:
        return 'OK'
    return '不符'


def extract_rows(rows, header_rows, mapping, file_name, sheet_name, vendor_name):
    """依表頭列與欄位對應，逐列擷取報價品項。對照 JS extractData()。"""
    data_start = max(header_rows) + 1
    combined_headers = combine_header_rows(rows, header_rows)

    def get(r, field):
        col = mapping.get(field, {}).get('columnIndex')
        if col is None or col >= len(r):
            return None
        return r[col]

    cleaned, logic_row = [], 0
    for idx, r in enumerate(rows[data_start:]):
        original_row = data_start + idx + 1  # 1-based，對齊 Excel 實際列號
        r = r or []
        if is_blank_row(r) or is_repeated_header(r, combined_headers) or is_noise_row(r):
            continue
        item = clean_text(get(r, 'item'))
        qty = to_number(get(r, 'qty'))
        price = to_number(get(r, 'price'))
        if not item and qty is None and price is None:
            continue  # 分類列／雜訊：靜默略過
        logic_row += 1
        vendor_cell = clean_text(get(r, 'vendor'))
        total = to_number(get(r, 'total'))
        cleaned.append({
            'vendor': vendor_cell or vendor_name,
            'vendor_from_column': bool(vendor_cell),
            'item': item,
            'spec': clean_text(get(r, 'spec')),
            'brand': clean_text(get(r, 'brand')),
            'model': clean_text(get(r, 'model')),
            'unit': clean_text(get(r, 'unit')),
            'qty': qty, 'price': price, 'total': total,
            'note': clean_text(get(r, 'note')),
            'qty_check': validate_total(qty, price, total),
            'file': file_name, 'sheet': sheet_name,
            'original_row': original_row, 'logic_row': logic_row,
        })
    return cleaned


def clean_quote_sheet(rows, file_name, sheet_name, vendor_name):
    """單一工作表的完整流程：表頭偵測 → 欄位對應 → 逐列擷取。

    回傳 (rows, status_dict)。批次（非互動）情境下，信心不足或缺必要欄位一律
    跳過並記錄原因，不自動臆測——與 boq_cleaner.py 找不到表頭時的慣例一致。
    """
    h = guess_header(rows)
    if not h['headerRows'] or h['score'] < 80 or 'item' not in {d['field'] for d in h['detected']}:
        return [], {'sheet': sheet_name, 'status': '略過（找不到可信表頭）',
                     'rows': len(rows), 'items': 0, 'header_row': None}

    mapping = build_mapping(h['detected'])
    missing = [f for f in REQUIRED_FIELDS if f not in mapping]
    if missing:
        return [], {'sheet': sheet_name,
                     'status': f'略過（缺必要欄位：{"、".join(missing)}）',
                     'rows': len(rows), 'items': 0,
                     'header_row': max(h['headerRows']) + 1}

    items = extract_rows(rows, h['headerRows'], mapping, file_name, sheet_name, vendor_name)
    return items, {'sheet': sheet_name, 'status': 'OK', 'rows': len(rows),
                    'items': len(items), 'header_row': max(h['headerRows']) + 1,
                    'columns': {k: v['columnIndex'] for k, v in mapping.items()}}


def clean_quote_workbook(path, sheet_names=None, skip_summary=True, vendor_name=None):
    """整份檔案的清洗入口，介面對照 boq_cleaner.clean_workbook()。

    vendor_name: 未指定時使用檔名（去副檔名）作為預設廠商名稱，供「廠商」欄
    缺漏時遞補；若工作表內有廠商欄位資料，仍以欄位內容優先。
    """
    import os
    default_vendor = vendor_name or os.path.splitext(os.path.basename(str(path)))[0]

    all_items, summary = [], []
    try:
        names = sheet_names or list_sheets(path)
    except Exception as e:
        summary.append({'sheet': str(path), 'status': f'讀取失敗：{e}',
                         'rows': 0, 'items': 0})
        return all_items, summary

    for sname in names:
        if skip_summary and ('總表' in sname or normalize_header(sname) == 'summary'):
            summary.append({'sheet': sname, 'status': '略過（總表）', 'rows': 0, 'items': 0})
            continue
        try:
            grid = read_grid(path, sname)
        except Exception as e:
            summary.append({'sheet': sname, 'status': f'讀取錯誤：{e}', 'rows': 0, 'items': 0})
            continue

        items, s = clean_quote_sheet(grid, os.path.basename(str(path)), sname, default_vendor)
        all_items.extend(items)
        summary.append(s)

    return all_items, summary


def items_to_rows(items):
    """轉成扁平 dict list，欄位順序固定，供 CSV/Excel 輸出使用。"""
    out = []
    for it in items:
        out.append({
            '檔案': it['file'], '工作表': it['sheet'],
            '原始列號': it['original_row'], '邏輯列': it['logic_row'],
            '廠商': it['vendor'], '品名': it['item'], '規格': it['spec'],
            '品牌': it['brand'], '型號': it['model'], '單位': it['unit'],
            '數量': it['qty'], '單價': it['price'], '複價': it['total'],
            '複價驗算': it['qty_check'], '備註': it['note'],
        })
    return out
