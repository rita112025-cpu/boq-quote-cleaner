"""
boq_cleaner.py — 工程標單清洗核心邏輯 v3

目標：把工程標單 Excel（含多層階層、父子拆列、規格說明文字列）整理成
「每列都是可比對品項」的乾淨表格。

v2 修正重點：
1. 跳過雙語/重複表頭列（例如 Particulars & Description / Unit / Quantity）。
2. 保留文字中的必要空白，不再把英文與規格全部黏在一起。
3. 支援前置小組名繼承：A / 1 / (01) 這類無單位數量列，若位於計價品項前，會成為下一批品項的語意前綴。
4. 支援後置規格說明：- / ‧ / 空白項次等說明列，會掛回上一筆計價品項的備註。
5. 補列 / 增列 這類無單位數量列改成前置語意，不再誤掛到上一筆品項。
6. v3 新增 context_open：區分「前置情境延伸」與「後置規格備註」，避免空白項次描述行掛錯上一筆計價品項。
7. v3.1 修正雙數量欄：部分標單有「數量 / 工地數量」兩欄，複價實際用工地數量計算。
   舊版取第一個非空 → 抓到設計數量，導致大量「數量×單價 ≠ 複價」。
   改用 _choose_qty 以複價一致性挑欄，並輸出「另一數量欄」與「複價驗算」欄。
"""
from __future__ import annotations

import re

# v1.1(F-3)：read_grid/list_sheets 已抽出至 excel_io.py（純 I/O，無清洗邏輯），
# 避免報價單批次包為了重用這兩個函式而必須夾帶整份 boq_cleaner.py。
from excel_io import read_grid, list_sheets
from dataclasses import dataclass, field
from typing import Optional, Any

CATEGORY_FORMAL = set('壹貳參肆伍陸柒捌玖拾')
CATEGORY_CASUAL = set('一二三四五六七八九十')
SUBTOTAL_RE = re.compile(r'(小\s*計|合\s*計|總\s*計)')
BRACKET_RE = re.compile(r'^[<〈].+[>〉]$')
LETTER_RE = re.compile(r'^[A-Za-z]$')
BULLET_RE = re.compile(r'^[-－–—‧•●◆◇◎※＊*]')
NUMERIC_GROUP_RE = re.compile(r'^(\d+(?:\.\d+)?|\(\d+\)|（\d+）|【\d+】)$')
EXPLICIT_CONTEXT_RE = re.compile(r'^(補列|增列|新增|追加)$')
# 部分業主標單風格：(一)(二) 中文括號小組；三.1 三.2 章節編號
CHINESE_BRACKET_RE = re.compile(r'^[\(（][一二三四五六七八九十壹貳參肆伍陸柒捌玖拾百千]+[\)）]$')
CHINESE_DECIMAL_RE = re.compile(r'^[一二三四五六七八九十壹貳參肆伍陸柒捌玖拾]+[\.．]\d+')
HEADER_WORD_RE = re.compile(
    r'(particulars|description|unit\s*price|quantity|\bunit\b|\btotal\b|項次|工程項目|項目及說明|品名|名稱|單位|數量|單價|複價)',
    re.I,
)

NAME_KEYS = ['品名', '名稱', '項目名稱', '工程項目及說明', '項次說明', '工作項目', '施工項目', 'Particulars', '項目']
UNIT_KEYS = ['單位', 'unit']
QTY_KEYS = ['數量', 'qty', 'quantity', '用量']
ITEMNO_KEYS = ['項次', '項目', 'item']
PRICE_KEYS = ['複價', 'total', '總價', '金額']
UNIT_PRICE_KEYS = ['單價', 'unit price']


@dataclass
class CleanItem:
    row_no: int
    breadcrumb: str
    name: str
    unit: str
    qty: Any
    unit_price: Any
    total_price: Any
    spec_notes: list[str] = field(default_factory=list)
    source_context: Optional[str] = None
    sheet: str = ''
    qty_alt: str = ''          # v3.1：另一個數量欄的值（例如設計數量 vs 工地數量）
    qty_check: str = ''        # v3.1：複價驗算結果 OK / 不符 / 無法驗算


def _text(v: Any) -> str:
    """輸出用清理：保留字詞間空白，只移除多餘空白。"""
    if v is None:
        return ''
    s = str(v).replace('\u3000', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    # Excel 數字 1.0 作為項次時顯示成 1
    if re.fullmatch(r'\d+\.0', s):
        return s[:-2]
    return s


def _compact(v: Any) -> str:
    """判斷用清理：移除全部空白並小寫。"""
    return re.sub(r'\s+', '', _text(v)).lower()


def _is_empty(v: Any) -> bool:
    return v is None or _text(v) == ''


def _is_numeric_like(v: Any) -> bool:
    if v is None or _text(v) == '':
        return False
    try:
        float(str(v).replace(',', ''))
        return True
    except Exception:
        return False


def _num(v: Any) -> Optional[float]:
    """轉數值；空字串、文字、NaN 都回 None。"""
    if _is_empty(v):
        return None
    try:
        import math
        f = float(str(v).replace(',', ''))
        return None if math.isnan(f) else f
    except Exception:
        return None


def _choose_qty(row, qty_cols, unit_price, total_price, get):
    """v3.1：挑正確的數量欄。

    部分標單有「數量 / 工地數量」兩欄，複價實際是用『工地數量』算的。舊版取第一個
    非空值 → 會抓到設計數量那欄，導致「數量 × 單價 ≠ 複價」。

    修正：有單價與複價時，選『數量 × 單價 最接近複價』的欄；否則退回第一個非空
    值（維持單一數量欄工作表的舊行為）。回傳 (chosen_value, alt_values)。
    """
    candidates = []
    for qc in qty_cols:
        v = get(row, qc)
        if not _is_empty(v):
            candidates.append(v)
    if not candidates:
        return None, []
    up, tp = _num(unit_price), _num(total_price)
    chosen = candidates[0]
    if up not in (None, 0) and tp is not None:
        best, best_err = None, None
        for v in candidates:
            q = _num(v)
            if q is None:
                continue
            err = abs(q * up - tp)
            if best_err is None or err < best_err:
                best, best_err = v, err
        if best is not None:
            chosen = best
    alt = [_text(v) for v in candidates if v is not chosen and _text(v) != _text(chosen)]
    return chosen, alt


def find_header_row(grid, max_scan=30):
    """掃前 N 列，用關鍵字評分找表頭列。回傳 (header_row_idx, col_map) 或 (None, None)。"""
    best_idx, best_score, best_cols = None, 0, None
    for i in range(min(max_scan, len(grid))):
        cells = [_compact(c) for c in grid[i]]
        cols = {}
        score = 0
        for key_list, tag, weight in [
            (NAME_KEYS, 'name', 2), (UNIT_KEYS, 'unit', 1),
            (QTY_KEYS, 'qty', 1), (ITEMNO_KEYS, 'itemno', 1),
            (PRICE_KEYS, 'price', 1), (UNIT_PRICE_KEYS, 'unit_price', 1),
        ]:
            for ci, c in enumerate(cells):
                if c and any(k.replace(' ', '').lower() in c for k in key_list) and tag not in cols:
                    cols[tag] = ci
                    score += weight
                    break
        # 至少 2 個不同欄位被識別，否則可能是含多個關鍵字的長描述列（假陽性）
        if score > best_score and len(set(cols.values())) >= 2:
            best_idx, best_score, best_cols = i, score, cols
    if best_score < 3:
        return None, None
    # 補充所有可能的數量欄。部分標單有「數量 / 工地數量」兩欄，部分補列只填工地數量。
    # 若只抓第一個「數量」欄，會漏掉這類真實計價品項。
    header_cells = [_compact(c) for c in grid[best_idx]]
    qty_cols = []
    for ci, c in enumerate(header_cells):
        if c and any(k.replace(' ', '').lower() in c for k in QTY_KEYS) and ci not in qty_cols:
            qty_cols.append(ci)
    if qty_cols:
        best_cols['qty_cols'] = qty_cols
    # name/itemno 衝突解決：若兩者指向同一欄（例如 header 只有 '項目' 一詞），
    # 以其他 NAME_KEYS 重新掃描，找出真正的品名欄。
    if best_cols.get('name') is not None and best_cols.get('name') == best_cols.get('itemno'):
        alt_name_keys = [k for k in NAME_KEYS if k != '項目']
        for ci, c in enumerate(header_cells):
            if ci == best_cols['itemno']:
                continue
            if c and any(k.replace(' ', '').lower() in c for k in alt_name_keys):
                best_cols['name'] = ci
                break
    return best_idx, best_cols


def classify_itemno(item_no: Any) -> str:
    s = _compact(item_no)
    shown = _text(item_no)
    if not s:
        return 'DETAIL'
    if len(shown) == 1 and shown in CATEGORY_FORMAL:
        return 'CATEGORY_1'
    if len(shown) == 1 and shown in CATEGORY_CASUAL:
        return 'CATEGORY_2'
    if BRACKET_RE.match(shown):
        return 'CATEGORY_2'
    if CHINESE_BRACKET_RE.match(shown):   # (一)(二)（三）等
        return 'CONTEXT_MAJOR'
    if LETTER_RE.match(shown):
        return 'CONTEXT_MAJOR'
    if EXPLICIT_CONTEXT_RE.match(shown):
        return 'CONTEXT_MINOR'
    if NUMERIC_GROUP_RE.match(shown):
        return 'CONTEXT_MINOR'
    if CHINESE_DECIMAL_RE.match(shown):   # 三.1 三.2 等章節編號
        return 'CATEGORY_2'
    return 'DETAIL'


def is_header_like(item_no: Any, name: Any, unit: Any, qty: Any, price: Any = None) -> bool:
    joined = ' '.join(_text(x) for x in [item_no, name, unit, qty, price] if not _is_empty(x))
    c = _compact(joined)
    if not c:
        return False
    # 雙語表頭：Particulars & Description / Unit / Quantity / Total
    if 'particulars' in c and 'description' in c:
        return True
    # 中文重複表頭
    if '工程項目及說明' in c and '單位' in c and '數量' in c:
        return True
    # 欄位值本身就是 Unit / Quantity / Total
    if _compact(unit) in {'unit', '單位'} or _compact(qty) in {'quantity', 'qty', '數量'}:
        return True
    # v3.2 修正（B-1）：原規則只要 name 欄含任一表頭關鍵字（如「數量」「名稱」
    # 「單位」）且沒有真實數量，就整列判為表頭，導致「數量依現場丈量為準」
    # 「設備名稱牌製作」這類真實說明文字被誤判丟棄，遺失規格備註。
    # 改為要求 item_no／name／unit／price 這幾欄「各自獨立」至少兩欄命中表頭
    # 關鍵字，才視為表頭列——真正的重複表頭通常多欄同時是欄名，一般說明句
    # 只會佔用單一欄位（name），可用欄數區分。
    field_hits = sum(
        1 for cell in (item_no, name, unit, price)
        if not _is_empty(cell) and HEADER_WORD_RE.search(_text(cell))
    )
    if field_hits >= 2 and not _is_numeric_like(qty):
        return True
    return False


def _join_parts(parts: list[str]) -> str:
    seen = []
    for p in parts:
        p = _text(p)
        if p and p not in seen:
            seen.append(p)
    return ' '.join(seen)


def clean_boq(grid, header_idx, col_map):
    name_col = col_map.get('name')
    unit_col = col_map.get('unit')
    qty_col = col_map.get('qty')
    qty_cols = col_map.get('qty_cols') or ([qty_col] if qty_col is not None else [])
    itemno_col = col_map.get('itemno')
    price_col = col_map.get('price')
    unit_price_col = col_map.get('unit_price')

    items: list[CleanItem] = []
    row_audit: list[dict] = []
    unclassified: list[tuple] = []
    breadcrumb = ['', '']
    context_parts: list[str] = []
    pending_notes: list[str] = []
    # v3：情境鏈是否仍在等待第一批計價品項消耗。
    # True：空白項次描述列應視為 CONTEXT_EXTEND，併入下一批品項情境。
    # False：空白項次描述列才視為上一筆計價品項的後置規格備註。
    context_open = False
    # v3.2 修正（B-2）：CATEGORY_1/CATEGORY_2 切換後、尚未產生本分類第一筆
    # 計價品項前，True 代表「跨過分類邊界」。此時若出現無數量的規格列，
    # 不可掛回上一分類最後一筆品項（那是錯誤歸屬），應暫存等待本分類
    # 第一筆品項出現後再掛上去。
    category_boundary = False

    def get(row, idx):
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    def audit(row_no, row_type, name, item_no='', reason=''):
        row_audit.append({
            'row_no': row_no,
            'row_type': row_type,
            'item_no': _text(item_no),
            'name': _text(name),
            'breadcrumb': _join_parts([b for b in breadcrumb if b]),
            'context': _join_parts(context_parts),
            'reason': reason,
        })

    for row_idx in range(header_idx + 1, len(grid)):
        row = grid[row_idx]
        row_no = row_idx + 1
        name = _text(get(row, name_col))
        item_no = get(row, itemno_col)
        if not name:
            audit(row_no, 'BLANK', name, item_no)
            continue

        unit_raw = get(row, unit_col)
        price_raw = get(row, price_col)
        unit_price_raw = get(row, unit_price_col)
        # v3.1：先讀價格，再用「數量×單價 最接近複價」挑正確數量欄（見 _choose_qty）。
        qty_raw, qty_alt = _choose_qty(row, qty_cols, unit_price_raw, price_raw, get)
        unit = _text(unit_raw)
        has_unit = bool(unit)
        has_qty = not _is_empty(qty_raw)

        if is_header_like(item_no, name, unit_raw, qty_raw, price_raw):
            audit(row_no, 'HEADER', name, item_no, 'repeated/bilingual header row')
            continue
        if SUBTOTAL_RE.search(name):
            audit(row_no, 'SUBTOTAL', name, item_no)
            continue

        if (has_unit or unit_col is None) and has_qty:
            full_context = _join_parts(context_parts)
            merged_name = _join_parts([full_context, name])
            item = CleanItem(
                row_no=row_no,
                breadcrumb=_join_parts([b for b in breadcrumb if b]),
                name=merged_name,
                unit=unit,
                qty=qty_raw,
                unit_price=unit_price_raw,
                total_price=price_raw,
                source_context=full_context or None,
                qty_alt='、'.join(qty_alt) if qty_alt else '',
            )
            # v3.1：複價驗算。q*單價 是否等於複價（1% 容差）。
            _q, _up, _tp = _num(qty_raw), _num(unit_price_raw), _num(price_raw)
            if _up in (None, 0) or _tp is None or _q is None:
                item.qty_check = '無法驗算'
            elif abs(_q * _up - _tp) <= abs(_tp) * 0.01 + 1:
                item.qty_check = 'OK'
            else:
                item.qty_check = '不符'
            if pending_notes:
                item.spec_notes.extend(pending_notes)
                pending_notes.clear()
            items.append(item)
            context_open = False
            category_boundary = False
            audit(row_no, 'PRICE_ITEM', name, item_no, 'has unit + qty')
            continue

        # 有單位但沒有數量、或有金額欄但數量空白：通常不是下一批品項的小組名，
        # 而是上一個計價品項的附屬規格 / 零價子項。
        # 不可讓它變成持續繼承 context，否則會污染後續品項。
        # v3.2（B-2）：剛跨過分類邊界（category_boundary=True）時，即使 items
        # 非空，也不可掛回上一分類最後一筆品項——那屬於錯誤歸屬，應暫存到
        # pending_notes，等本分類第一筆品項出現後再掛上去。
        if (has_unit or not _is_empty(price_raw) or not _is_empty(unit_price_raw)) and not has_qty:
            if items and not category_boundary:
                items[-1].spec_notes.append(name)
                audit(row_no, 'SPEC_NOTE_AFTER_ITEM', name, item_no, 'has unit/price but missing qty')
            else:
                pending_notes.append(name)
                audit(row_no, 'SPEC_NOTE_BEFORE_ITEM', name, item_no,
                      'has unit/price but missing qty' + (' (category boundary)' if category_boundary else ''))
            continue

        kind = classify_itemno(item_no)
        if kind == 'CATEGORY_1':
            breadcrumb[0] = name
            breadcrumb[1] = ''
            context_parts.clear()
            pending_notes.clear()
            context_open = False
            category_boundary = True
            audit(row_no, 'CATEGORY_1', name, item_no)
        elif kind == 'CATEGORY_2':
            breadcrumb[1] = name
            context_parts.clear()
            pending_notes.clear()
            context_open = False
            category_boundary = True
            audit(row_no, 'CATEGORY_2', name, item_no)
        elif kind == 'CONTEXT_MAJOR':
            context_parts[:] = [name]
            pending_notes.clear()
            context_open = True
            audit(row_no, 'CONTEXT_MAJOR', name, item_no)
        elif kind == 'CONTEXT_MINOR':
            # 數字/補列/增列但沒有單位數量：通常是下一批品項的小組名。
            # 若是明顯補列/增列，會替換舊 context；否則掛在現有 major context 後面。
            shown = _text(item_no)
            if EXPLICIT_CONTEXT_RE.match(shown):
                context_parts[:] = [name]
            else:
                if context_parts:
                    # 同層小組切換：保留第一層 major context，替換第二層小組名。
                    context_parts[:] = context_parts[:1] + [name]
                else:
                    context_parts[:] = [name]
            pending_notes.clear()
            context_open = True
            audit(row_no, 'CONTEXT_MINOR', name, item_no)
        else:
            # v3：空白項次描述列有兩種完全不同的語意：
            # 1) 前置情境尚未被計價品項消耗：它是情境延伸，應併入 context_parts。
            # 2) 前置情境已被計價品項消耗：它才是上一筆計價品項的後置規格備註。
            if _compact(item_no) == '' and context_open and context_parts:
                context_parts.append(name)
                audit(row_no, 'CONTEXT_EXTEND', name, item_no, 'blank item_no while context is open')
            elif items and not category_boundary and (BULLET_RE.match(name) or not context_parts or _compact(item_no) == '' or not context_open):
                items[-1].spec_notes.append(name)
                audit(row_no, 'SPEC_NOTE_AFTER_ITEM', name, item_no)
            else:
                # 第一個計價品項前的補充說明，先暫存，掛到下一個計價品項。
                pending_notes.append(name)
                audit(row_no, 'SPEC_NOTE_BEFORE_ITEM', name, item_no)
                if not items:
                    unclassified.append((row_no, name))

    return items, unclassified, row_audit


def clean_workbook(path, sheet_names=None, skip_summary=True):
    all_items: list[CleanItem] = []
    all_audit: list[dict] = []
    summary = []
    try:
        names = sheet_names or list_sheets(path)
    except Exception as e:
        summary.append({'sheet': str(path), 'status': f'讀取失敗：{e}', 'rows': 0, 'items': 0, 'unclassified': 0})
        return all_items, summary, all_audit
    for sname in names:
        if skip_summary and ('總表' in sname or _compact(sname) == 'summary'):
            summary.append({'sheet': sname, 'status': '略過（總表）', 'rows': 0, 'items': 0, 'unclassified': 0})
            continue
        try:
            grid = read_grid(path, sname)
        except Exception as e:
            summary.append({'sheet': sname, 'status': f'讀取錯誤：{e}', 'rows': 0, 'items': 0, 'unclassified': 0})
            continue
        header_idx, col_map = find_header_row(grid)
        if header_idx is None:
            summary.append({'sheet': sname, 'status': '略過（找不到可信表頭）', 'rows': len(grid), 'items': 0, 'unclassified': 0})
            continue
        items, unclassified, row_audit = clean_boq(grid, header_idx, col_map)
        for it in items:
            it.sheet = sname
        for a in row_audit:
            a['sheet'] = sname
        all_items.extend(items)
        all_audit.extend(row_audit)
        summary.append({
            'sheet': sname,
            'status': 'OK',
            'rows': len(grid),
            'items': len(items),
            'unclassified': len(unclassified),
            'header_row': header_idx + 1,
            'columns': col_map,
        })
    return all_items, summary, all_audit


def items_to_rows(items):
    out = []
    for it in items:
        out.append({
            '工作表': getattr(it, 'sheet', ''),
            '原始列號': it.row_no,
            '大分類路徑': it.breadcrumb,
            '品名（含繼承語意）': it.name,
            '單位': it.unit,
            '數量': it.qty,
            '另一數量欄': getattr(it, 'qty_alt', ''),
            '複價驗算': getattr(it, 'qty_check', ''),
            '單價': it.unit_price,
            '複價': it.total_price,
            '繼承來源（供核對）': it.source_context or '',
            '規格備註': '；'.join(it.spec_notes) if it.spec_notes else '',
        })
    return out
