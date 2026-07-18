"""
test_quote_cleaner.py — quote_cleaner.py 單元測試

驗證範圍：Python 移植版與原始 JS 版（合併版 HTML「報價單處理」模組）
行為一致。這是「移植正確性」的驗證，不是「真實報價單準確度」的驗證——
後者需要真實樣本累積，見 quote_cleaner.py 檔頭說明。
"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from quote_cleaner import (
    normalize_header, clean_text, to_number, classify_cell,
    is_blank_row, is_noise_row, combine_header_rows, is_repeated_header,
    detect_header_field, score_header_row, guess_header, build_mapping,
    validate_total, extract_rows, clean_quote_sheet, clean_quote_workbook,
    items_to_rows, HEADER_RULES, REQUIRED_FIELDS,
)


class TestUtils(unittest.TestCase):
    """對照 JS normalizeHeader/cleanText/toNumber/classifyCell 的既有行為。"""

    def test_normalize_header(self):
        self.assertEqual(normalize_header(None), '')
        self.assertEqual(normalize_header('  品名  '), '品名')
        self.assertEqual(normalize_header('Unit Price'), 'unitprice')
        self.assertEqual(normalize_header('單位：'), '單位')
        self.assertEqual(normalize_header('（備註）'), '備註')
        self.assertEqual(normalize_header('數　量'), '數量')  # 全形空格（NFKC）

    def test_clean_text(self):
        self.assertEqual(clean_text(None), '')
        self.assertEqual(clean_text('  PVC  管  '), 'PVC 管')
        self.assertEqual(clean_text('第一行\n第二行'), '第一行 第二行')
        self.assertEqual(clean_text(100), '100')
        self.assertEqual(clean_text(1.0), '1.0')  # 與 JS String(1.0)="1" 不同，
        # 因為 Python float(1.0) 轉字串是 "1.0"；此差異只影響數字型儲存格被誤
        # 當文字欄位讀取的邊界情況，不影響 to_number() 的數值解析。

    def test_to_number(self):
        self.assertIsNone(to_number(None))
        self.assertIsNone(to_number(''))
        self.assertEqual(to_number(100), 100.0)
        self.assertEqual(to_number(3.14), 3.14)
        self.assertEqual(to_number('1,234'), 1234.0)
        self.assertEqual(to_number('1,234.5'), 1234.5)
        self.assertIsNone(to_number('abc'))
        self.assertIsNone(to_number(float('nan')))
        self.assertIsNone(to_number(float('inf')))
        self.assertEqual(to_number('$100'), 100.0)
        self.assertEqual(to_number('8500元'), 8500.0)  # 結尾「元」：正常後綴用法

    def test_to_number_currency_middle_rejected(self):
        # v1.1 修正（Q-2）：「元」出現在數字中間不再被靜默移除，避免
        # 「3元5」這類黏連文字被誤解析成 35（靜默產生錯誤數值）。
        self.assertIsNone(to_number('3元5'))
        self.assertIsNone(to_number('元100'))  # 前綴「元」不是常見寫法，不支援

    def test_to_number_whitespace_as_thousand_sep(self):
        # 空白仍視為千分位分隔符移除（沿用既有行為，非本次修正範圍）
        self.assertEqual(to_number('1 000'), 1000.0)

    def test_classify_cell(self):
        self.assertEqual(classify_cell(None), 'empty')
        self.assertEqual(classify_cell(''), 'empty')
        self.assertEqual(classify_cell(100), 'number')
        self.assertEqual(classify_cell('1,234'), 'number')
        self.assertEqual(classify_cell('abc'), 'text')
        self.assertEqual(classify_cell('品名'), 'text')

    def test_is_blank_row(self):
        self.assertTrue(is_blank_row([]))
        self.assertTrue(is_blank_row(None))
        self.assertTrue(is_blank_row([None, '', '  ']))
        self.assertFalse(is_blank_row([None, 'x', '']))

    def test_is_noise_row(self):
        self.assertTrue(is_noise_row(['小計', None, 100]))
        self.assertTrue(is_noise_row(['合計金額', 5000]))
        self.assertTrue(is_noise_row(['第3頁', None]))
        self.assertTrue(is_noise_row(['承攬廠商簽章欄', None]))
        self.assertFalse(is_noise_row(['PVC管', '2吋', 100]))


class TestHeaderDetection(unittest.TestCase):
    """對照 JS detectHeaderField/scoreHeaderRow/guessHeader/buildMapping。"""

    def test_detect_header_field_exact(self):
        r = detect_header_field('品名')
        self.assertEqual(r['field'], 'item')
        self.assertEqual(r['matchType'], 'exact')
        self.assertEqual(r['score'], 30)

    def test_detect_header_field_partial(self):
        r = detect_header_field('單價(未稅)')
        self.assertIsNotNone(r)
        self.assertEqual(r['field'], 'price')
        self.assertEqual(r['matchType'], 'partial')

    def test_detect_header_field_none(self):
        # 純敘述性文字，且不含任何欄位別名的子字串，才應該回 None。
        # 注意：'工程名稱' 這類字串會因含「名稱」（item 別名之一）被 partial 命中，
        # 這是原始 JS 演算法既有的 substring 比對行為，不是本測試要驗證的對象。
        self.assertIsNone(detect_header_field('地址電話傳真'))
        self.assertIsNone(detect_header_field(None))
        self.assertIsNone(detect_header_field(''))

    def test_partial_match_can_overreach_on_substrings(self):
        # 記錄一個已知、繼承自 JS 原始邏輯的寬鬆特性：只要文字包含欄位別名的
        # 子字串就會 partial 命中，因此「工程名稱」會被誤判為 item 相關欄位。
        # 這不是 Python 移植造成的差異，是移植前就存在的行為，先記錄下來，
        # 之後有真實報價單案例顯示這會造成誤判時，再回頭調整比對規則。
        r = detect_header_field('工程名稱')
        self.assertIsNotNone(r)
        self.assertEqual(r['field'], 'item')
        self.assertEqual(r['matchType'], 'partial')

    def test_score_header_row_basic(self):
        row = ['廠商', '品名', '規格', '單位', '數量', '單價', '複價']
        s = score_header_row(row)
        self.assertGreater(s['score'], 80)
        self.assertIn('item', s['fields'])
        self.assertIn('vendor', s['fields'])

    def test_score_header_row_data_row_scores_low(self):
        # 一般資料列不該被誤判為表頭
        row = ['範例廠商A', 'PVC管', '2吋', '支', 100, 85, 8500]
        s = score_header_row(row)
        self.assertLess(s['score'], 80)

    def test_score_header_row_title_penalty(self):
        row = ['工程名稱：某某大樓機電工程', None, None, None]
        s = score_header_row(row)
        self.assertLess(s['score'], 0)

    def test_combine_header_rows_dedup(self):
        rows = [['單價', '複價'], ['未稅', '未稅']]
        out = combine_header_rows(rows, [0, 1])
        self.assertEqual(out, ['單價 未稅', '複價 未稅'])

    def test_combine_header_rows_no_dup_when_same(self):
        rows = [['品名'], ['品名']]  # 兩列相同，不應重複串接
        out = combine_header_rows(rows, [0, 1])
        self.assertEqual(out, ['品名'])

    def test_guess_header_picks_best_row(self):
        rows = [
            ['報價單', None, None, None],
            ['工程名稱：測試案', None, None, None],
            ['廠商', '品名', '單位', '數量', '單價', '複價'],
            ['範例廠商A', 'PVC管', '支', 100, 85, 8500],
        ]
        h = guess_header(rows)
        self.assertEqual(h['headerRows'], [2])
        fields = {d['field'] for d in h['detected']}
        self.assertIn('item', fields)
        self.assertIn('vendor', fields)

    def test_guess_header_double_row(self):
        # 單列各自分數都不足以成為表頭（row0 只有 2 個精確命中、欄數<3 被扣分；
        # row1 有廠商/品名但「未稅」不命中任何欄位），合併後才湊出足夠分數。
        rows = [
            [None, None, '單價', '複價'],
            ['廠商', '品名', '未稅', '未稅'],
            ['範例廠商A', 'PVC管', 85, 8500],
        ]
        h = guess_header(rows)
        self.assertEqual(h['headerRows'], [0, 1])

    def test_build_mapping_first_hit_wins(self):
        detected = [
            {'field': 'item', 'columnIndex': 0, 'matchType': 'exact'},
            {'field': 'item', 'columnIndex': 5, 'matchType': 'partial'},
        ]
        m = build_mapping(detected)
        self.assertEqual(m['item']['columnIndex'], 0)

    def test_is_repeated_header(self):
        headers = ['廠商', '品名', '單位', '數量', '單價', '複價']
        repeated_row = ['廠商', '品名', '單位', '數量', '單價', '複價']
        self.assertTrue(is_repeated_header(repeated_row, headers))
        data_row = ['範例廠商A', 'PVC管', '支', 100, 85, 8500]
        self.assertFalse(is_repeated_header(data_row, headers))


class TestValidateTotal(unittest.TestCase):
    def test_ok(self):
        self.assertEqual(validate_total(100, 85, 8500), 'OK')

    def test_mismatch(self):
        self.assertEqual(validate_total(5, 50, 999), '不符')

    def test_within_tolerance(self):
        # 100 × 85 = 8500，容差 |8500|*1%+1 = 86，8550 在容差內
        self.assertEqual(validate_total(100, 85, 8550), 'OK')

    def test_cannot_validate(self):
        self.assertEqual(validate_total(None, 85, 8500), '無法驗算')
        self.assertEqual(validate_total(100, None, 8500), '無法驗算')
        self.assertEqual(validate_total(100, 85, None), '無法驗算')

    def test_price_zero_cannot_validate(self):
        # v1.1 修正（Q-1）：單價 0 一律視為無法驗算，與 boq_cleaner.py
        # 的 `_up in (None, 0)` 判斷對齊。單價 0 通常代表「未報價／併入
        # 他項」，不該因為 0×qty-0=0 落在容差內就標記為「OK」。
        self.assertEqual(validate_total(5, 0, 0), '無法驗算')
        self.assertEqual(validate_total(100, 0, 500), '無法驗算')


class TestExtractRows(unittest.TestCase):
    """對照 JS extractData()：整合測試，模擬一份小型合成報價單。"""

    def setUp(self):
        self.rows = [
            ['廠商', '品名', '規格', '單位', '數量', '單價', '複價'],           # row0 表頭
            ['範例廠商A', 'PVC管', '2吋', '支', 100, 85, 8500],                  # row1 OK
            ['範例廠商A', '不鏽鋼閘閥', '3吋', '只', 20, 2350, 47000],          # row2 OK
            [None, None, None, None, None, None, None],                        # row3 空白
            ['範例廠商A', '銅球閥', '3/4吋', '只', 50, 0, 0],                    # row4 單價0
            ['廠商', '品名', '規格', '單位', '數量', '單價', '複價'],           # row5 重複表頭
            [None, '小計', None, None, None, None, 55500],                     # row6 雜訊
            ['範例廠商A', '電線', '', '米', 500, 18, 9100],                      # row7 複價不符
        ]

    def test_basic_extraction(self):
        h = guess_header(self.rows)
        mapping = build_mapping(h['detected'])
        items = extract_rows(self.rows, h['headerRows'], mapping,
                              'test.xlsx', 'Sheet1', 'default_vendor')
        # 應該有 4 筆（row1,2,4,7），排除空白/重複表頭/小計
        self.assertEqual(len(items), 4)

    def test_original_row_traceability(self):
        h = guess_header(self.rows)
        mapping = build_mapping(h['detected'])
        items = extract_rows(self.rows, h['headerRows'], mapping,
                              'test.xlsx', 'Sheet1', 'default_vendor')
        rows_no = [it['original_row'] for it in items]
        self.assertEqual(rows_no, [2, 3, 5, 8])  # 1-based

    def test_qty_check_values(self):
        h = guess_header(self.rows)
        mapping = build_mapping(h['detected'])
        items = extract_rows(self.rows, h['headerRows'], mapping,
                              'test.xlsx', 'Sheet1', 'default_vendor')
        checks = {it['item']: it['qty_check'] for it in items}
        self.assertEqual(checks['PVC管'], 'OK')
        self.assertEqual(checks['不鏽鋼閘閥'], 'OK')
        self.assertEqual(checks['電線'], '不符')  # 500*18=9000 ≠ 9100

    def test_vendor_from_column(self):
        h = guess_header(self.rows)
        mapping = build_mapping(h['detected'])
        items = extract_rows(self.rows, h['headerRows'], mapping,
                              'test.xlsx', 'Sheet1', 'default_vendor')
        self.assertTrue(all(it['vendor'] == '範例廠商A' for it in items))

    def test_vendor_fallback_to_filename(self):
        # 沒有廠商欄位時，應遞補檔名
        rows = [
            ['品名', '單位', '數量', '單價', '複價'],
            ['PVC管', '支', 100, 85, 8500],
        ]
        h = guess_header(rows)
        mapping = build_mapping(h['detected'])
        items = extract_rows(rows, h['headerRows'], mapping,
                              '廠商甲.xlsx', 'Sheet1', '廠商甲')
        self.assertEqual(items[0]['vendor'], '廠商甲')
        self.assertFalse(items[0]['vendor_from_column'])


class TestCleanQuoteSheet(unittest.TestCase):
    def test_low_confidence_skipped(self):
        rows = [['foo', 'bar'], ['baz', 'qux']]
        items, summary = clean_quote_sheet(rows, 'x.xlsx', 'Sheet1', 'x')
        self.assertEqual(items, [])
        self.assertIn('找不到可信表頭', summary['status'])

    def test_missing_required_field_skipped(self):
        # 表頭需先通過信心門檻（score>=80 且偵測到 item），才會進入「缺必要
        # 欄位」判斷；這裡刻意組出高分表頭但不含「單價」，觸發該分支。
        rows = [
            ['廠商', '品名', '規格', '品牌', '型號', '備註', '單位', '數量'],
            ['範例廠商A', 'PVC管', '2吋', '南亞', 'X1', '', '支', 100],
        ]
        items, summary = clean_quote_sheet(rows, 'x.xlsx', 'Sheet1', 'x')
        self.assertEqual(items, [])
        self.assertIn('缺必要欄位', summary['status'])
        self.assertIn('price', summary['status'])

    def test_ok_sheet(self):
        rows = [
            ['廠商', '品名', '單位', '數量', '單價', '複價'],
            ['範例廠商A', 'PVC管', '支', 100, 85, 8500],
        ]
        items, summary = clean_quote_sheet(rows, 'x.xlsx', 'Sheet1', 'x')
        self.assertEqual(summary['status'], 'OK')
        self.assertEqual(len(items), 1)
        self.assertEqual(summary['header_row'], 1)


class TestItemsToRows(unittest.TestCase):
    def test_field_order_and_keys(self):
        rows = [
            ['廠商', '品名', '單位', '數量', '單價', '複價'],
            ['範例廠商A', 'PVC管', '支', 100, 85, 8500],
        ]
        items, _ = clean_quote_sheet(rows, 'x.xlsx', 'Sheet1', 'x')
        out = items_to_rows(items)
        self.assertEqual(out[0]['品名'], 'PVC管')
        self.assertEqual(out[0]['廠商'], '範例廠商A')
        self.assertEqual(out[0]['複價驗算'], 'OK')
        expected_keys = {'檔案', '工作表', '原始列號', '邏輯列', '廠商', '品名',
                          '規格', '品牌', '型號', '單位', '數量', '單價', '複價',
                          '複價驗算', '備註'}
        self.assertEqual(set(out[0].keys()), expected_keys)


class TestNoRealDependencyOnBoqCore(unittest.TestCase):
    """v1.1(F-3) 更新：quote_cleaner.py 現在完全不 import boq_cleaner.py
    （改從 excel_io.py 取得純 I/O 讀檔函式），兩套工具徹底獨立，不再共用
    同一份 474 行的標單核心檔案，避免版本漂移風險。"""

    def test_no_boq_cleaner_import(self):
        import quote_cleaner
        with open(quote_cleaner.__file__, encoding='utf-8') as f:
            src = f.read()
        self.assertNotIn('from boq_cleaner', src,
                          'quote_cleaner.py 不應再 import boq_cleaner.py（改用 excel_io.py）')
        self.assertNotIn('import boq_cleaner', src)

    def test_uses_excel_io_for_io(self):
        import quote_cleaner
        with open(quote_cleaner.__file__, encoding='utf-8') as f:
            src = f.read()
        self.assertIn('from excel_io import', src)

    def test_no_boq_cleaning_functions_imported(self):
        import quote_cleaner
        with open(quote_cleaner.__file__, encoding='utf-8') as f:
            src = f.read()
        forbidden = ['classify_itemno', 'clean_boq', '_choose_qty', 'CleanItem']
        for name in forbidden:
            self.assertNotIn(name, src, f'quote_cleaner.py 不應引用標單核心的 {name}')


if __name__ == '__main__':
    unittest.main(verbosity=2)
