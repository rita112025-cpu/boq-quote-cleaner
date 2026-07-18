"""
工程報價單批次清洗 CLI — v1.0

用法：
    # 明確列出檔案
    python export_cleaned_quotes.py A.xlsx B.xls --outdir output

    # 指向資料夾，批次處理裡面全部 Excel（預設遞迴子資料夾）
    python export_cleaned_quotes.py --folder ./報價單資料夾 --outdir output

    # 資料夾模式但不遞迴子資料夾
    python export_cleaned_quotes.py --folder ./報價單資料夾 --no-recursive --outdir output

    # 檔案與資料夾可併用
    python export_cleaned_quotes.py extra.xlsx --folder ./報價單資料夾 --outdir output

    # 指定廠商名稱（單一檔案時較適用；批次時預設用各檔案的檔名當廠商名）
    python export_cleaned_quotes.py 範例廠商A報價.xlsx --vendor 範例廠商A --outdir output

與 export_cleaned.py（標單批次 CLI）結構刻意保持一致（--folder / --no-recursive /
逐檔錯誤隔離 / batch_log.csv / XLSX 控制字元清理），但兩者完全獨立，互不 import
清洗邏輯，只共用 excel_io.py 最底層的純 I/O 讀檔函式（不再依賴整份 boq_cleaner.py）。

【重要】quote_cleaner.py 的欄位偵測邏輯尚未用大量真實報價單驗證過準確度
（不同於 boq_cleaner.py 已用多份真實標單回歸驗證）。批次跑完後，請務必抽查
batch_log.csv 的狀態欄與 clean_items.csv 的「複價驗算」欄，確認欄位辨識正確，
不要未經核對就直接拿去比價或议价。
"""
import argparse
import os
import re
import sys
import time
import pandas as pd
from quote_cleaner import clean_quote_workbook, items_to_rows

SUPPORTED_EXT = ('.xlsx', '.xls', '.xlsm')  # v1.1(F-4): 移除 .xlsb，openpyxl 不支援該格式，宣稱支援但實測必失敗

# XLSX（ECMA-376／XML 1.0）不允許的控制字元：0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F
_ILLEGAL_XLSX_CHARS_RE = re.compile('[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')


def sanitize_for_excel(value):
    if isinstance(value, str):
        return _ILLEGAL_XLSX_CHARS_RE.sub('', value)
    return value


def sanitize_df_for_excel(df):
    if df.empty:
        return df
    return df.map(sanitize_for_excel)


def is_temp_lock_file(path):
    return os.path.basename(path).startswith('~$')


def collect_files(file_args, folder, recursive):
    """副檔名比對一律轉小寫再比較，跨平台行為一致（理由同 export_cleaned.py）。"""
    files = list(file_args or [])

    if folder:
        if not os.path.isdir(folder):
            print(f'錯誤：找不到資料夾 {folder}', file=sys.stderr)
            sys.exit(1)
        if recursive:
            for root, _dirs, names in os.walk(folder):
                for name in names:
                    if name.lower().endswith(SUPPORTED_EXT):
                        files.append(os.path.join(root, name))
        else:
            for name in os.listdir(folder):
                full = os.path.join(folder, name)
                if os.path.isfile(full) and name.lower().endswith(SUPPORTED_EXT):
                    files.append(full)

    files = sorted({os.path.normpath(f) for f in files if not is_temp_lock_file(f)})
    return files


def classify_file_result(summary):
    """判斷整份檔案的批次結果：'ok' / 'skip' / 'error'。理由同標單 CLI 的
    F-2 修正——單純略過（找不到可信表頭、缺必要欄位、總表）不是錯誤，
    不該讓整批 exit code 非 0；只有真正的讀取錯誤才算失敗。"""
    statuses = [str(s.get('status', '')) for s in summary]
    if any(s == 'OK' for s in statuses):
        return 'ok'
    if any(s.startswith('讀取錯誤') or s.startswith('讀取失敗') for s in statuses):
        return 'error'
    return 'skip'


def main():
    ap = argparse.ArgumentParser(description='工程報價單批次清洗 CLI')
    ap.add_argument('files', nargs='*', help='Excel 檔案路徑（可與 --folder 併用，也可單獨使用）')
    ap.add_argument('--folder', help='資料夾路徑：批次處理裡面全部 .xlsx/.xls/.xlsm/.xlsb')
    ap.add_argument('--no-recursive', action='store_true', help='--folder 模式下不遞迴子資料夾（預設遞迴）')
    ap.add_argument('--outdir', default='output', help='輸出資料夾')
    ap.add_argument('--include-summary', action='store_true', help='不要略過總表工作表')
    ap.add_argument('--vendor', help='統一指定廠商名稱（覆蓋檔名遞補；適合單一廠商多檔案的情境）')
    ap.add_argument('--strict', action='store_true',
                     help='連「略過（無可用工作表）」也視為失敗並讓 exit code 非 0（預設只有讀取錯誤才失敗）')
    args = ap.parse_args()

    files = collect_files(args.files, args.folder, recursive=not args.no_recursive)
    if not files:
        print('找不到任何 Excel 檔案。請確認 files 參數或 --folder 路徑是否正確。', file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.outdir, exist_ok=True)

    all_rows, all_summary = [], []
    ok_files, skip_files, error_files = 0, 0, 0
    t0 = time.time()

    print(f'共找到 {len(files)} 個檔案，開始批次清洗（報價單）…\n')

    for idx, f in enumerate(files, 1):
        base = os.path.basename(f)
        print(f'[{idx}/{len(files)}] {base} ...', end=' ', flush=True)
        file_t0 = time.time()

        try:
            items, summary = clean_quote_workbook(
                f, skip_summary=not args.include_summary, vendor_name=args.vendor)
        except Exception as e:
            print(f'讀取失敗：{e}')
            all_summary.append({'sheet': '(整份檔案)', 'status': f'讀取失敗：{e}',
                                 'rows': 0, 'items': 0, 'file': base})
            error_files += 1
            continue

        for r in items_to_rows(items):
            all_rows.append(r)
        for s in summary:
            s['file'] = base
            all_summary.append(s)

        result = classify_file_result(summary)
        if result == 'ok':
            ok_files += 1
        elif result == 'error':
            error_files += 1
        else:
            skip_files += 1

        elapsed = time.time() - file_t0
        print(f'完成，{len(items)} 項（{elapsed:.2f}s）')

    total_elapsed = time.time() - t0

    clean_df = pd.DataFrame(all_rows)
    summary_df = pd.DataFrame(all_summary)

    batch_log_cols = ['file', 'sheet', 'status', 'rows', 'header_row', 'items']
    batch_log_df = summary_df.reindex(columns=batch_log_cols) if not summary_df.empty else pd.DataFrame(columns=batch_log_cols)

    clean_csv = os.path.join(args.outdir, 'clean_quote_items.csv')
    batch_log_csv = os.path.join(args.outdir, 'batch_log.csv')
    xlsx_path = os.path.join(args.outdir, 'structured_quote_result.xlsx')

    clean_df.to_csv(clean_csv, index=False, encoding='utf-8-sig')
    batch_log_df.to_csv(batch_log_csv, index=False, encoding='utf-8-sig')

    clean_df_x = sanitize_df_for_excel(clean_df)
    summary_df_x = sanitize_df_for_excel(summary_df)
    batch_log_df_x = sanitize_df_for_excel(batch_log_df)
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        batch_log_df_x.to_excel(writer, index=False, sheet_name='BatchLog')
        summary_df_x.to_excel(writer, index=False, sheet_name='Summary')
        clean_df_x.to_excel(writer, index=False, sheet_name='CleanItems')

    print()
    print('=' * 50)
    print(f'批次完成，耗時 {total_elapsed:.1f} 秒')
    print(f'檔案：成功 {ok_files} / 略過（無可用工作表，非錯誤）{skip_files} / '
          f'讀取錯誤 {error_files} / 共 {len(files)}')
    print(f'品項總數：{len(clean_df)}')
    if not clean_df.empty:
        mismatch = int((clean_df['複價驗算'] == '不符').sum())
        no_check = int((clean_df['複價驗算'] == '無法驗算').sum())
        print(f'複價驗算：不符 {mismatch} 筆／無法驗算 {no_check} 筆（建議人工核對）')
    print(f'clean_quote_items -> {clean_csv}')
    print(f'batch_log         -> {batch_log_csv}')
    print(f'excel             -> {xlsx_path}')
    print()
    print('提醒：欄位偵測尚未用大量真實報價單驗證過，請抽查 batch_log.csv 與')
    print('      複價驗算欄，確認辨識結果正確後再用於比價或議價。')

    if skip_files:
        print(f'\n提醒：有 {skip_files} 個檔案沒有任何可用工作表（找不到可信表頭／'
              f'缺必要欄位／僅總表），已略過，詳見 batch_log.csv。這不算錯誤，'
              f'exit code 不受影響。', file=sys.stderr)

    if error_files:
        print(f'\n錯誤：有 {error_files} 個檔案讀取失敗，詳見 batch_log.csv', file=sys.stderr)
        sys.exit(1)
    if args.strict and skip_files:
        print(f'\n--strict：略過檔案數 > 0，視為失敗。', file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
