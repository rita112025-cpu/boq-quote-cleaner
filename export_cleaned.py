"""
工程標單／報價單批次清洗 CLI — v1.1

用法：
    # 明確列出檔案（原有用法，向下相容）
    python export_cleaned.py A.xlsx B.xls --outdir output

    # 指向資料夾，批次處理裡面全部 Excel（預設遞迴子資料夾）
    python export_cleaned.py --folder ./標單資料夾 --outdir output

    # 資料夾模式但不遞迴子資料夾
    python export_cleaned.py --folder ./標單資料夾 --no-recursive --outdir output

    # 檔案與資料夾可併用
    python export_cleaned.py extra.xlsx --folder ./標單資料夾 --outdir output

清洗核心（boq_cleaner.py）完全不變動；本檔只負責批次蒐集檔案、
逐檔呼叫、錯誤隔離與彙整輸出，不修改任何清洗規則或判斷邏輯。
"""
import argparse
import os
import re
import sys
import time
import pandas as pd
from boq_cleaner import clean_workbook, items_to_rows

SUPPORTED_EXT = ('.xlsx', '.xls', '.xlsm')  # v1.1(F-4): 移除 .xlsb，openpyxl 不支援該格式，宣稱支援但實測必失敗

# XLSX（ECMA-376／XML 1.0）不允許的控制字元：0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, 0x7F
# （保留 \t \n \r）。真實標單常見於 OCR／複製貼上殘留的隱藏控制碼，例如 U+0002。
_ILLEGAL_XLSX_CHARS_RE = re.compile('[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]')


def sanitize_for_excel(value):
    """移除 openpyxl／XLSX 不允許的控制字元，避免整批 Excel 輸出因單一儲存格中斷。
    只影響 Excel 輸出；CSV 不受影響，保留原始資料（含控制字元）。"""
    if isinstance(value, str):
        return _ILLEGAL_XLSX_CHARS_RE.sub('', value)
    return value


def sanitize_df_for_excel(df):
    if df.empty:
        return df
    return df.map(sanitize_for_excel)


def is_temp_lock_file(path):
    """Excel 開啟中會產生 ~$檔名.xlsx 鎖定暫存檔，讀取會失敗，應直接跳過。"""
    return os.path.basename(path).startswith('~$')


def collect_files(file_args, folder, recursive):
    """彙整待處理檔案清單：明確列出的檔案 + 資料夾掃描結果，去重、排序、排除鎖定檔。

    副檔名比對一律轉小寫再比較（不用 glob('*.xls') 直接比對），因為 Linux／macOS
    的 glob 對副檔名大小寫敏感，會靜默漏掉如 .XLS（全大寫）的檔案；Windows 則不敏感。
    若依賴 glob 原生行為，同一支程式在不同作業系統上會找到不同數量的檔案且不會報錯，
    是很難察覺的資料遺漏。改用手動遍歷＋小寫比對可確保跨平台行為一致。
    """
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

    # 去重、排除 Excel 鎖定暫存檔、排序（確保每次執行順序一致，方便對照 log）
    files = sorted({os.path.normpath(f) for f in files if not is_temp_lock_file(f)})
    return files


def classify_file_result(summary):
    """判斷整份檔案的批次結果：'ok' / 'skip' / 'error'。

    v1.1 修正（F-2）：原本只要沒有任何 OK 工作表就一律算 fail_files，導致
    「這份檔案本來就不是標單」（例如封面、說明頁，正確地被略過）跟「檔案
    真的讀取失敗」被混為一談，一起計入失敗、一起讓整批 exit code=1。對排
    程／CI 依 exit code 判斷成敗的情境，資料夾裡混一張說明頁就會讓整批被
    誤判失敗。現在只有真正的讀取錯誤才算失敗；單純略過（找不到可信表頭、
    總表）不影響 exit code。
    """
    statuses = [str(s.get('status', '')) for s in summary]
    if any(s == 'OK' for s in statuses):
        return 'ok'
    if any(s.startswith('讀取錯誤') or s.startswith('讀取失敗') for s in statuses):
        return 'error'
    return 'skip'


def main():
    ap = argparse.ArgumentParser(description='工程標單／報價單批次清洗 CLI')
    ap.add_argument('files', nargs='*', help='Excel 檔案路徑（可與 --folder 併用，也可單獨使用）')
    ap.add_argument('--folder', help='資料夾路徑：批次處理裡面全部 .xlsx/.xls/.xlsm/.xlsb')
    ap.add_argument('--no-recursive', action='store_true', help='--folder 模式下不遞迴子資料夾（預設遞迴）')
    ap.add_argument('--outdir', default='output', help='輸出資料夾')
    ap.add_argument('--include-summary', action='store_true', help='不要略過總表工作表')
    ap.add_argument('--strict', action='store_true',
                     help='連「略過（無可用工作表）」也視為失敗並讓 exit code 非 0（預設只有讀取錯誤才失敗）')
    args = ap.parse_args()

    files = collect_files(args.files, args.folder, recursive=not args.no_recursive)
    if not files:
        print('找不到任何 Excel 檔案。請確認 files 參數或 --folder 路徑是否正確。', file=sys.stderr)
        sys.exit(1)

    os.makedirs(args.outdir, exist_ok=True)

    all_rows, all_summary, all_audit = [], [], []
    ok_files, skip_files, error_files = 0, 0, 0
    t0 = time.time()

    print(f'共找到 {len(files)} 個檔案，開始批次清洗…\n')

    for idx, f in enumerate(files, 1):
        base = os.path.basename(f)
        print(f'[{idx}/{len(files)}] {base} ...', end=' ', flush=True)
        file_t0 = time.time()

        # 逐檔錯誤隔離：單一檔案不論任何原因失敗，都不中斷整批處理。
        # clean_workbook 內部已對「整份檔案讀取失敗」與「單一工作表讀取失敗／
        # 找不到表頭」個別捕捉並記錄在 summary，這裡的 try/except 是最外層保險
        # （例如檔案權限問題等 clean_workbook 內部未預期的例外）。
        try:
            items, summary, audit = clean_workbook(f, skip_summary=not args.include_summary)
        except Exception as e:
            print(f'讀取失敗：{e}')
            all_summary.append({
                'sheet': '(整份檔案)', 'status': f'讀取失敗：{e}',
                'rows': 0, 'items': 0, 'unclassified': 0, 'file': base,
            })
            error_files += 1
            continue

        file_item_count = 0
        for r in items_to_rows(items):
            r['來源檔案'] = base
            all_rows.append(r)
            file_item_count += 1
        for s in summary:
            s['file'] = base
            all_summary.append(s)
        for a in audit:
            a['file'] = base
            all_audit.append(a)

        # 三分類：成功／略過（無錯誤，例如封面或找不到可信表頭）／讀取錯誤。
        # 只有讀取錯誤才計入失敗，略過不影響 exit code（見 classify_file_result）。
        result = classify_file_result(summary)
        if result == 'ok':
            ok_files += 1
        elif result == 'error':
            error_files += 1
        else:
            skip_files += 1

        elapsed = time.time() - file_t0
        print(f'完成，{file_item_count} 項（{elapsed:.2f}s）')

    total_elapsed = time.time() - t0

    clean_df = pd.DataFrame(all_rows)
    summary_df = pd.DataFrame(all_summary)
    audit_df = pd.DataFrame(all_audit)

    # 批次總覽（每檔每工作表一列，直接反映 clean_workbook 原生的 status 文字，
    # 不另外發明英文狀態碼，維持與清洗核心原生 status 一致的用語）
    batch_log_cols = ['file', 'sheet', 'status', 'rows', 'header_row', 'items', 'unclassified']
    batch_log_df = summary_df.reindex(columns=batch_log_cols) if not summary_df.empty else pd.DataFrame(columns=batch_log_cols)

    clean_csv = os.path.join(args.outdir, 'clean_items.csv')
    audit_csv = os.path.join(args.outdir, 'row_audit.csv')
    batch_log_csv = os.path.join(args.outdir, 'batch_log.csv')
    xlsx_path = os.path.join(args.outdir, 'structured_result.xlsx')

    clean_df.to_csv(clean_csv, index=False, encoding='utf-8-sig')
    audit_df.to_csv(audit_csv, index=False, encoding='utf-8-sig')
    batch_log_df.to_csv(batch_log_csv, index=False, encoding='utf-8-sig')

    # Excel（XLSX/ECMA-376）不允許部分控制字元；CSV 無此限制，保留原始資料。
    # 清理只影響 structured_result.xlsx，不回寫、不影響 CSV 或清洗結果本身。
    clean_df_x = sanitize_df_for_excel(clean_df)
    summary_df_x = sanitize_df_for_excel(summary_df)
    audit_df_x = sanitize_df_for_excel(audit_df)
    batch_log_df_x = sanitize_df_for_excel(batch_log_df)
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
        batch_log_df_x.to_excel(writer, index=False, sheet_name='BatchLog')
        summary_df_x.to_excel(writer, index=False, sheet_name='Summary')
        clean_df_x.to_excel(writer, index=False, sheet_name='CleanItems')
        audit_df_x.to_excel(writer, index=False, sheet_name='RowAudit')

    print()
    print('=' * 50)
    print(f'批次完成，耗時 {total_elapsed:.1f} 秒')
    print(f'檔案：成功 {ok_files} / 略過（無可用工作表，非錯誤）{skip_files} / '
          f'讀取錯誤 {error_files} / 共 {len(files)}')
    print(f'品項總數：{len(clean_df)}')
    print(f'clean_items -> {clean_csv}')
    print(f'row_audit   -> {audit_csv}')
    print(f'batch_log   -> {batch_log_csv}')
    print(f'excel       -> {xlsx_path}')

    if skip_files:
        print(f'\n提醒：有 {skip_files} 個檔案沒有任何可用工作表（找不到可信表頭／'
              f'僅總表），已略過，詳見 batch_log.csv。這不算錯誤，exit code 不受影響。',
              file=sys.stderr)

    if error_files:
        print(f'\n錯誤：有 {error_files} 個檔案讀取失敗，詳見 batch_log.csv', file=sys.stderr)
        sys.exit(1)
    if args.strict and skip_files:
        print(f'\n--strict：略過檔案數 > 0，視為失敗。', file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
