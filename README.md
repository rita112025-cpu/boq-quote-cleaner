BOQ Semantic Cleaning Pipeline - Synthetic Edition
Cleans hierarchical MEP BOQ / quote sheets. Key decisions: dual-quantity disambiguation via price consistency, state machine for pre/post remarks, full row-level audit trail. All samples are synthetic
# boq-quote-cleaner — 工程標單／報價單清洗工具

把混亂的工程標單（BOQ）與廠商報價單 Excel，批次清洗成「每列都是可比對品項」的乾淨表格。純 Python CLI，無外部服務依賴。

> 本庫所有樣本均為人工合成的示範資料，不含任何真實專案、公司或廠商資訊。

---

## 為什麼需要這個工具

工程標單不是表格，是「長得像表格的階層式文件」。這是產業通病，不分公司：

- **多層階層與父子拆列**：`壹 → (一) → A → 1` 的分類語意散落在不同列，品項本身那一列往往看不出它屬於哪個系統。
- **語意繼承**：`A 高壓配電盤設備`、`1 MOF PANEL` 這種「小組名」列自己沒有數量單價，但下面所有品項都隱含這個前綴——直接逐列讀取會得到一堆無法辨識的「配電盤本體」。
- **雙語／重複表頭**：每頁重印表頭、中英文表頭混排（`Particulars & Description / Unit / Quantity`）。
- **雙數量欄**：「數量」與「工地數量」並存，複價實際只用其中一欄計算——取錯欄，整份成本比對就是錯的。
- **規格說明散落**：`- 含耐壓測試報告` 這種列可能是上一筆品項的備註，也可能是下一批品項的情境延伸，位置完全一樣、語意完全相反。

人工整理大型標單通常耗時、難稽核，而且錯誤不容易即時被發現——直到比價結果出現異常。

---

## 三個核心技術決策

### 1. 雙數量欄挑欄：用複價一致性反推

標單同時有「數量」「工地數量」兩欄時，哪一欄才是計價用的？表頭文字不可靠（每家叫法不同），欄位順序也不可靠。

唯一可靠的線索是**複價本身**：複價是用哪一欄算出來的，哪一欄就是計價欄。

```
對每一列：有單價與複價時，
    選「數量 × 單價 最接近複價」的那一欄作為計價數量
    另一欄輸出到「另一數量欄」（設計 vs 工地差異，可追變更）
單一數量欄的工作表維持原行為（第一個非空值）
```

同時輸出「複價驗算」欄（`OK` / `不符` / `無法驗算`，容差 `|複價|×1% + 1`），讓原始檔本身填錯的列浮上來交人工判斷，而不是被工具靜默吞掉。

實際效果可以用 `samples_synthetic/範例標單.xlsx` 驗證：`配電盤本體` 一列數量欄填 2、工地數量填 3、複價 150,000 = 3 × 50,000，工具會正確取 3，並把 2 記在「另一數量欄」。

### 2. 前置情境 vs 後置備註：狀態機

這是整個清洗最難的部分。同樣一列「空白項次＋純文字描述」，語意取決於它出現的位置：

```
1  MOF PANEL              ← 小組名（前置情境）
   含基礎型鋼及固定五金      ← 情境「延伸」：屬於下面所有品項
   配電盤本體  面 3 ...     ← 計價品項（消耗情境）
   - 含耐壓測試報告          ← 後置「備註」：只屬於上面那一筆
```

用位置規則（縮排、順序）判斷都會在真實檔案上翻車。解法是一個小型狀態機：

- `context_open` 旗標：小組名列宣告後開啟；第一筆計價品項出現時關閉。開啟期間的空白項次描述列是**情境延伸**（併入之後所有品項的繼承語意），關閉後才是**後置備註**（只掛回上一筆品項）。
- `category_boundary` 旗標：大分類（`壹`、`貳`）切換後到本分類第一筆品項出現前，規格列一律暫存，不誤掛回**上一個分類**的最後一筆品項。

分類（breadcrumb）、情境（context）、備註（notes）三層各自有明確的生命週期與清空時機，而不是一個「上一列是什麼」的 heuristic。

### 3. 逐列 Audit Trail：清洗過程必須可稽核

清洗工具最危險的不是報錯，是**靜默做錯**。所以每一列——包含被丟棄的列——都會留下分類紀錄輸出到 `row_audit.csv`：

```
row_no │ row_type            │ name                 │ breadcrumb / context │ reason
    7  │ CONTEXT_EXTEND      │ 含基礎型鋼及固定五金    │ ...                  │ blank item_no while context is open
    8  │ PRICE_ITEM          │ 配電盤本體             │ 電氣設備工程 / ...    │ has unit + qty
    9  │ SPEC_NOTE_AFTER_ITEM│ - 含耐壓測試報告       │ ...                  │
   14  │ SUBTOTAL            │ 小計                  │                      │
```

任何一筆清洗結果看起來怪，都能從 audit 反查「這一列當時被判成什麼、為什麼」。這個設計在開發期間直接抓出過兩個會靜默丟資料的 bug（表頭誤判、跨分類誤掛），修正前後的差異也全靠 audit 的分類統計量化驗證。

---

## 安裝與使用

```bash
pip install -r requirements.txt
```

**標單清洗**：

```bash
python export_cleaned.py --folder ./標單資料夾 --outdir output
python export_cleaned.py 範例標單.xlsx --outdir output
```

**報價單清洗**（扁平品項列表，獨立的欄位偵測邏輯）：

```bash
python export_cleaned_quotes.py --folder ./報價單資料夾 --outdir output
python export_cleaned_quotes.py A.xlsx --vendor 廠商名 --outdir output
```

兩者共同的批次特性：

- 逐檔錯誤隔離：單一檔案失敗不中斷整批，記錄在 `batch_log.csv` 繼續處理。
- 「略過」≠「失敗」：找不到可信表頭等正常略過不影響 exit code（適合排程／CI）；要讓略過也算失敗加 `--strict`。
- 自動排除 `~$` 鎖定暫存檔、副檔名不分大小寫。
- XLSX 輸出自動清除 Office 不允許的控制字元（複製貼上殘留常見）；CSV 保留原始位元組。
- 表頭信心不足或缺必要欄位時**整張略過並記錄原因，不臆測欄位**。

## 立即體驗（合成樣本）

```bash
python export_cleaned.py samples_synthetic/boq/範例標單.xlsx --outdir out_boq
python export_cleaned_quotes.py --folder samples_synthetic/quotes --outdir out_quote
```

預期結果見 `samples_synthetic/README.md`，可逐項核對。

## 測試

```bash
python -m unittest discover -s tests
```

## 驗證狀態（誠實揭露）

- `boq_cleaner.py`（標單）：核心邏輯以多份真實標單逐列回歸驗證過（原始資料涉及機密，不隨本庫公開；本庫附合成標單樣本供獨立驗證行為）。
- `quote_cleaner.py`（報價單）：目前驗證了「與原始 JS 實作行為一致」（38 項單元測試）＋合成樣本端到端，**尚未經大量真實報價單驗證欄位偵測準確度**。第一次使用請先小量試跑並核對 `batch_log.csv` 與「複價驗算」欄。

## 已知限制

- `_choose_qty` 逐列獨立挑欄，理論上同一工作表不同列可能取到不同數量欄；實務回歸未觀察到問題，屬理論邊界。
- 報價單表頭門檻（信心分數）與必要欄位（品名／單位／數量／單價）是刻意保守的設計：寧可整張略過並記錄，不臆測。沒有單位欄的報價單會被略過。
- `detect_header_field` 部分包含比對偏寬鬆（如「工程名稱」會因含「名稱」被視為品名欄候選），繼承自原始 JS 版行為，測試檔已記錄。

## 架構

```
excel_io.py               純 I/O 讀檔（.xlsx/.xlsm 走 openpyxl、.xls 走 xlrd）
boq_cleaner.py            標單清洗核心（階層狀態機、雙數量欄、audit）
quote_cleaner.py          報價單清洗核心（表頭評分、欄位映射、複價驗算）
export_cleaned.py         標單批次 CLI
export_cleaned_quotes.py  報價單批次 CLI
samples_synthetic/        合成示範資料（全部人工生成）
tests/                    單元測試
```

兩套清洗核心完全獨立、互不 import，只共用 `excel_io.py`——標單是階層式文件、報價單是扁平列表，清洗規則本質不同，請勿把兩種檔案混丟給同一支工具。

## License

MIT
