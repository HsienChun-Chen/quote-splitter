# 標單拆分詢價單工具

## 專案說明

把公家單位的投標標價清單（Excel / PDF / Word）自動拆分成每家廠商的詢價單。
產出格式與「剛剛好室內裝修有限公司」既有的報價單範本完全一致。

---

## 檔案結構

```
├── 標單拆分工具.html     ← 主工具（瀏覽器直接打開使用，無需安裝）
├── build_quote.py        ← Python 版核心邏輯（本機執行，可搭配 LibreOffice 轉 PDF）
└── README.md
```

---

## 網頁工具使用方式

1. 用瀏覽器打開 `標單拆分工具.html`
2. **步驟 1**：上傳標單（支援 `.xls` / `.xlsx` / `.pdf` / `.docx`）
3. **步驟 2**：確認工程資訊（自動從標單帶入）
4. **步驟 3**：新增廠商，把品項分配給各廠商（可批次勾選指派）
5. **步驟 4**：點「產出所有詢價單」→ 下載 ZIP

ZIP 裡每家廠商有獨立的 `.xlsx`。超過 15 項自動分成 `_P1.xlsx`、`_P2.xlsx`。

---

## Python 版使用方式（build_quote.py）

### 安裝依賴

```bash
pip install openpyxl
```

### 範本檔

需要原始的 `活頁簿1.xlsx`（報價單範本），請放在同一目錄下。

### 基本用法

```python
from build_quote import build_quote

items = [
    {"is_section": True, "desc": "PVC地磚工程"},
    {"no": 1, "desc": "地坪,自平水泥+15*90cmPVC地磚A", "unit": "M2", "qty": 1402.00, "note": "-"},
    {"no": 2, "desc": "側封板(H=20CM,面板橫向鋪設)",    "unit": "M",  "qty": 46.00,   "note": "-"},
]

project = {
    "contact": "陳憲君",
    "tel":     "0912-687-638",
    "mail":    "a7788995720@gmail.com",
    "name":    "屏東區管理處暨營運所辦公廳興建工程",
    "site":    "屏東市林森路"
}

build_quote(
    template_path="活頁簿1.xlsx",
    output_path="詢價單_A廠商.xlsx",
    vendor_name="A廠商",
    project=project,
    items=items
)
```

### 轉 PDF（需要 LibreOffice）

```bash
# macOS
brew install libreoffice
soffice --headless --convert-to pdf 詢價單_A廠商.xlsx

# Ubuntu
sudo apt install libreoffice
soffice --headless --convert-to pdf 詢價單_A廠商.xlsx
```

---

## 關鍵設計說明

### 排版邏輯

| 問題 | 解法 |
|------|------|
| 範本第 4,8,11,13,15 列是 B:I 合併的分類列 | unmerge → 重設樣式 → 重建 B:C / H:I 合併 |
| 廠商名稱放哪（LOGO 占 B2:H2） | 覆蓋 I2 的 TODAY() 公式 |
| 灰色淡底殘留（29-32 列合計區） | fill 設為 none |
| 外部連結公式 `=OFFSET([1]毛管...` 造成 #REF | 掃描並清除所有含 `[1]` 的公式 |
| 超過 15 項時版面崩版 | 每頁固定 15 項，多餘列壓縮到 3pt |
| 品名過長文字截斷 | D 欄 wrap_text，依字數估算列高 |

### 分頁邏輯

- 每頁最多 **15 項**（含分類列）
- 最後一項剛好滿 15 時**不切頁**（避免空白第二頁）
- 換頁時帶出當前分類名稱 +「（續）」

### 欄位對應

| Excel 欄 | 內容 |
|---------|------|
| B:C（合併）| 項次 |
| D | 品項名稱（wrap_text 自動換行） |
| E | 數量 + 單位（如 `1402.00 M2`） |
| F | 單價（留空，廠商填寫） |
| G | 複價（留空，廠商填寫） |
| H:I（合併）| 備註 |

---

## 可調整的參數（build_quote.py 頂部）

```python
MAX_ROWS       = 15   # 每頁最多幾項（含分類列）
DESC_COL_CHARS = 18   # D 欄每行容納約幾個中文字
BASE_ROW_H     = 18   # 單行基礎高度（pt）
```

---

## 標單解析支援格式

| 格式 | 解析方式 | 可靠度 |
|------|---------|--------|
| `.xlsx` / `.xls` | 自動找「詳細表」分頁，識別項次/項目/數量/單位表頭 | ★★★★★ |
| `.docx` | 解析 XML，找表格標題列 | ★★★★☆ |
| `.pdf` | 文字提取，正則匹配數量+單位 | ★★★☆☆ |

政府標單通常是 Excel，解析最穩定。

---

## 已知限制

- PDF 解析對掃描版 / 複雜版面效果有限
- 列高估算是近似值（18 字/行），實際視字型和字元寬度有差
- 輸出 PDF 建議用 Excel「另存新檔 → PDF」，比工具直接產生的品質更好

---

## 如果要擴充

- **要改一頁項數**：修改 `MAX_ROWS`
- **要加更多工程資訊欄位**（如業主、工程編號）：在 `fill_sheet` 裡對應的列加 `ws["D42"]` 之類
- **要支援其他範本**：替換 `活頁簿1.xlsx`，並重新確認 `FIRST_ROW`、合併欄位位置
