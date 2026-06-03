# quote-splitter

## 專案說明

把公家單位的投標標價清單（Excel / PDF / Word）自動拆分成每家廠商的詢價單。
產出格式與「剛剛好室內裝修有限公司」既有的報價單範本完全一致。

## 檔案

| 檔案 | 說明 |
|------|------|
| `標單拆分工具.html` | 主工具，瀏覽器直接打開，無需安裝 |
| `build_quote.py` | Python 版核心邏輯，可搭配 LibreOffice 轉 PDF |
| `README.md` | 詳細使用說明與設計文件 |

## 使用方式

1. 用瀏覽器打開 `標單拆分工具.html`
2. 上傳標單（`.xls` / `.xlsx` / `.pdf` / `.docx`）
3. 確認工程資訊 → 分配品項給各廠商 → 產出 ZIP（每廠商一份 `.xlsx`）

超過 15 項自動分頁為 `_P1.xlsx`、`_P2.xlsx`。

## 關鍵參數（build_quote.py）

```python
MAX_ROWS       = 15   # 每頁最多幾項（含分類列）
DESC_COL_CHARS = 18   # D 欄每行容納約幾個中文字
BASE_ROW_H     = 18   # 單行基礎高度（pt）
```

## 注意事項

- 修改前先 Read 確認現有內容
- 需要原始報價單範本 `活頁簿1.xlsx` 才能跑 Python 版
- Secrets / 廠商資料不得 commit 至版本控制
