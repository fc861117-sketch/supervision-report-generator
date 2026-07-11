# 督導報告表生成

這是一個可部署到 GitHub Pages 的靜態網站。輸入督導日期、時間、督導人員與督導單位後，瀏覽器會套用 `template.docx` 產生 Word 檔。

網站使用本地 `vendor/pizzip.min.js`，不依賴外部 CDN。

## 使用方式

1. 開啟 `index.html`
2. 輸入日期與時間
3. 按「產生 Word 檔」

## GitHub Pages

將此資料夾推到 GitHub repository 後，到 repository 的 `Settings` -> `Pages`，選擇 `main` branch 與 root folder，即可發布網站。

## 範本

`template.docx` 由原始範本去識別化後保留版面。網站會替換日期、時間、督導人員與督導單位欄位。
