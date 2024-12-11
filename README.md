# 中國快訊PDF生成與合併程式
<div align="center">
    <img src="https://hackmd.io/_uploads/r1pXuw6nA.png" width="200" alt="TEJ">
    <img src="https://hackmd.io/_uploads/rkHVRqanC.png" width="200" alt="TWCC">
</div>

## 下載所有程式
```
git clone git@tej-gitlab.tejwin.com:YJB/CN-NEWS-PDF.git
```

## 安裝套件
```
pip install -r requirements.txt
```

## 執行程式
```
python CCRQM_PDF.py
```
![2024-09-10 18.02.27](https://hackmd.io/_uploads/HkQstc63R.gif)

## 使用方式
1. **TEJ API KEY**：
打開程式碼，找到以下程式碼 -> 將 `"Your Key"` 替換為您的 TEJ API Key。
```Python
tejapi.ApiConfig.api_key  = "Your KEY"
tejapi.ApiConfig.api_base = "http://10.10.10.66"
```

2. **設置日期範圍**：
Terminal 輸入起迄日期，格式為 `"2024-01-01"`

3. **設置字體**：
程式中使用的字體為 `SimSun`，若讀取不到字體請將ttc檔之路徑絕對路徑。
```Python
# 定義字體
pdfmetrics.registerFont(TTFont('SimSun', 'simsun.ttc'))
```

4. **文件路徑**：
若讀取不到檔案請改為絕對路徑。
- `pdf_file`：生成的 PDF 保存路徑。
- `logo_path`：PDF 頂部 Logo 圖片的路徑。
- `pdfs_to_merge`：需要合併的 PDF 文件路徑列表。

## 文件說明
| 文件名稱                   | 說明                           |
| ---------------------- | ---------------------------- |
| 中国快讯.pdf               | 從資料庫導出的表格文件。                 |
| CCRQM.pdf              | 關於CCRQM的介紹文件。                |
| CCRQM 沪深上市企业风险事件早报.pdf | 最終生成的報告文件，包含了中國快訊與風險事件的整合內容。 |