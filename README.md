# Outlook Markdown 匯出工具

這是一個 Windows 桌面版 Outlook 郵件匯出工具。程式會透過選單讓你指定日期範圍，讀取 `scan_set.ini` 裡列出的 Outlook 資料夾，將符合條件的郵件匯出成 Markdown 檔案，並把附件與郵件內圖片集中保存到 `附件` 資料夾。

## 重要需求

- 作業系統：Windows。
- 郵件程式：Windows 桌面版 Outlook，不是網頁版 Outlook。
- Python：建議使用 Python 3.14 或更新版本。
- 執行程式前，請先開啟 Outlook 桌面程式，並確認郵件已同步到本機。
- 第一次讀取 Outlook 時，Outlook 可能會跳出安全性提示，請允許程式存取郵件資料。

## 安裝 Python

如果電腦尚未安裝 Python，請先到 Python 官方網站下載並安裝：

[https://www.python.org/downloads/](https://www.python.org/downloads/)

安裝時建議勾選 `Add python.exe to PATH`。安裝完成後，在 PowerShell 確認版本：

```powershell
py --version
```

若能看到類似以下結果，代表 Python 可正常使用：

```text
Python 3.14.4
```

## 安裝必要套件

第一次使用前，請在 PowerShell 執行：

```powershell
py -m pip install pywin32 beautifulsoup4 markdownify
```

套件用途：

- `pywin32`：讓 Python 讀取 Windows 桌面版 Outlook。
- `beautifulsoup4`：協助處理 HTML 郵件內容。
- `markdownify`：將 HTML 郵件轉換成 Markdown。

## 檔案說明

- `outlook_export_menu.py`：主要 Python 選單程式。
- `scan_set.ini`：指定要掃描的 Outlook 資料夾。
- `README.md`：使用說明。
- `export_outlook_to_markdown.ps1`：早期 PowerShell 版本，保留備用。

## 設定掃描資料夾

程式只會掃描 `C:\Codex\scan_set.ini` 裡列出的 Outlook 資料夾。每行填一個資料夾名稱：

```ini
Inbox
收件匣
客戶郵件
```

如果不同信箱裡有同名資料夾，建議使用完整路徑，避免掃到不想要的資料夾：

```ini
\\your.name@example.com\Inbox\客戶郵件
```

設定規則：

- 每行一個資料夾。
- 空白行會略過。
- `#` 開頭的行會當成註解。
- 可以使用資料夾名稱或完整 Outlook 資料夾路徑。

## 啟動程式

請先開啟 Outlook 桌面程式，然後在 PowerShell 執行：

```powershell
py C:\Codex\outlook_export_menu.py
```

程式會顯示選單：

```text
1. 依日期範圍匯出 scan_set.ini 指定資料夾郵件
2. 查看目前設定與匯出紀錄
3. 修改輸出資料夾
4. 離開
```

## 匯出流程

1. 編輯 `scan_set.ini`，填入要掃描的 Outlook 資料夾。
2. 開啟 Windows 桌面版 Outlook。
3. 執行 `py C:\Codex\outlook_export_menu.py`。
4. 選擇功能 `1`。
5. 輸入開始日期與結束日期，格式為 `YYYY-MM-DD`，例如 `2026-05-14`。
6. 等待程式掃描指定資料夾並輸出 Markdown。

## 輸出位置

預設輸出到：

```text
C:\Codex\outlook_markdown
```

輸出結構範例：

```text
outlook_markdown\
  .export_manifest.json
  20260514_093000_王小明_會議紀錄.md
  20260514_101500_客戶A_報價確認.md
  附件\
    20260514_101500_客戶A_報價確認_01_quote.xlsx
    20260514_101500_客戶A_報價確認_02_image.png
```

## 檔名規則

郵件 Markdown 檔名格式：

```text
日期時間_寄信人_主旨.md
```

例如：

```text
20260514_101500_客戶A_報價確認.md
```

不適合放進 Windows 檔名的符號會自動替換，過長的寄件人或主旨也會自動截短。

## 附件與圖片

- 所有附件會集中存放在 `附件` 資料夾。
- 一般附件會以 Markdown 連結方式寫入郵件檔。
- 圖片附件會以 Markdown 圖片語法寫入郵件檔。
- 郵件內嵌圖片會嘗試依 Outlook 的 Content-ID 對應到已保存的圖片檔。

Markdown 範例：

```markdown
## Attachments

- [報價單.xlsx](附件/20260514_101500_客戶A_報價確認_01_報價單.xlsx)
- ![image.png](附件/20260514_101500_客戶A_報價確認_02_image.png)
```

## 避免重複匯出

程式會在輸出資料夾建立 `.export_manifest.json`，記錄已匯出的郵件。下次使用同樣日期範圍時，已匯出的郵件會自動略過。

若早期版本已匯出郵件但附件缺失，本版會自動補跑缺少附件版本標記的郵件，補完後才恢復略過規則。

## 常見問題

### 一定要開 Outlook 嗎？

建議一定要先開啟 Windows 桌面版 Outlook。Outlook 已開啟時，程式讀取郵件比較穩定，也比較容易處理安全性提示。

### 可以讀 Outlook 網頁版嗎？

不行。這個工具是讀取本機 Outlook 桌面程式中的資料。

### 沒有看到某些郵件怎麼辦？

請確認 Outlook 已經把那些郵件同步到本機，也確認該資料夾有寫在 `scan_set.ini` 裡。

### 附件或圖片沒有出現怎麼辦？

請先確認郵件本身在 Outlook 裡可以正常開啟附件。若同一封郵件已由早期版本匯出過，本版會自動補跑附件；如果仍有問題，可以刪除輸出資料夾中的 `.export_manifest.json` 後重新匯出，但這會讓所有郵件重新處理一次。

### 日期格式怎麼填？

請使用：

```text
YYYY-MM-DD
```

例如：

```text
2026-05-14
```

## 開發者備註

主要程式入口是 `outlook_export_menu.py`。目前使用 Outlook COM 介面讀取本機郵件，並使用 `markdownify` 轉換 HTML 郵件內容。匯出狀態保存在 `.export_manifest.json`，附件版本由程式內的 `ATTACHMENT_EXPORT_VERSION` 控制。
