# Universal Markdown 極簡轉檔器

一款以 **PyQt6** 製作的 Windows 桌面 GUI 小工具，專門用來將常見文件格式快速轉成 Markdown。

本專案整合了兩條成熟技術路線：

- `mammoth`：負責 Word `.docx -> Markdown` 的既有穩定流程
- Microsoft 的 [markitdown](https://github.com/microsoft/markitdown)：負責多格式文件轉 Markdown 與 OCR / Vision 擴充能力

`markitdown` 本身非常強大，但原生定位偏向 CLI 與程式整合。對一般使用者而言，拖曳檔案、切換 OCR 模式、管理 Provider、測試模型連線，仍然不夠直觀。因此本專案將這套能力整理為一個 **深色系、可拖曳、可設定 Provider 的桌面 GUI 工具**。

---

## 專案特色

- 深色系 PyQt6 GUI，支援拖曳檔案直接轉換
- `.docx` 支援雙模式
  - `Mammoth 保留圖片`
  - `MarkItDown OCR`
- 其他格式支援 `MarkItDown 快速模式 / OCR 增強模式`
- 背景執行緒轉換，不會因 OCR 或大型檔案而讓整個視窗卡死
- 轉換中有明確互動提示，避免誤判為當機
- 支援多組 OpenAI-compatible Provider
- 可在 GUI 內新增 Provider、測試連線、讀取模型
- API Key 不寫入 `settings.json`，而是儲存在 **Windows Credential Manager**
- 可打包成 **Portable** 版本，方便在沒有 Python 環境的 Windows 電腦上使用

---

## 支援格式與路由邏輯

| 檔案格式 | 引擎 | 說明 |
| --- | --- | --- |
| `.docx` | `Mammoth` 或 `MarkItDown OCR` | 可在 GUI 中切換 |
| `.pdf` | `MarkItDown` | 可切換快速模式 / OCR 增強模式 |
| `.pptx` | `MarkItDown` | 可切換快速模式 / OCR 增強模式 |
| `.xlsx`, `.xls` | `MarkItDown` | 可切換快速模式 / OCR 增強模式 |
| `.csv` | `MarkItDown` | 可切換快速模式 / OCR 增強模式 |
| `.txt`, `.json`, `.xml`, `.html`, `.htm`, `.epub` | `MarkItDown` | 可直接轉為 Markdown |

---

## 系統需求

- Windows 10 / 11
- Python 3.10 以上
- 文件中的 Python 指令以虛擬環境 `.venv` 為例

說明：

- 目前 API Key 的安全儲存依賴 **Windows Credential Manager**
- 若只使用 Portable 版，則不需要另外安裝 Python

---

## 主要依賴套件

執行時主要使用：

- `PyQt6`
- `mammoth`
- `openai`
- `python-dotenv`
- `markitdown[pdf,pptx,xlsx,xls]`
- `markitdown-ocr`

若要自行打包 Portable 版，另外需要：

- `pyinstaller`

---

## 安裝方式

### 1. 建立虛擬環境

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. 安裝套件

```powershell
python -m pip install --upgrade pip
python -m pip install PyQt6 mammoth openai python-dotenv "markitdown[pdf,pptx,xlsx,xls]" markitdown-ocr
```

若需自行打包 Portable 版，再額外安裝：

```powershell
python -m pip install pyinstaller
```

---

## 執行方式

### 方式 A：直接以 Python 啟動

```powershell
.\.venv\Scripts\python.exe app.py
```

### 方式 B：使用批次檔啟動

可直接雙擊：

- `UniversalMarkdown.bat`

---

## 首次使用設定

首次啟動後，可先在 GUI 內點選「設定」：

1. 新增或選擇一組 Provider
2. 輸入 `Service Provider`
3. 輸入 `Base URL`
4. 輸入 `Model`
5. 輸入 `API Key`
6. 可使用「讀取模型」或「測試連線」確認設定是否正確

### Provider 說明

本工具支援任意 **OpenAI-compatible API Provider**，不綁死特定廠商。

目前內建只是方便起步的範例模板，例如：

- `NVIDIA_NIM`
- `Google_Gemini`

亦可自行新增其他相容服務。

---

## 設定與安全設計

### `settings.json` 會儲存什麼？

非敏感設定會寫入：

- `_UserSettings/settings.json`

內容包含：

- Provider 名稱
- Base URL
- Model
- GUI 選項
- 當前模式設定

### API Key 儲存在哪裡？

API Key **不會** 明文寫入 `settings.json`，而是儲存在：

- **Windows Credential Manager**

這樣的好處是：

- 分享原始碼或 Portable 版本時，不會把 API Key 一起帶出
- 每位使用者都可在各自的電腦上填入自己的金鑰

### `.env` 的用途

專案提供：

- `.env.example`

用途是：

- 提供 Base URL 範本
- 方便本機開發

分享版本不應放入真實 API Key。

---

## 使用流程

1. 開啟工具
2. 在設定頁面中設定至少一組 Provider
3. 選擇 `.docx` 模式
4. 選擇其他格式的 OCR 模式
5. 將檔案拖曳到視窗中
6. 等待背景轉換完成
7. 點選「開啟輸出資料夾」查看結果

---

## DOCX 模式說明

### Mammoth 保留圖片

適合：

- 保留 Word 內嵌圖片輸出
- 沿用既有 `docx -> markdown` 工作流程

### MarkItDown OCR

適合：

- Word 內含大量截圖、流程圖、圖片文字
- 需要更強的 Vision / OCR 理解能力

---

## 其他格式 OCR 模式說明

### 快速模式

適合：

- 以速度為優先
- 純文字為主的文件
- 不特別需要圖片 OCR

### OCR 增強模式

適合：

- PDF / PPTX 中含大量圖片、截圖、掃描頁
- 需要透過 Vision API 解析圖像內容

注意：

- OCR 增強模式會比快速模式慢很多，尤其在 PDF、PPTX 上更明顯
- 這是正常現象，不代表程式當機

---

## 打包 Portable 版本

專案已附上打包批次檔：

- `Build_UniversalMarkdown_Exe.bat`

執行後會輸出：

- `dist/UniversalMarkdown_Portable/`

其中主要執行檔為：

- `dist/UniversalMarkdown_Portable/UniversalMarkdown_Portable.exe`

### 重要說明

- 正式釋出版採用 **Portable 資料夾版**
- 需要保留整個 `UniversalMarkdown_Portable` 資料夾一起使用
- 發佈附件為：
  - `dist/UniversalMarkdown_Portable.zip`

---

## 專案與發佈內容

### 原始碼內容

Repository 主要包含：

- `app.py`
- `converter.py`
- `UniversalMarkdown.bat`
- `Build_UniversalMarkdown_Exe.bat`
- `.env.example`
- `README.md`

### Release 附件

Release 附件包含：

- `UniversalMarkdown_Portable.zip`
### 
---

## 專案結構

```text
MarkItDown/
├─ app.py
├─ converter.py
├─ UniversalMarkdown.bat
├─ Build_UniversalMarkdown_Exe.bat
├─ .env.example
├─ _UserSettings/
└─ dist/
   └─ UniversalMarkdown_Portable/
```

---

## 致謝

- Microsoft [markitdown](https://github.com/microsoft/markitdown)
- `mammoth`
- `PyQt6`

---

## 授權

本專案採用 `MIT License`。

第三方套件與相依元件仍各自適用其原始授權條款，使用前請一併確認。
