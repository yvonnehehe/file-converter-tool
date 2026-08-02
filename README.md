# 📄 File Converter Tool

一個給 Windows 使用的簡易檔案格式轉換小工具，純視窗介面操作，不用打指令。

[![Build Windows EXE](https://github.com/yvonnehehe/file-converter-tool/actions/workflows/build-exe.yml/badge.svg)](https://github.com/yvonnehehe/file-converter-tool/actions/workflows/build-exe.yml)

---

## ✨ 功能

| 分頁 | 功能 | 需要額外軟體 |
|---|---|---|
| 格式轉換 | JPG / PNG → PDF | 不需要 |
| 格式轉換 | Word (docx) → PDF | 需要安裝 Microsoft Word |
| 格式轉換 | Excel (xlsx / xls) → PDF | 需要安裝 Microsoft Excel |
| 格式轉換 | PDF → JPG 圖片（每頁一張） | 不需要 |
| 格式轉換 | PDF → Word (docx) | 不需要 |
| PDF 合併 / 拆分 | 拆成單頁 PDF | 不需要 |
| PDF 合併 / 拆分 | 多個 PDF 合併成一個 | 不需要 |

操作方式很單純：選檔案 → 選要轉換的格式 → 選輸出位置 → 完成，全程視窗點擊，不用打任何指令。

---

## ⬇️ 下載使用（不需要安裝 Python）

1. 到 [**Actions**](https://github.com/yvonnehehe/file-converter-tool/actions/workflows/build-exe.yml) 頁面，點進最新一次成功的 run
2. 頁面下方 **Artifacts** 區塊，下載 `FileConverter-exe`（zip 檔）
3. 解壓縮後得到 `FileConverter.exe`，雙擊即可執行

> 需要登入 GitHub 帳號才能下載 Artifact，且 Artifact 保存 90 天。如果需要長期公開的下載連結，可以改發布 [Release](https://github.com/yvonnehehe/file-converter-tool/releases)。

---

## 🛠 開發 / 自行打包

**需求套件**（`requirements.txt`）：

```
pillow
docx2pdf
pywin32
PyMuPDF
pdf2docx
pypdf
```

**本機直接執行（需要 Python）：**

```bash
pip install -r requirements.txt
python file_converter_gui.py
```

**本機打包成 exe：**

雙擊 `build.bat`（Windows 環境），完成後 exe 會出現在 `dist/FileConverter.exe`。

**雲端自動打包：**

每次 push 到 `main` 分支，[GitHub Actions](.github/workflows/build-exe.yml) 會自動在雲端 Windows 環境打包一份新的 `FileConverter.exe`，不需要自己有 Windows 電腦。

---

## 📁 檔案結構

```
.
├── file_converter_gui.py        # 主程式（tkinter 視窗介面）
├── requirements.txt              # Python 套件需求
├── build.bat                     # 本機一鍵打包腳本（Windows）
└── .github/workflows/
    └── build-exe.yml             # 雲端自動打包設定
```

---

## ⚠️ 注意事項

- Word / Excel 轉 PDF 是透過呼叫本機安裝的 Microsoft Office（COM 自動化）完成，所以執行的電腦需要有安裝 Office。
- 其餘功能皆為純 Python 套件實作，不依賴任何外部軟體。
