[English](README.md) | [繁體中文](README.zh.md)

# docx

一個用於建立、讀取、編輯和操作 Word 文件 (.docx) 的 Claude Code 技能。

## 說明

此技能提供全面的 DOCX 文件處理 -- 使用 JavaScript (docx-js) 建立新文件、透過 XML 操作編輯現有文件（解包/編輯/重新打包工作流程）、使用 pandoc 讀取內容、新增追蹤修訂和批註，以及依據 Office Open XML 結構描述驗證輸出。

## 功能特色

- **建立新文件**: 使用 docx-js 產生 .docx 檔案，完全控制樣式、表格、圖片、頁首/頁尾、頁面版面和目錄
- **編輯現有文件**: 將 .docx 解包為 XML，進行精確編輯，再以驗證重新打包
- **讀取與擷取內容**: 使用 pandoc 擷取文字，或以原始 XML 取得完整保真度
- **追蹤修訂**: 加入插入、刪除和格式變更，支援正確的紅線標記
- **批註**: 插入批註和串連回覆，使用正確的 XML 結構
- **格式轉換**: 透過 LibreOffice 將 .doc 轉為 .docx，匯出為 PDF 或圖片
- **驗證**: 依據 ISO/IEC 29500 和 Microsoft 擴充結構描述進行驗證

## 安裝

```bash
cp -r docx ~/.claude/skills/
```

### 相依套件

```bash
# 文件建立
npm install -g docx

# 文字擷取
brew install pandoc

# 格式轉換（選用）
brew install --cask libreoffice

# 圖片轉換（選用）
brew install poppler
```

## 使用方式

- 「建立一份專業的 Word 報告」
- 「編輯這個 .docx 檔案並加入追蹤修訂」
- 「從 document.docx 擷取文字」
- 「在 contract.docx 加入批註」
- 「把這個 .doc 檔案轉成 .docx」

## 檔案結構

```
docx/
  SKILL.md              -- 主要技能指引
  LICENSE.txt           -- 授權條款
  references/           -- （保留供參考文件使用）
  scripts/
    __init__.py
    accept_changes.py   -- 接受所有追蹤修訂（透過 LibreOffice）
    comment.py          -- 在解包的文件中新增批註
    office/
      pack.py           -- 重新打包 XML 為 .docx 並驗證
      unpack.py         -- 將 .docx 解包為可編輯的 XML
      validate.py       -- 基於結構描述的 OOXML 驗證
      soffice.py        -- LibreOffice 轉換包裝器
      helpers/          -- XML 操作工具
      schemas/          -- ISO/IEC 29500 和 Microsoft XSD 結構描述
      validators/       -- 格式專屬驗證器 (docx, pptx, redlining)
    templates/          -- 批註用 XML 範本
  assets/               -- （保留供未來資源使用）
```

## 授權

專有授權。完整條款請見 LICENSE.txt。

## 來源

改編自 [anthropics/skills](https://github.com/anthropics/skills/tree/main/skills/docx)。
