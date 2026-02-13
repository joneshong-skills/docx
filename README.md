[English](README.md) | [繁體中文](README.zh.md)

# docx

A Claude Code skill for creating, reading, editing, and manipulating Word documents (.docx files).

## Description

This skill provides comprehensive DOCX document handling -- creating new documents with JavaScript (docx-js), editing existing documents via XML manipulation (unpack/edit/repack workflow), reading content with pandoc, adding tracked changes and comments, and validating output against Office Open XML schemas.

## What It Does

- **Create new documents**: Generate .docx files using docx-js with full control over styles, tables, images, headers/footers, page layout, and table of contents
- **Edit existing documents**: Unpack .docx into XML, make precise edits, and repack with validation
- **Read and extract content**: Use pandoc for text extraction or raw XML access for full fidelity
- **Tracked changes**: Add insertions, deletions, and formatting changes with proper redlining
- **Comments**: Insert comments and threaded replies with proper XML structure
- **Convert formats**: Convert .doc to .docx via LibreOffice, export to PDF or images
- **Validate**: Schema-based validation against ISO/IEC 29500 and Microsoft extensions

## Install

```bash
cp -r docx ~/.claude/skills/
```

### Dependencies

```bash
# Document creation
npm install -g docx

# Text extraction
brew install pandoc

# Format conversion (optional)
brew install --cask libreoffice

# Image conversion (optional)
brew install poppler
```

## Usage

- "Create a professional report as a Word document"
- "Edit this .docx file and add tracked changes"
- "Extract the text from document.docx"
- "Add comments to the contract.docx"
- "Convert this .doc file to .docx"

## File Structure

```
docx/
  SKILL.md              -- Main skill instructions
  LICENSE.txt           -- License terms
  references/           -- (reserved for reference docs)
  scripts/
    __init__.py
    accept_changes.py   -- Accept all tracked changes (via LibreOffice)
    comment.py          -- Add comments to unpacked documents
    office/
      pack.py           -- Repack XML into .docx with validation
      unpack.py         -- Unpack .docx into editable XML
      validate.py       -- Schema-based OOXML validation
      soffice.py        -- LibreOffice wrapper for conversions
      helpers/          -- XML manipulation utilities
      schemas/          -- ISO/IEC 29500 and Microsoft XSD schemas
      validators/       -- Format-specific validators (docx, pptx, redlining)
    templates/          -- XML templates for comments
  assets/               -- (reserved for future assets)
```

## License

Proprietary. See LICENSE.txt for complete terms.

## Source

Adapted from [anthropics/skills](https://github.com/anthropics/skills/tree/main/skills/docx).
