# FileRecoveryTool

Intelligent post-recovery organization tool for Office documents.

This tool is designed for scenarios after file recovery where original filenames are lost, contents are mixed, and manual identification is difficult. It analyzes file content, hashes, formats, and naming rules to help users quickly identify primary files, duplicates, and corrupted files, and generate suggested filenames and analysis reports.

---

![screenshot](https://github.com/Dorlamon/FileRecoveryTool/blob/main/screenshot.en-us.png)

---

## Core Purpose

When original filenames are lost after disk recovery (e.g., missing MFT), this tool helps:

- Infer reasonable filenames from content
- Detect duplicate files
- Classify Primary / Duplicate / Unique / Corrupted files
- Preview rename results
- Execute actual renaming
- Export HTML analysis reports
- Organize output files based on rules

---

## Main Features

- Supported formats:
  - docx, xlsx, pptx, doc, xls, ppt, rtf, pdf, txt
- SHA256 deduplication
- Content fingerprint comparison
- Full Excel worksheet parsing
- Intelligent Excel naming
- File classification
- Rename simulation & execution
- HTML report export
- Open latest report
- Safe organize mode (Copy / Move)
- Primary-only mode
- Organize by extension
- Quality score
- Bilingual UI
- Encryption detection (Office / PDF / RTF)

---

## Project Structure

FileRecoveryTool/
- OfficeRecoveryToolkit.cmd
- OfficeRecoveryToolkit.ps1
- OfficeEncryptionProbe.Program.cs
- OfficeEncryptionProbe.csproj
- GUIDE.txt
- WHAT'S NEW.txt
- README.md

---

## Requirements

- Windows 10 / 11
- PowerShell 5.1
- Microsoft Office compatibility pack recommended

---

## Usage

Run:

OfficeRecoveryToolkit.cmd

---

## Workflow

1. Run tool
2. Set scan folder
3. Scan
4. Review results
5. Export report
6. Simulate rename
7. Execute rename

---

## File Classification

- Primary
- Duplicate
- Unique
- Corrupted

---

## Output

- HTML report
- CSV preview
- Rename results
- Organized files

---

## Safety Tips

- Always scan first
- Review report
- Use simulation
- Backup important data

---

## Limitations

- Windows only
- Legacy formats may need conversion
- Accuracy depends on file integrity

---

## Version

v5.8.5.12
