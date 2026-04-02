---
name: quarto-gost-espd
description: Use when creating or revising technical documentation in this repository for GOST 19.xxx or GOST 2.105-style text documents with DOCX/PDF output and Word template post-processing.
---

# QuartoGost ESPD

Use this skill for explanatory notes, software documentation, operating
instructions, technical descriptions, and similar text documents.

## Default source

- `templates\espd\espd-template.qmd`

## Build command

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType espd -InputFile <file.qmd> -OutputDir <dir> -Name <name> -EmbedFonts`

## Rules

1. Keep service pages in `resources\reference-docs\espd\reference.docx`.
2. Do not remove `%MAINTEXT%` or `%TOC%` from the reference DOCX.
3. Structure the main body by numbered sections, subsections, points, subpoints, and appendices according to GOST 2.105-95 section 4.1.
4. Use labels for figures, tables, equations, and sections.
5. Use Julia blocks for computed tables or charts.
6. Preserve the logic of title sheet, approval sheet, and change-registration sheet when the document type requires them.
7. When a document is large, it may be split into parts or books; keep numbering and naming consistent across those parts.
8. Keep charts in the grayscale document profile and preserve the exported `svg` files for reuse in technical appendices and related documents.
9. Fill cover and approval-sheet requisites through the `gost:` block in front matter so `%DOC_TITLE%`, `%DOC_CODE%`, `%APPROVER_TITLE%` and related placeholders are populated automatically.
10. Keep placeholders only for variable requisites such as names, positions, dates, codes, document designations, and approval data; do not replace static form text like ministry names, `УТВЕРЖДАЮ`, `СОГЛАСОВАНО`, or signature underlines with placeholders.

Read [required-sections.md](references/required-sections.md) when you need a
baseline structure.
Before finalizing, run [norm-control.md](references/norm-control.md).
