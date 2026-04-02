---
name: quarto-gost-article
description: Use when creating or revising an article-like document in this repository as an extensible QuartoGost format derived from the shared bibliography, cross-reference, and Julia workflow.
---

# QuartoGost Article

Use this skill for papers, analytical notes, white papers, and other new text
formats that should reuse the project structure without the heavier GOST-specific
service pages.

## Default source

- `templates\article\article-template.qmd`

## Build command

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType article -InputFile <file.qmd> -OutputDir <dir> -Name <name>`

## Rules

1. Reuse bibliography, Julia chunks, and cross-reference patterns from the common partials.
2. Keep the structure simple unless the target publication imposes stricter rules.
3. If the user later needs journal-specific Word styles, add a dedicated reference DOCX and template folder instead of overloading the generic article template.
4. For charts, use the shared plotting helper with the `article` profile so the document gets grayscale figures plus separate TIFF 600 dpi exports.
