---
name: quarto-gost-study-guide
description: Use when creating or revising a study guide or educational manual in this repository with Quarto, DOCX/PDF output, and baseline alignment to GOST 7.60-2003 and GOST R 7.0.4-2020.
---

# QuartoGost Study Guide

Use this skill for `учебное пособие`, `учебно-методическое пособие`, course
materials, lecture-based manuals, and similar educational editions.

## Default source

- `templates\study-guide\study-guide-template.qmd`

## Build command

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType study-guide -InputFile <file.qmd> -OutputDir <dir> -Name <name>`

## Rules

1. Follow the baseline structure in [structure.md](references/structure.md).
2. Treat `ГОСТ 7.60-2003` as the definition source for what counts as a study guide.
3. Treat `ГОСТ Р 7.0.4-2020` as the source for title-page and output-details expectations.
4. Also use the local file `ref\datalab-output-Требования к авторскому оригиналу_pdf.pdf.md` for concrete formatting and heading rules.
5. Check local publisher or university rules for grif, reviewers, annotation, ISBN, and publication block.
6. Use Julia chunks for reproducible calculations, examples, charts, and tables.
7. Keep the didactic apparatus explicit: introduction, theory, worked examples, exercises, control questions, references, appendices.
8. Remember that the cover is supplied as a separate DOCX file and the main reference DOCX covers title, verso, contents, body insertion, and output details.
9. If the text uses `custom-style`, keep the shared style names from the project-wide style contract.
10. Use grayscale academic charts for the body text and keep the companion `svg` files when the material may later be reused in articles or LMS content.

Before finalizing, run [norm-control.md](references/norm-control.md).
