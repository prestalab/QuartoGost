---
name: quarto-gost-report
description: Use when creating or revising a research report in this repository under GOST 7.32-2017 with mandatory report structure, bibliography, counters, and DOCX/PDF output.
---

# QuartoGost Report

Use this skill for reports on research, design, or engineering work in the
GOST 7.32 structure.

## Default source

- `templates\report\report-template.qmd`

## Build command

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType report -InputFile <file.qmd> -OutputDir <dir> -Name <name> -EmbedFonts -Counters`

## Rules

1. Preserve the report reference DOCX markers used for counters.
2. Include the mandatory structural elements from GOST 7.32-2017: title sheet, list of executors, abstract, contents, introduction, main part, conclusion, sources, appendices.
3. Add terms, definitions, abbreviations, and designation lists when the subject requires them.
4. Use `.bib` for sources and cite them inline.
5. Prefer reproducible Julia-generated figures and tables over manual screenshots.
6. Keep the abstract quantitative and compliant with the standard's sequence.
7. Keep numbering, page placement, figure/table/formula captions, and appendix designations aligned with the GOST rules below.
8. Use the grayscale plotting profile for charts so figures remain suitable for DOCX/PDF and are also exported to `svg`.
9. Fill title-sheet requisites through the `gost:` block in front matter so `%REPORT_TITLE%`, `%TOPIC_TITLE%`, `%RESEARCH_CODE%`, `%LEADER_NAME%` and related placeholders are populated automatically.
10. Use placeholders only for changing requisites such as organization, report title, topic, names, positions, dates, registration numbers, and codes; keep static template text and service labels as literal text in the reference DOCX.

Read [required-sections.md](references/required-sections.md) before drafting.
Before finalizing, run [norm-control.md](references/norm-control.md).
