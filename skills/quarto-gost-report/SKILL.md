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

1. Use `quarto-gost-scientific-editor` for general text editing, references, captions, and Quarto structure; this skill is only for GOST 7.32-specific report requirements.
2. Preserve the report reference DOCX markers used for counters.
3. Include the mandatory structural elements from GOST 7.32-2017: title sheet, list of executors, abstract, contents, introduction, main part, conclusion, sources, appendices.
4. Add terms, definitions, abbreviations, and designation lists when the subject requires them.
5. Keep the abstract quantitative and compliant with the standard's sequence.
6. Fill title-sheet requisites through the `gost:` block in front matter so `%REPORT_TITLE%`, `%TOPIC_TITLE%`, `%RESEARCH_CODE%`, `%LEADER_NAME%` and related placeholders are populated automatically.
7. Use placeholders only for changing requisites such as organization, report title, topic, names, positions, dates, registration numbers, and codes; keep static template text and service labels as literal text in the reference DOCX.

Read [required-sections.md](references/required-sections.md) before drafting.
Before finalizing, run [norm-control.md](references/norm-control.md).
