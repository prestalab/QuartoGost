---
name: quarto-gost-scientific-editor
description: Edit, review, and tighten Russian scientific or academic text inside this QuartoGost repository. Use when the task is to improve structure, headings, terminology, captions, references, bibliography, cross-references, notes, or editorial consistency in `.qmd` sources for reports, dissertations, synopses, study guides, presentations, and related academic documents.
---

# QuartoGost Scientific Editor

Use this skill for editorial work on source `qmd` files in this project.

## Core rule

Edit the Quarto source, not the generated `docx` or `pptx`, unless the task is
explicitly about reference templates or output formatting.

## Workflow

1. Identify the document family: `report`, `dissertation`, `synopsis`, `study-guide`, `presentation`, `espd`, or `article`.
2. Read [checklist.md](references/checklist.md) for the compact editorial pass.
3. Read [quarto-conventions.md](references/quarto-conventions.md) before changing headings, figures, tables, citations, notes, or `custom-style`.
4. Load the document-family-specific references only when needed:
   - `skills\quarto-gost-report\references\required-sections.md`
   - `skills\quarto-gost-report\references\norm-control.md`
   - `skills\quarto-gost-dissertation\references\academic-structure.md`
   - `skills\quarto-gost-dissertation\references\norm-control.md`
   - `skills\quarto-gost-espd\references\required-sections.md`
   - `skills\quarto-gost-espd\references\norm-control.md`
   - `skills\quarto-gost-study-guide\references\structure.md`
   - `skills\quarto-gost-study-guide\references\norm-control.md`
5. Keep title data, front matter blocks, and placeholder keys stable unless the task is to change them deliberately.

## Editing priorities

1. Fix broken structure first: wrong heading levels, missing sections, duplicated headings, inconsistent appendices.
2. Fix scientific-document mechanics next: captions, cross-references, citations, bibliography calls, notes, abbreviations, and terminology.
3. Tighten language after structure is correct: remove repetition, improve transitions, shorten bulky sentences, and align style across sections.

## Review priorities

When asked to review, prioritize:

1. Structural mismatches against the target document family.
2. Broken Quarto mechanics: missing labels, dead references, malformed citations, wrong `custom-style`, or notes placed outside slide scope.
3. Editorial inconsistencies: abbreviations, capitalization, numbering, caption wording, appendix naming, and bibliography style drift.

## Output discipline

- Preserve Russian academic tone unless the user asks for simplification.
- Prefer concrete rewrites over abstract advice when the text is local and editable.
- Keep comments for presentation handouts in source QMD:
  - `notes:` in front matter for the title slide
  - `::: {.notes}` inside normal slides
- Do not invent bibliography entries; use existing `.bib` keys or flag missing data.
