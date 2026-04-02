---
name: quarto-gost-presentation
description: Use when creating or revising a presentation in this repository with Quarto to PPTX export, concise slide structure, and Julia-generated visuals.
---

# QuartoGost Presentation

Use this skill for defense slides, stage reports, conference decks, and internal
presentations derived from the same source materials as the written documents.

## Default source

- `templates\presentation\presentation-template.qmd`

## Build command

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType presentation -InputFile <file.qmd> -OutputDir <dir> -Name <name>`

## Rules

1. Keep one main point per slide.
2. Prefer figures, schemes, and short evidence-heavy bullets over dense text.
3. Use Julia plots for reproducible charts.
4. For charts, use the shared plotting helper so presentation figures stay in the color academic scheme.
5. Start from `resources\reference-pptx\reference.pptx`, which is derived from `ref\Russian-Phd-LaTeX-Dissertation-Template\presentation.tex`, `Presentation\styles.tex`, and `Presentation\title.tex`.
6. Keep the title slide metadata aligned with the dissertation: title, organization, city/year, presenter, and supervisor should match the written materials.
7. If a corporate or university PPTX master is required, adapt `resources\reference-pptx\reference.pptx` rather than styling each slide manually.
8. Put speaker comments for the handout into the source QMD: use front matter `notes:` for the title slide and `::: {.notes}` blocks inside individual slides.
9. The presentation build also produces `*-handout.docx`, a two-slides-per-page A4 handout with comments derived from the same QMD in the spirit of `presentation_handout.tex`.

Read [slides.md](references/slides.md) when you need slide guidance.
