---
name: quarto-gost-dissertation
description: Use when creating or revising a dissertation or dissertation synopsis in this repository according to GOST R 7.0.11-2011 and the academic structure inherited from the reference dissertation project.
---

# QuartoGost Dissertation

Use this skill for dissertation manuscripts and dissertation abstracts.

## Default sources

- Dissertation: `templates\dissertation\dissertation-template.qmd`
- Synopsis: `templates\synopsis\synopsis-template.qmd`

## Build commands

Dissertation:
`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType dissertation -InputFile <file.qmd> -OutputDir <dir> -Name <name>`

Synopsis:
`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType synopsis -InputFile <file.qmd> -OutputDir <dir> -Name <name>`

## Rules

1. Use `quarto-gost-scientific-editor` for general text editing, bibliography calls, captions, and Quarto mechanics; this skill is only for dissertation- and synopsis-specific academic constraints.
2. Follow the structure in [academic-structure.md](references/academic-structure.md).
3. Keep the introduction aligned with the required academic elements from GOST R 7.0.11-2011.
4. Present defense statements clearly and concretely.
5. Use appendices for supporting evidence, acts, listings, and supplementary tables.
6. For the synopsis, check both the cover and reverse-side mandatory fields before finalizing.
7. For the synopsis, fill cover and reverse-side fields through the `synopsis:` block in front matter; the project maps those values into the A5 reference DOCX derived from `ref\Russian-Phd-LaTeX-Dissertation-Template\synopsis.tex`.
8. For the dissertation, fill title-page fields through the `dissertation:` block in front matter; the project maps those values into the A4 reference DOCX derived from `ref\Russian-Phd-LaTeX-Dissertation-Template\dissertation.tex`, `Dissertation\setup.tex`, and `Dissertation\title.tex`.

Before finalizing, run [norm-control.md](references/norm-control.md).
