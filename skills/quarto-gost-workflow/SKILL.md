---
name: quarto-gost-workflow
description: Use when working in this repository to choose a QuartoGost document template, fill it with content, run the Windows build scripts, and keep output aligned with the project's structure and automation flow.
---

# QuartoGost Workflow

Use this skill when the task is to create, update, or build documentation inside
this repository.

## Workflow

1. Read [project-layout.md](references/project-layout.md) if you need the repo map.
2. When the task is norm-sensitive, align the draft with the extracted standard notes:
   GOST 2.105-95 for ESPD-style text documents,
   GOST 7.32-2017 for research reports,
   GOST R 7.0.11-2011 for dissertations and synopses.
3. Choose the document type:
   `espd`, `report`, `dissertation`, `synopsis`, `presentation`, `envelopes`, `article`, `study-guide`.
4. Start from the matching file in `templates\...` or run
   `scripts\new-document.ps1`.
5. Keep reusable content in `templates\common\partials` and shared data in
   `templates\common\data`.
6. For `espd` and `report`, remember that title and service pages live in
   `resources\reference-docs\...\reference.docx`.
7. Build with `scripts\build.ps1`.
8. If the request includes calculations or charts, prefer Julia code blocks.
9. Before finalizing, check that mandatory structural elements from the standard are present.
10. If the draft uses `custom-style`, verify that the target `reference.docx` contains the same style names.
11. For charts, follow the shared plotting contract in [plotting.md](references/plotting.md).
12. For `espd` and `report`, prefer filling title and approval data through the `gost:` block in QMD front matter rather than editing the DOCX cover by hand.
13. When a `reference.docx` must be refreshed, update only the document family you are working on by using `scripts\update-reference-docs.ps1 -DocumentType <type>` instead of regenerating all reference templates.
14. When the task is specifically editorial, structural, or wording-related inside `qmd`, use `quarto-gost-scientific-editor` as the common text-editing layer and then apply the document-family skill only for format-specific constraints.

## Build rules

- Use `scripts\check-environment.ps1` before diagnosing missing tools.
- Use `scripts\init-julia.ps1` when Julia packages are not installed.
- Use `-NoWordPostprocess` only for debugging raw Quarto output.
- Use `-Counters` for reports where template counters must be updated.

## Output discipline

- Keep section names close to the relevant GOST structure.
- Use cross-reference labels for figures, tables, equations, and sections.
- Put bibliography entries in `.bib` instead of formatting the list manually.
- Treat `ref\Russian-Phd-LaTeX-Dissertation-Template\Documents\*.md` as the local normative extraction source for skills work in this repo.
- Use the shared custom-style contract from [custom-styles.md](references/custom-styles.md) so style names stay stable across document families.
