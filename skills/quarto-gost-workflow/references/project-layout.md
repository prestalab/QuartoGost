## Quick map

- `templates\...`: source templates by document type
- `templates\common\partials`: reusable fragments
- `resources\reference-docs`: Word templates for final DOCX assembly
- `scripts\build.ps1`: main build entry point
- `scripts\new-document.ps1`: copy a starter template
- `scripts\init-julia.ps1`: install Julia packages
- `scripts\check-environment.ps1`: verify Quarto, Julia, Pandoc, Word COM
- `scripts\generate-reference-docs.ps1`: refresh common DOCX reference templates and shared style names
- `skills\...`: project-specific skills for agents
- `ref\Russian-Phd-LaTeX-Dissertation-Template\Documents\...`: extracted normative texts used to align report and dissertation-oriented skills

## Selection guide

- `espd`: technical documentation in the GOST 19.xxx / GOST 2.105 family
- `report`: research report under GOST 7.32-2017
- `dissertation`: full dissertation manuscript
- `synopsis`: dissertation abstract
- `presentation`: PPTX slides
- `envelopes`: mailing envelopes from a TSV list
- `article`: extensibility example for papers
- `study-guide`: educational manuals and study guides

## DOCX style contract

- All reference DOCX files should expose the same custom style names.
- Formatting may differ by document family, but style names should remain stable.
- See `references/custom-styles.md` in the workflow skill for the shared mapping.

## Normative sources inside `ref`

- `datalab-output-GOST 2.105-95.pdf.md`: section and list numbering, title and approval sheet logic, change-registration guidance
- `datalab-output-2021-11gost_7.32-2017.pdf.md`: report structure, abstract requirements, page layout, numbering, appendices
- `datalab-output-GOST R 7.0.11-2011.pdf.md`: dissertation and synopsis structure, layout, page numbering, references
- `datalab-output-Def_positions.pdf.md`: drafting rules for defense statements
