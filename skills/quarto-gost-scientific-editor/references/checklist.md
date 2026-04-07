## Compact checklist

Use this list for a fast editorial pass over a `qmd` document.

### 1. Structure

- Heading levels are sequential and not skipped.
- Mandatory sections for the document family are present.
- Appendix headings and unnumbered sections are used consistently.

### 2. Quarto mechanics

- Figures, tables, equations, and sections have stable labels where needed.
- Cross-references point to existing labels.
- Citations use `.bib` keys and consistent citation syntax.
- `custom-style` names match the shared contract from the project.

### 3. Editorial consistency

- Terms and abbreviations are introduced once and reused consistently.
- Captions for tables and figures follow one naming pattern.
- Numbering, lists, and appendix names do not drift across chapters.
- Capitalization of institutions, positions, and headings is uniform.

### 4. Document-family specifics

- `report`: abstract, abbreviations, conclusion, sources, appendices.
- `dissertation`: introduction elements, chapter logic, conclusion, references, appendices.
- `synopsis`: concise defense-oriented structure and reverse-side metadata alignment.
- `study-guide`: pedagogical structure, annotations, review/approval blocks if required.
- `presentation`: one main point per slide, short bullets, notes available for handout.

### 5. Final pass

- Remove placeholder wording that accidentally leaked into narrative text.
- Check that front matter values still match the document content.
- Keep generated-output concerns out of the prose unless the section is explicitly methodological.
