## Quarto editing conventions

### Headings

- Use `#`, `##`, `###` according to the logical structure, not visual size.
- For unnumbered sections, use Quarto attributes or the document's established pattern.

### Cross-references

- Figures: `@fig-...`
- Tables: `@tbl-...`
- Equations: `@eq-...`
- Sections: `@sec-...` when the project uses section labels explicitly

Do not mention references in prose if the target object has no label.

### Bibliography

- Use `[@key]`, `[@key1; @key2]`, or the local style already present in the file.
- Keep bibliography data in `.bib`, not as manually typed final entries.
- In this project bibliography formatting is produced automatically by Quarto/Pandoc citeproc with a CSL file, not by `biblatex`, `biber`, or `.bst`.
- Use `::: {#refs}` to place the generated bibliography exactly where the document needs the final list.
- For synopsis author-publication lists, prefer a dedicated `.bib` file such as `resources\bibliography\author-works.bib` and avoid mixing those entries with unrelated external citations in the same document.
- Treat `utf8gost71u.bst` and `ugost2008mod.bst` from the LaTeX reference project as formatting references only; direct reuse in Quarto is not supported.

### Figures and tables

- Keep captions informative and short.
- Prefer source-generated graphics and tables over pasted screenshots.
- For presentation slides, do not overload captions; keep the detail in `::: {.notes}`.

### Notes for presentation handouts

- Title slide notes come from front matter:

```yaml
notes: |
  Comment for the title slide.
```

- Normal slide notes come from slide-local blocks:

```markdown
::: {.notes}
Comment for this slide.
:::
```

### Custom styles

Only use style names that exist in the project contract, for example:

- `UnnumberedHeadingOne`
- `UnnumberedHeadingOneNoTOC`
- `UnnumberedHeadingTwo`
- `Figure`
- `ReferenceItem`
- `MyCustomStyle`

Do not invent new style names unless the task includes updating the corresponding
reference templates.
