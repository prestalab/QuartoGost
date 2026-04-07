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
