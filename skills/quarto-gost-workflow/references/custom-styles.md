## Shared `custom-style` contract

The repository keeps the same style names across DOCX templates. This lets the
same Markdown and Quarto sources work across multiple document profiles while
only the visual design changes.

### Final paragraph and table styles expected in reference DOCX files

- `UnnumberedHeading1`
- `UnnumberedHeading1NoTOC`
- `UnnumberedHeading2`
- `GostKeywords`
- `Figure`
- `ReferenceItem`
- `MyCustomStyle`
- `Source Code`
- `Captioned Figure`
- `First Paragraph`
- `TableStyleContributors`
- `TableStyleAbbreviations`
- `TableStyleGost`
- `TableStyleGostNoHeader`

### Marker styles used from Quarto / Pandoc custom-style attributes

- `UnnumberedHeadingOne`
- `UnnumberedHeadingOneNoTOC`
- `UnnumberedHeadingTwo`
- `AppendixHeadingOne`
- `ContributorsTable`
- `AbbreviationsTable`
- `GostKeywords`
- `Figure`
- `ReferenceItem`
- `MyCustomStyle`

### Usage guide

- `UnnumberedHeadingOne`: unnumbered level-1 heading that should still appear in the contents.
- `UnnumberedHeadingOneNoTOC`: unnumbered level-1 heading that should stay out of the contents.
- `UnnumberedHeadingTwo`: unnumbered level-2 heading.
- `AppendixHeadingOne`: first heading of an appendix.
- `ContributorsTable`: marker for the executors table.
- `AbbreviationsTable`: marker for the abbreviations table.
- `GostKeywords`: keyword block in an abstract or annotation.
- `Figure`: wrapper for an unnumbered figure.
- `ReferenceItem`: bibliography container or bibliography paragraph style.
- `MyCustomStyle`: fallback user-defined style for special formatting.
- `Normal`: use this built-in style when you need to reset paragraph formatting inside complex tables or list-heavy cells.
