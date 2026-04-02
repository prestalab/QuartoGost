## Core report structure from GOST 7.32

- титульный лист
- список исполнителей
- реферат
- содержание
- термины и определения
- перечень сокращений и обозначений
- введение
- основная часть отчёта
- заключение
- список использованных источников
- приложения

## Practical guidance

- The bold items in the standard are usually mandatory.
- In this repository, title and service pages come from the reference DOCX.
- Keep the abstract concise but quantitative when counts or scope are known.

## Title sheet checklist from GOST 7.32-2017

- ministry, ведомство, or parent structure if applicable
- full and short organization name
- UDC index
- research registration number
- report registration number
- agreement and approval marks with dates and signatures
- document type: report on research work
- R&D title
- report title
- report kind: interim or final
- program or topic code
- book number when there are multiple books
- supervisor details
- place and year

## List of executors

- Include surnames and initials, positions, degrees, titles, signatures, and role in report preparation.
- If the report has only one executor, their details may be moved to the title sheet and the separate executors section may be omitted.

## Abstract requirements

- Include total report volume, number of books, illustrations, tables, sources, and appendices.
- Include 5 to 15 keywords or keyword phrases.
- The abstract text should cover the object of study or development, purpose, methods, results, practical significance, and possible applications.
- If some parts do not apply, keep the sequence but omit the absent items.
- Target size: about 850 printed characters and not more than one typewritten page in the source standard.

## Formatting checkpoints from section 6

- Format: A4 by default; A3 is allowed for large figures or tables.
- Printing: one side of the sheet.
- Default spacing: one-and-a-half spacing.
- For very large final reports, one-line spacing is allowed in the standard.
- Font size: at least 12 pt; Times New Roman is recommended in the extracted standard text.
- Margins: left 30 mm, right 15 mm, top 20 mm, bottom 20 mm.
- Paragraph indent: 1.25 cm.
- Page numbering: Arabic numerals, continuous through the whole report including appendices.
- Title sheet counts in pagination but does not display the number.
- Page number is placed at the bottom center.

## Numbering and references

- Sections use Arabic numerals without a trailing dot.
- Subsections, points, and subpoints are numbered hierarchically.
- Enumerations use a dash; if text refers to items, use lowercase Russian letters with a parenthesis.
- Contents must include introduction, all named sections/subsections/points, conclusion, sources, and appendices with page numbers.
- Figures are placed after the first reference or on the next page and must be referenced in text.
- Figures are numbered either consecutively or within a section; captions are centered below the figure.
- Tables are numbered either consecutively or within a section; appendix tables use appendix-prefixed numbering.
- Formulas are placed on separate lines with at least one free line above and below.
- Explanations of symbols go immediately below a formula and begin with `где` without a colon.
- Formula numbers are placed at the right in parentheses.
- Bibliographic references use square brackets and should map to the numbered source list.

## Appendices

- Appendices are designated by letters: `ПРИЛОЖЕНИЕ А`, `ПРИЛОЖЕНИЕ Б`, ...
- Appendix pages keep continuous pagination with the rest of the report.
- Appendices must appear in the contents with designation, status, and title.
- Figures, tables, formulas, and internal sections in appendices use appendix-prefixed numbering.

## Shared custom styles

- Use the shared contract in `skills\quarto-gost-workflow\references\custom-styles.md`.
- In report-like documents the most common custom styles are:
  `UnnumberedHeadingOne`, `UnnumberedHeadingOneNoTOC`, `AppendixHeadingOne`,
  `UnnumberedHeadingTwo`, `ContributorsTable`, `AbbreviationsTable`,
  `GostKeywords`, and `ReferenceItem`.
