## Dissertation manuscript

- титульный лист
- оглавление
- введение
- основная часть
- заключение
- список сокращений и условных обозначений
- словарь терминов
- список литературы
- список иллюстративного материала
- приложения

## Introduction checklist

- актуальность темы
- степень разработанности темы
- цель и задачи исследования
- научная новизна
- теоретическая и практическая значимость
- методология и методы
- положения, выносимые на защиту
- достоверность и апробация

## Synopsis checklist

- общая характеристика работы
- основное содержание по главам
- заключение
- публикации автора

## Dissertation formatting checkpoints from GOST R 7.0.11-2011

- Main text is divided into chapters and paragraphs, or sections and subsections, numbered with Arabic numerals.
- Each chapter or section starts on a new page.
- Headings are centered, without a period at the end, and without word breaks.
- Headings are separated from the text by three line intervals.
- Dissertation is prepared on A4 sheets, one-sided.
- Line spacing: one-and-a-half.
- Font size: 12-14 pt.
- Margins: left 25 mm, right 10 mm, top 20 mm, bottom 20 mm.
- Paragraph indent: equal throughout the text and equal to five characters.
- All pages, including figures and appendices, are numbered continuously.
- Title page counts as page one but does not display the number.
- Page number is placed in the center of the upper margin.

## Illustrations, tables, formulas

- Illustrations are placed after the first textual reference or on the next page, or in appendices if needed.
- Tables are placed after the first textual reference or on the next page, or in appendices if needed.
- Tables may be numbered continuously or within a chapter/section.
- Formula symbols should follow the relevant national standards.
- Explanations of symbols should be given in the text or directly below the formula.
- Bibliographic references in the dissertation text should follow GOST R 7.0.5.

## Defense statements from local reference `Def_positions`

- Defense statements must be clear, concrete, and reflect the essence of the scientific results.
- Avoid vague wording like "a new method is proposed that improves..."
- Include distinguishing features of the new results and the author's contribution.
- Include both the essence of the result and its scientific and practical significance.
- For methods, algorithms, models, and technologies, state not only the name but also the substantive novelty and the comparison basis.
- Good formulations let a specialist understand novelty and value from the statements alone.
- Recommended lead verbs include `установлено`, `обнаружено`, `доказано`, and similar result-oriented forms.
- Keep the same defense statements in the dissertation introduction and in the synopsis.

## Synopsis structure and fields

- The synopsis includes a cover and the text of the synopsis.
- The text contains general characteristics of the work, main content, and conclusion.
- The cover must include the status `на правах рукописи`, author name, dissertation title, specialty code and name, sought degree, city, and year.
- The reverse side must include organization, supervisor or consultant, opponents, leading organization, defense date and place, library, mailing date, and academic secretary name.
- The list of publications by the author on the dissertation topic should be formatted bibliographically.
- In this Quarto project that list should be generated automatically from `.bib` data, preferably from a dedicated file such as `resources\bibliography\author-works.bib`, with the output placed through `::: {#refs}`.
- Grouped biblatex-style sublists from the LaTeX reference project are not reproduced directly here; if separate grouped lists are required, that is a future CSL/Lua-filter customization task.
- The synopsis is intended for printing/duplication, and output details should follow GOST R 7.0.4 when needed.

## Shared custom styles

- Use the shared contract in `skills\quarto-gost-workflow\references\custom-styles.md`.
- In dissertation-family documents the most common custom styles are:
  `UnnumberedHeadingOne`, `UnnumberedHeadingOneNoTOC`, `UnnumberedHeadingTwo`,
  `AppendixHeadingOne`, and `ReferenceItem`.
