QuartoGost
==========

1. Назначение

Проект собирает в одном репозитории шаблоны, ресурсы и Windows-скрипты для
автоматической подготовки электронной документации на базе Quarto с расчётами
и графиками в Julia.

Основой послужили два исходных проекта из каталога `ref`:

- `ref\gostdown`:
  шаблоны и PowerShell-постобработка для документов по ГОСТ 19.xxx и ГОСТ 7.32;
- `ref\Russian-Phd-LaTeX-Dissertation-Template`:
  структура диссертации, автореферата, презентации и конвертов, а также
  нормативные материалы в каталоге `Documents`.

В текущем проекте собраны:

- шаблон текстового документа по ГОСТ 19.xxx (ЕСПД);
- шаблон отчёта о НИР по ГОСТ 7.32-2017;
- шаблон диссертации по ГОСТ Р 7.0.11-2011;
- шаблон автореферата;
- шаблон презентации с экспортом в PPTX;
- сценарий генерации конвертов по TSV-списку рассылки;
- шаблон учебного пособия;
- базовый шаблон статьи как пример расширения проекта;
- Windows-сценарии сборки и инициализации;
- skills для AI-агентов.

Для `espd` и `report` Word-шаблоны титульных и служебных страниц теперь
берутся как прямые копии исходных DOCX из `ref\gostdown`, а не как заново
сгенерированные документы. Это сделано, чтобы сохранить исходные стили
обложек без дрейфа форматирования.

2. Ключевые возможности

- единая структура проекта для разных видов документов;
- хранение исходного текста в `qmd`/Markdown;
- поддержка библиографии в формате BibTeX (`.bib`);
- автоматическая нумерация разделов, таблиц, рисунков, формул;
- перекрёстные ссылки на рисунки, таблицы, формулы и разделы;
- расчёты и построение графиков средствами Julia;
- построение академических графиков через `CairoMakie.jl + MakiePublication.jl`;
- экспорт документов в DOCX;
- экспорт презентаций в PPTX;
- экспорт DOCX в PDF через Microsoft Word COM;
- постобработка Word для сценариев ГОСТ 19.xxx и ГОСТ 7.32 с использованием
  служебных страниц и стилей из `reference.docx`;
- простое добавление новых типов документов, например статей.

3. Структура проекта

.
|-- _quarto.yml
|-- readme.txt
|-- build-*.cmd
|-- resources
|   |-- bibliography
|   |-- csl
|   |-- filters
|   |-- reference-docs
|   |   |-- espd
|   |   |-- report
|   |   |-- dissertation
|   |   `-- synopsis
|   |-- reference-pptx
|   `-- assets
|-- scripts
|   |-- build.ps1
|   |-- new-document.ps1
|   |-- init-julia.ps1
|   |-- check-environment.ps1
|   |-- postprocess-word.ps1
|   `-- julia
|       `-- Project.toml
|-- templates
|   |-- common
|   |   |-- data
|   |   `-- partials
|   |-- espd
|   |-- report
|   |-- dissertation
|   |-- synopsis
|   |-- presentation
|   |-- envelopes
|   |-- study-guide
|   `-- article
|-- skills
|   |-- quarto-gost-workflow
|   |-- quarto-gost-espd
|   |-- quarto-gost-report
|   |-- quarto-gost-dissertation
|   |-- quarto-gost-presentation
|   |-- quarto-gost-envelopes
|   |-- quarto-gost-study-guide
|   `-- quarto-gost-article
`-- ref

4. Требования к окружению под Windows

Обязательные компоненты:

- Windows;
- Quarto;
- Julia;
- Microsoft Word 2010+ с доступным COM API;
- шрифты, применяемые в корпоративных шаблонах, если требуется строгое
  соответствие внешнему виду исходных DOCX.

Желательные компоненты:

- Pandoc в PATH, если используется отдельно от поставки Quarto;
- Microsoft PowerPoint, если впоследствии понадобится дополнительная ручная
  доработка презентаций или создание собственного `reference.pptx`.

Проверка окружения:

`powershell -ExecutionPolicy Bypass -File scripts\check-environment.ps1`

Инициализация Julia-пакетов:

`powershell -ExecutionPolicy Bypass -File scripts\init-julia.ps1`

Выборочное обновление reference DOCX:

`powershell -ExecutionPolicy Bypass -File scripts\update-reference-docs.ps1 -DocumentType report`

Примеры:

- обновить только `espd`:
  `powershell -ExecutionPolicy Bypass -File scripts\update-reference-docs.ps1 -DocumentType espd`
- обновить только `synopsis`:
  `powershell -ExecutionPolicy Bypass -File scripts\update-reference-docs.ps1 -DocumentType synopsis`
- обновить только `dissertation` и `study-guide`:
  `powershell -ExecutionPolicy Bypass -File scripts\update-reference-docs.ps1 -DocumentType dissertation,study-guide`
- обновить только `presentation`:
  `powershell -ExecutionPolicy Bypass -File scripts\update-reference-docs.ps1 -DocumentType presentation`
- обновить всё:
  `powershell -ExecutionPolicy Bypass -File scripts\update-reference-docs.ps1 -DocumentType all`

Обновление reference DOCX из `gostdown` напрямую:

`powershell -ExecutionPolicy Bypass -File scripts\refresh-gostdown-reference-docs.ps1`

Сценарий:

- копирует `ref\gostdown\demo-template-espd.docx` в
  `resources\reference-docs\espd\reference.docx`;
- копирует `ref\gostdown\demo-template-report.docx` в
  `resources\reference-docs\report\reference.docx`;
- заменяет только переменные реквизиты на плейсхолдеры вида `%DOC_TITLE%`,
  `%APPROVER_TITLE%`, `%RESEARCH_CODE%` и т.п., не меняя оформление абзацев;
- не заменяет постоянный текст макета: названия министерств, слова
  `УТВЕРЖДАЮ`, `СОГЛАСОВАНО`, `ЛИСТ УТВЕРЖДЕНИЯ`, служебные подписи,
  подчеркивания и прочие неизменяемые элементы формы.

Выборочная работа с `gostdown`-шаблонами:

- только `espd`:
  `powershell -ExecutionPolicy Bypass -File scripts\refresh-gostdown-reference-docs.ps1 -DocumentType espd`
- только `report`:
  `powershell -ExecutionPolicy Bypass -File scripts\refresh-gostdown-reference-docs.ps1 -DocumentType report`

Выборочная генерация внутренних reference DOCX:

- только `dissertation`:
  `set QUARTOGOST_REFERENCE_TYPES=dissertation && powershell -ExecutionPolicy Bypass -File scripts\generate-reference-docs.ps1`
- только `study-guide`:
  `set QUARTOGOST_REFERENCE_TYPES=study-guide && powershell -ExecutionPolicy Bypass -File scripts\generate-reference-docs.ps1`
- только `synopsis`:
  `set QUARTOGOST_REFERENCE_TYPES=synopsis && powershell -ExecutionPolicy Bypass -File scripts\generate-reference-docs.ps1`
- только `presentation`:
  `set QUARTOGOST_REFERENCE_TYPES=presentation && powershell -ExecutionPolicy Bypass -File scripts\generate-reference-docs.ps1`

Графический стек Julia:

- `CairoMakie.jl` используется как основной backend для публикационных графиков;
- `MakiePublication.jl` задаёт академическую тему оформления;
- для презентаций применяется цветная схема;
- для текстовых документов применяются оттенки серого;
- для статей дополнительно сохраняются TIFF-версии графиков с разрешением 600 dpi.

5. Быстрый старт

1. Проверьте окружение:

   `powershell -ExecutionPolicy Bypass -File scripts\check-environment.ps1`

2. Инициализируйте Julia:

   `powershell -ExecutionPolicy Bypass -File scripts\init-julia.ps1`

3. Запустите один из демонстрационных сценариев:

   `build-espd-demo.cmd`

   `build-report-demo.cmd`

   `build-dissertation-demo.cmd`

   `build-synopsis-demo.cmd`

   `build-presentation-demo.cmd`

   `build-envelopes-demo.cmd`

   `build-study-guide-demo.cmd`

4. Готовые файлы будут созданы в соответствующем подкаталоге `build\...`.

6. Шаблоны и назначение

`templates\espd\espd-template.qmd`

- текстовый документ по ГОСТ 19.xxx / ГОСТ 2.105-95;
- подходит для пояснительных записок, описаний программ, руководств и прочей
  документации, где нужны служебные страницы из DOCX-шаблона.
- содержит пример блока `gost:` для подстановки титульных реквизитов.

`templates\report\report-template.qmd`

- отчёт о НИР по ГОСТ 7.32-2017;
- ориентирован на структуру: реферат, сокращения, введение, основная часть,
  заключение, список источников, приложения.
- содержит пример блока `gost:` для подстановки титульных реквизитов.

`templates\dissertation\dissertation-template.qmd`

- диссертация по ГОСТ Р 7.0.11-2011;
- повторяет академическую логику проекта-источника, но в Quarto/DOCX;
- шаблон и `reference.docx` выровнены по `ref\Russian-Phd-LaTeX-Dissertation-Template\dissertation.tex`,
  `Dissertation\setup.tex` и `Dissertation\title.tex`: титульный лист,
  оглавление, поля `2.5 / 1 / 2 / 2 см`, полуторный интервал.

`templates\synopsis\synopsis-template.qmd`

- автореферат диссертации;
- краткая структура для тиражируемой версии;
- шаблон и `reference.docx` выровнены по `ref\Russian-Phd-LaTeX-Dissertation-Template\synopsis.tex`:
  формат A5, одинарный интервал, титульный лист, оборот и выходные сведения.

`templates\presentation\presentation-template.qmd`

- презентация с экспортом в PPTX;
- подходит для защиты, доклада, отчёта по этапу работ;
- шаблон и `reference.pptx` выровнены по
  `ref\Russian-Phd-LaTeX-Dissertation-Template\presentation.tex`,
  `Presentation\styles.tex` и `Presentation\title.tex`: титульный слайд,
  содержательный layout, синяя академическая схема и широкоформатный кадр;
- вместе с презентацией автоматически формируются раздаточные материалы
  `*-handout.pdf` по логике `presentation_handout.tex`.

`templates\envelopes\envelopes-template.qmd`

- заготовка и справка по генерации конвертов;
- фактическая генерация выполняется через TSV и `scripts\build.ps1`.

`templates\article\article-template.qmd`

- пример того, как проект расширяется на статьи и смежные форматы.

`templates\study-guide\study-guide-template.qmd`

- учебное пособие;
- подходит для учебных и учебно-методических материалов с экспортом в DOCX/PDF;
- учитывает базовые требования к виду издания и выходным сведениям, но локальные издательские правила нужно проверять отдельно.

`resources\reference-docs\study-guide\reference.docx`

- DOCX-шаблон учебного пособия;
- включает титульный лист, оборот титульного листа, маркер содержания `%TOC%`,
  маркер основного текста `%MAINTEXT%` и страницу выходных сведений;
- оформлен по требованиям из `ref\datalab-output-Требования к авторскому оригиналу_pdf.pdf.md`.

7. Нормативные чеклисты

7.1 ГОСТ 2.105-95 и ЕСПД

- разделы нумеруются арабскими цифрами без точки;
- подразделы, пункты и подпункты нумеруются иерархически;
- перечисления оформляют через тире, буквы со скобкой, при детализации — цифры со скобкой;
- при необходимости применяют титульный лист и лист утверждения;
- для текстовых документов рекомендуется лист регистрации изменений;
- при большом объёме документ можно делить на части и книги;
- служебные листы и фирменные реквизиты лучше вести через `reference.docx`.

7.2 ГОСТ 7.32-2017 для отчёта о НИР

- обязательные элементы: титульный лист, список исполнителей, реферат, содержание, введение, основная часть, заключение, список использованных источников, приложения;
- поля: левое 30 мм, правое 15 мм, верхнее 20 мм, нижнее 20 мм;
- межстрочный интервал обычно полуторный, шрифт не менее 12 pt;
- номер страницы ставят внизу по центру; титульный лист учитывается, но номер на нём не печатают;
- рисунки и таблицы размещают после первого упоминания или на следующей странице;
- формулы выносят в отдельную строку, сверху и снизу оставляют по одной свободной строке;
- пояснения под формулой начинают со слова `где` без двоеточия;
- приложения обозначают буквами и включают в содержание.

7.3 ГОСТ Р 7.0.11-2011 для диссертации и автореферата

- диссертацию печатают на одной стороне листа A4 через 1.5 интервала;
- размер шрифта 12-14 pt;
- поля: левое 25 мм, правое 10 мм, верхнее 20 мм, нижнее 20 мм;
- номер страницы ставят в центре верхнего поля;
- главы начинают с новой страницы;
- заголовки центрируют, не ставят точку в конце и не переносят слова;
- автореферат должен включать обложку, общую характеристику работы, основное содержание, заключение и публикации автора;
- положения, выносимые на защиту, должны быть конкретными и совпадать по смыслу в диссертации и автореферате.

7.4 Учебное пособие

- по ГОСТ 7.60-2003 учебное пособие — учебное издание, дополняющее или частично/полностью заменяющее учебник;
- по ГОСТ Р 7.0.4-2020 на титульной странице учебного издания обычно указывают авторов, заглавие, вид издания, сведения о рекомендации или допуске, издание, место, издателя и год;
- локальные требования вуза, НМС, УМО, редакционно-издательского отдела или издательства имеют приоритет для грифа, рецензентов, ISBN, аннотации и выпускных данных.
- по локальному документу `Требования к авторскому оригиналу`:
  поля 1,9/1,9/1,9/2,4 см, нижний колонтитул 1,5 см, Times New Roman, 16 кегль,
  одинарный интервал, выравнивание по ширине, абзац 1-1,5 см, автоматические переносы;
- каждая глава начинается с новой страницы, подразделы продолжаются на текущей;
- в разделе должно быть не менее двух подразделов;
- каждая глава учебного пособия завершается материалом на закрепление;
- заголовки 1 уровня набираются заглавными буквами, заголовки 2+ уровня —
  строчными; рекомендована индексационная нумерация;
- формулы центрируются, важные формулы нумеруются в формате `глава.номер`,
  номер ставят у правого поля;
- таблицы и рисунки размещают после абзаца со ссылкой на них.

7.5 Унификация DOCX-стилей и `custom-style`

Во всех DOCX-шаблонах проекта используются одинаковые имена пользовательских
стилей. Это сделано для того, чтобы одни и те же `custom-style` из Markdown /
Quarto работали одинаково по смыслу во всех типах документов, а различалось
только оформление.

Базовые имена стилей, которые должны существовать во всех `reference.docx`:

- абзацные / итоговые: `UnnumberedHeading1`, `UnnumberedHeading1NoTOC`,
  `UnnumberedHeading2`, `GostKeywords`, `Figure`, `ReferenceItem`,
  `MyCustomStyle`, `Source Code`, `Captioned Figure`, `First Paragraph`;
- табличные: `TableStyleContributors`, `TableStyleAbbreviations`,
  `TableStyleGost`, `TableStyleGostNoHeader`;
- стили-маркеры для `custom-style`: `UnnumberedHeadingOne`,
  `UnnumberedHeadingOneNoTOC`, `UnnumberedHeadingTwo`, `AppendixHeadingOne`,
  `ContributorsTable`, `AbbreviationsTable`, `GostKeywords`, `Figure`,
  `ReferenceItem`, `MyCustomStyle`.

Когда использовать `custom-style`:

- `UnnumberedHeadingOne`
  Ненумерованный заголовок первого уровня, который должен попасть в оглавление.
  Подходит для `Введение`, `Заключение`, `Список использованных источников`,
  `Обозначения и сокращения`.

- `UnnumberedHeadingOneNoTOC`
  Ненумерованный заголовок первого уровня, который не должен попадать в
  оглавление. Подходит для `Реферат`, `Аннотация`, `Содержание`,
  `Список исполнителей`, служебных заголовков.

- `UnnumberedHeadingTwo`
  Ненумерованный подзаголовок второго уровня. Обычно используется внутри
  приложений или служебных разделов.

- `AppendixHeadingOne`
  Заголовок приложения. Используется для первого заголовка приложения и
  позволяет корректно считать количество приложений при постобработке.

- `ContributorsTable`
  Метка для таблицы списка исполнителей. Ставится в первой ячейке таблицы,
  чтобы Word-постобработка применила специальный табличный стиль.

- `AbbreviationsTable`
  Метка для таблицы сокращений и обозначений. Ставится в первой ячейке таблицы.

- `GostKeywords`
  Абзац или блок ключевых слов в реферате/аннотации.

- `Figure`
  Блок для рисунка без подписи и номера. Используется редко, когда нужен
  просто вставной графический объект.

- `ReferenceItem`
  Стиль контейнера библиографических записей при выводе списка литературы.

- `MyCustomStyle`
  Резервный пользовательский стиль. Применяется, когда нужно особое оформление,
  не покрытое базовыми средствами Markdown.

- `Normal`
  Используется для принудительного возврата к обычному абзацному стилю внутри
  сложных таблиц и вложенных списков.

8. Скрипт сборки

Основной сценарий:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 ...`

Параметры:

- `-DocumentType`
  Тип документа. Допустимые значения:
  `espd`, `report`, `dissertation`, `synopsis`, `presentation`,
  `envelopes`, `article`, `study-guide`.

- `-InputFile`
  Путь к входному `qmd`.
  Если параметр не указан, для большинства типов используется шаблон по
  умолчанию из каталога `templates`.

- `-OutputDir`
  Каталог для результатов сборки.

- `-Format`
  `all`, `docx`, `pdf`, `pptx`.
  Для презентации `all` означает `pptx`.
  Для остальных типов `all` означает `docx + pdf`.

- `-Name`
  Базовое имя выходных файлов без расширения.

- `-ReferenceDoc`
  Явный путь к DOCX-шаблону.
  Используется главным образом для `espd` и `report`, где Word-постобработка
  вставляет основной текст в `reference.docx`.

- `-EmbedFonts`
  Просит Word встраивать TrueType-шрифты в итоговый DOCX/PDF.

- `-Counters`
  Включает подстановку счётчиков страниц, рисунков, таблиц, источников,
  приложений. Для `report` обычно нужно включать.

- `-NoWordPostprocess`
  Отключает сборку через Word-шаблон и оставляет только прямой вывод Quarto.
  Полезно для диагностики.

- `-AddressList`
  Используется только для `envelopes`.
  Путь к TSV-файлу со списком рассылки.

- `-SenderName`
  Используется только для `envelopes`.
  Подпись отправителя.

- `-SenderAddress`
  Используется только для `envelopes`.
  Почтовый адрес отправителя.

- `-JuliaProject`
  Путь к Julia project directory.
  По умолчанию используется `scripts\julia`.

- `-Quarto`
  Имя или путь к исполняемому файлу Quarto.

9. Типовые сценарии запуска

ГОСТ 19.xxx:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType espd -InputFile templates\espd\espd-template.qmd -OutputDir build\espd -Name my-espd -EmbedFonts`

ГОСТ 7.32:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType report -InputFile templates\report\report-template.qmd -OutputDir build\report -Name my-report -EmbedFonts -Counters`

Диссертация:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType dissertation -InputFile templates\dissertation\dissertation-template.qmd -OutputDir build\dissertation -Name thesis`

Демо:

`build-dissertation-demo.cmd`

Автореферат:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType synopsis -InputFile templates\synopsis\synopsis-template.qmd -OutputDir build\synopsis -Name synopsis`

Презентация:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType presentation -InputFile templates\presentation\presentation-template.qmd -OutputDir build\presentation -Name defense-slides`

Демо:

`build-presentation-demo.cmd`

Результат:

- `defense-slides.pptx`
- `defense-slides-handout.pdf`

Конверты:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType envelopes -AddressList templates\common\data\sample-addresses.tsv -OutputDir build\envelopes -Name mailing -SenderName "Организация" -SenderAddress "Адрес отправителя"`

Статья:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType article -InputFile templates\article\article-template.qmd -OutputDir build\article -Name paper`

Учебное пособие:

`powershell -ExecutionPolicy Bypass -File scripts\build.ps1 -DocumentType study-guide -InputFile templates\study-guide\study-guide-template.qmd -OutputDir build\study-guide -Name study-guide`

С отдельной обложкой:

- основной шаблон: `resources\reference-docs\study-guide\reference.docx`
- отдельная обложка: `resources\reference-docs\study-guide\cover-template.docx`

10. Создание нового документа из шаблона

Команда:

`powershell -ExecutionPolicy Bypass -File scripts\new-document.ps1 -DocumentType report -Destination work\reports -Name stage-1-report`

Результат:

- будет создан каталог, если его ещё нет;
- шаблон `qmd` будет скопирован в указанную папку;
- дальше вы редактируете новый файл и собираете его через `scripts\build.ps1`.

11. Заполнение титульных полей из Quarto

Для `espd` и `report` плейсхолдеры в `reference.docx` заполняются из блока
`gost:` в YAML front matter файла `qmd`.

Пример:

````
---
title: "Описание программы"
gost:
  approver_title: "Директор Центра цифровой вёрстки"
  approver_name: "И. И. Сидоров"
  doc_title: "СИСТЕМА АВТОМАТИЧЕСКОЙ ПОДГОТОВКИ ЭЛЕКТРОННОЙ ДОКУМЕНТАЦИИ"
  doc_code: "QG.2026-01 13 01"
---
````

Правила:

- имя ключа `gost` в `qmd` преобразуется в плейсхолдер Word по схеме
  `%UPPER_SNAKE_CASE%`;
- например `doc_title` заменяет `%DOC_TITLE%`;
- плейсхолдерами должны быть только меняющиеся реквизиты документа:
  должности, фамилии, коды, обозначения, названия работ, даты, шифры,
  номера и другие переменные поля;
- постоянный текст формы и декоративные элементы `reference.docx` не нужно
  превращать в плейсхолдеры;
- блок `gost:` должен содержать простые строковые значения;
- если ключ не задан, в итоговом DOCX останется исходный плейсхолдер, что
  удобно для нормоконтроля и проверки незаполненных реквизитов.

Для `synopsis` используется отдельный блок `synopsis:` в front matter.

Пример:

````
---
title: "Автореферат диссертации"
synopsis:
  manuscript_status: "На правах рукописи"
  thesis_author: "Фамилия Имя Отчество"
  thesis_title: "Название диссертации"
  specialty_line_1: "Специальность 2.3.1 --- ..."
  thesis_degree: "кандидата технических наук"
  thesis_city: "Москва"
  thesis_year: "2026"
---
````

Ключи из `synopsis:` так же преобразуются в плейсхолдеры Word по схеме
`%UPPER_SNAKE_CASE%`.

12. Как редактировать служебные страницы DOCX

Для типов `espd` и `report` структура сборки двухступенчатая:

1. Quarto формирует временный DOCX из `qmd`;
2. Word COM открывает `reference.docx`, вставляет в него основной текст,
   перестраивает оглавление, обновляет счётчики и сохраняет итоговый DOCX/PDF.

Файлы для редактирования:

- `resources\reference-docs\espd\reference.docx`
- `resources\reference-docs\report\reference.docx`

В них можно править:

- титульные листы;
- лист утверждения;
- лист регистрации изменений;
- фирменные стили организации;
- служебные надписи и поля.

Важно:

- не удаляйте маркер `%MAINTEXT%`, если хотите сохранять автоматическую вставку
  основного текста;
- не удаляйте маркер `%TOC%`, если хотите автоматическое оглавление;
- для отчёта также используются маркеры `%NPAGES%`, `%NTABLES%`,
  `%NFIGURES%`, `%NREFERENCES%`, `%NAPPENDICES%`, `%NCHAPTERS%`.
- реквизиты титульных листов для `espd` и `report` лучше менять не вручную
  в DOCX, а через блок `gost:` в `qmd`.

13. Добавление таблиц

Обычная Markdown-таблица:

| Параметр | Значение |
|---|---|
| alpha | 1.25 |
| beta  | 2.50 |

Таблица из Julia:

````
```{julia}
#| label: tbl-results
#| tbl-cap: "Результаты эксперимента"
using DataFrames
DataFrame(Параметр = ["alpha", "beta"], Значение = [1.25, 2.50])
```
````

Ссылка на таблицу:

`см. @tbl-results`

14. Добавление графиков и рисунков

График, построенный в Julia:

````
```{julia}
#| label: fig-curve
#| fig-cap: "График зависимости"
using CairoMakie
...
```
````

Во всех основных шаблонах уже есть скрытый инициализационный Julia-блок:

- он подключает `scripts\julia\plotting.jl`;
- включает `CairoMakie.jl + MakiePublication.jl`;
- активирует профиль документа.

Правило профилей:

- `presentation`: цветная академическая схема, рендер в PNG для надёжной вставки в PPTX;
- `espd`, `report`, `dissertation`, `synopsis`, `study-guide`: академическая схема в оттенках серого, дополнительный экспорт графиков в `generated-figures\*.svg`;
- `article`: та же серая схема, экспорт в `generated-figures\*.svg` и отдельно в `generated-figures\*.tiff` с разрешением 600 dpi.

Рекомендуемый шаблон графика:

````
```{julia}
#| label: fig-curve
#| fig-cap: "График зависимости"
using CairoMakie

x = range(0, 10, length = 200)
y = sin.(x)

fig = Figure(size = (900, 450))
ax = Axis(fig[1, 1], xlabel = "x", ylabel = "sin(x)")
lines!(ax, x, y, linewidth = 3)
quarto_gost_export_assets(fig, "fig-curve")
fig
```
````

Важно:

- не задавайте цвета вручную без необходимости, чтобы профиль документа мог
  автоматически применить нужную цветовую схему;
- для текстовых документов встраиваемый вариант графика должен быть векторным
  (`svg`);
- для статей TIFF 600 dpi создаётся как дополнительный файл для передачи в
  издательские системы, даже если в DOCX вставляется `svg`.

Ссылка на рисунок:

`см. @fig-curve`

Вставка внешнего рисунка:

`![Подпись рисунка](path/to/image.png){#fig-image width=80%}`

Ссылка:

`см. @fig-image`

15. Добавление формул

Встроенная формула:

`$a^2 + b^2 = c^2$`

Выключная формула:

````
$$
E = mc^2
$$ {#eq-energy}
````

Ссылка на формулу:

`см. @eq-energy`

16. Библиография и ссылки на источники

Файл библиографии по умолчанию:

`resources\bibliography\references.bib`

Вставка ссылки:

- `[@ivanov2023]`
- `[@doe2024, c. 45]`
- `см. [@gostdown]`

Список литературы формируется автоматически при наличии `.bib` и `csl`.

17. Ссылки на разделы

Идентификатор у раздела:

`# Методика {#sec-method}`

Ссылка в тексте:

`см. @sec-method`

18. Работа с include и расширением структуры

Для повторно используемых фрагментов применяйте:

`{{< include ../common/partials/example-elements.qmd >}}`

Рекомендуемый подход:

- держать общие фрагменты в `templates\common\partials`;
- выносить повторяемые данные в `templates\common\data`;
- создавать новый каталог в `templates\...` для каждого нового типа документа;
- по мере роста проекта добавлять новые wrapper-скрипты или профили сборки.

19. Презентации и PPTX

Шаблон презентации:

`templates\presentation\presentation-template.qmd`

Особенности:

- экспорт идёт напрямую в PPTX;
- используется `resources\reference-pptx\reference.pptx`, автоматически
  сформированный по мотивам `presentation.tex`;
- после сборки автоматически создаётся `*-handout.pdf` с двумя слайдами на
  странице A4;
- графики Julia вставляются в слайды через профиль `presentation`;
- для презентаций используется цветная академическая схема
  `CairoMakie.jl + MakiePublication.jl`;
- при необходимости шаблон `reference.pptx` можно доработать в PowerPoint,
  сохранив те же layout-и и общую структуру.

20. Конверты для рассылки

Для конвертов используется TSV-файл без строки заголовков.

Формат строк:

`Индекс<TAB>Город<TAB>Адрес<TAB>Организация<TAB>Адресат`

Пример:

`101000<TAB>г. Москва<TAB>ул. Примерная, д. 1<TAB>ФГБУН ...<TAB>Учёному секретарю`

Сценарий сборки преобразует список адресов в многостраничный DOCX.

21. Примеры

Готовые примеры с заполненными титульными полями и демонстрацией таблиц,
рисунков, формул, графиков и `custom-style`:

- `examples\espd-demo\main.qmd`
- `examples\report-demo\main.qmd`
- `examples\synopsis-demo\main.qmd`

Команды запуска:

- `build-espd-demo.cmd`
- `build-report-demo.cmd`
- `build-synopsis-demo.cmd`

В этих примерах показаны:

- заполнение реквизитов обложки через `gost:`;
- построение графиков Julia;
- таблицы и формулы с кросс-ссылками;
- использование `custom-style="UnnumberedHeadingOne"` и
  `custom-style="MyCustomStyle"`.

22. Skills для AI-агентов

В каталоге `skills` находятся навыки для автоматизированной работы с проектом:

- `quarto-gost-workflow`
  общий навык выбора шаблона, редактирования и сборки;
- `quarto-gost-espd`
  подготовка документации по ГОСТ 19.xxx / ГОСТ 2.105-95;
- `quarto-gost-report`
  подготовка отчёта о НИР по ГОСТ 7.32-2017;
- `quarto-gost-dissertation`
  подготовка диссертации и автореферата;
- `quarto-gost-presentation`
  подготовка презентаций;
- `quarto-gost-envelopes`
  подготовка конвертов для рассылки.
- `quarto-gost-article`
  подготовка статей и других расширяемых текстовых форматов.
- `quarto-gost-study-guide`
  подготовка учебных пособий.

Каждый навык содержит:

- `SKILL.md` с кратким алгоритмом;
- каталог `references` с опорными правилами по структуре и применению.

Для графиков во всех навыках действует единое правило:

- презентации: цветная академическая схема;
- остальные документы: оттенки серого;
- в текстовых документах графики дополнительно сохраняются в `svg`;
- в статьях дополнительно формируются `tiff` 600 dpi.

23. Источники требований

При проектировании структуры учитывались материалы:

- `ref\Russian-Phd-LaTeX-Dissertation-Template\Documents\datalab-output-2021-11gost_7.32-2017.pdf.md`
- `ref\Russian-Phd-LaTeX-Dissertation-Template\Documents\datalab-output-GOST 2.105-95.pdf.md`
- `ref\Russian-Phd-LaTeX-Dissertation-Template\Documents\datalab-output-GOST R 7.0.11-2011.pdf.md`
- `ref\Russian-Phd-LaTeX-Dissertation-Template\Documents\datalab-output-Def_positions.pdf.md`

Дополнительно использованы:

- ГОСТ 7.60-2003 для определения вида издания `учебное пособие`;
- ГОСТ Р 7.0.4-2020 для состава выходных сведений учебного издания.
- `ref\datalab-output-Требования к авторскому оригиналу_pdf.pdf.md`
  для локальных требований к стилям, заголовкам, формулам, таблицам,
  иллюстрациям и структуре учебного пособия.

24. Важные замечания

- в текущем проекте Quarto-часть ориентирована на воспроизводимую сборку и
  расширение под новые типы документов;
- точное визуальное соответствие локальным нормативам организации почти всегда
  требует правки `reference.docx`;
- для `espd` и `report` критична доступность Microsoft Word COM;
- для презентации фирменный шаблон PPTX не включён в репозиторий и должен быть
  добавлен отдельно при необходимости.
