Файл `reference.pptx` формируется скриптом
`scripts\generate-presentation-reference-pptx.ps1` по мотивам
`ref\Russian-Phd-LaTeX-Dissertation-Template\presentation.tex`,
`Presentation\styles.tex` и `Presentation\title.tex`.

В шаблоне заложены:

- широкоформатный кадр 16:9;
- синяя академическая цветовая схема;
- титульный, содержательный и секционный layout;
- базовая нижняя служебная линия и номер слайда.

Раздаточные материалы формируются отдельно скриптом
`scripts\export-presentation-handout.ps1` на основе уже собранного `pptx`:
комментарии берутся из того же исходного `qmd`, а миниатюры слайдов
экспортируются из `pptx`; затем формируется `*-handout.docx` с раскладкой
2 слайда на страницу A4.

Если требуется фирменное оформление, доработайте этот файл в Microsoft
PowerPoint и сохраните его по тому же пути:

resources\reference-pptx\reference.pptx

После этого Quarto-шаблон презентации будет использовать обновлённый слайд-мастер
автоматически.
