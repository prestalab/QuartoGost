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
слайды экспортируются как изображения и раскладываются по 2 на страницу A4
в файл `*-handout.pdf`.

Если требуется фирменное оформление, доработайте этот файл в Microsoft
PowerPoint и сохраните его по тому же пути:

resources\reference-pptx\reference.pptx

После этого Quarto-шаблон презентации будет использовать обновлённый слайд-мастер
автоматически.
