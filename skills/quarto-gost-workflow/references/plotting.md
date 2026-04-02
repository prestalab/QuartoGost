## Julia plotting contract

- Use `CairoMakie.jl` as the rendering backend for project charts.
- Use `MakiePublication.jl` themes for academic-looking figures.
- Do not hardcode colors unless the task explicitly requires a custom palette.
- Let the document profile control the visual mode.

## Profiles

- `presentation`
  Use a color academic palette and render figures for PPTX insertion.

- `espd`, `report`, `dissertation`, `synopsis`, `study-guide`
  Use grayscale styling and export a companion `svg` file to `generated-figures`.

- `article`
  Use grayscale styling, export `svg` for the document workflow, and also save
  a separate `tiff` file at 600 dpi for journal or proceedings submission.

## Template pattern

Each top-level template should contain a hidden Julia setup chunk that:

1. loads `scripts\julia\plotting.jl`;
2. imports `.QuartoGostPlots`;
3. calls `quarto_gost_activate!(...)` with the right profile.

## Figure authoring rule

- build the figure with Makie objects;
- call `quarto_gost_export_assets(fig, "<stem>")`;
- return `fig` as the last expression in the chunk.
