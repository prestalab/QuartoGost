module QuartoGostPlots

using CairoMakie
using MakiePublication
using FileIO
using PNGFiles
using TiffImages

export quarto_gost_activate!,
       quarto_gost_export_assets,
       quarto_gost_save_article_tiff,
       quarto_gost_save_svg,
       quarto_gost_current_profile

const GRAY_PALETTE = [
    "#111111",
    "#333333",
    "#555555",
    "#777777",
    "#999999",
    "#BBBBBB"
]

const CURRENT_PROFILE = Ref(:text)

function quarto_gost_current_profile()
    return CURRENT_PROFILE[]
end

function text_theme()
    return MakiePublication.theme_aps(colors = GRAY_PALETTE)
end

function article_theme()
    return MakiePublication.theme_aps(colors = GRAY_PALETTE)
end

function presentation_theme()
    return MakiePublication.theme_aps(colors = MakiePublication.seaborn_deep())
end

function quarto_gost_theme(profile::Symbol)
    if profile == :presentation
        return presentation_theme()
    elseif profile == :article
        return article_theme()
    else
        return text_theme()
    end
end

function quarto_gost_activate!(profile::Symbol = :text)
    CURRENT_PROFILE[] = profile
    CairoMakie.activate!(type = profile == :presentation ? "png" : "svg")
    Makie.set_theme!(quarto_gost_theme(profile))
    return profile
end

function ensure_assets_dir(dir::AbstractString = "generated-figures")
    mkpath(dir)
    return dir
end

function quarto_gost_save_svg(fig, stem::AbstractString; dir::AbstractString = "generated-figures")
    dir = ensure_assets_dir(dir)
    path = joinpath(dir, string(stem, ".svg"))
    CairoMakie.save(path, fig; pt_per_unit = 1)
    return path
end

function quarto_gost_save_article_tiff(fig, stem::AbstractString; dir::AbstractString = "generated-figures", dpi::Integer = 600)
    dir = ensure_assets_dir(dir)
    png_path = joinpath(dir, string(stem, ".png"))
    tiff_path = joinpath(dir, string(stem, ".tiff"))

    CairoMakie.save(png_path, fig; px_per_unit = dpi / 72)
    image = FileIO.load(png_path)
    FileIO.save(tiff_path, image)
    rm(png_path; force = true)

    return tiff_path
end

function quarto_gost_export_assets(fig, stem::AbstractString; dir::AbstractString = "generated-figures")
    profile = quarto_gost_current_profile()

    if profile == :presentation
        return (;)
    elseif profile == :article
        svg = quarto_gost_save_svg(fig, stem; dir = dir)
        tiff = quarto_gost_save_article_tiff(fig, stem; dir = dir, dpi = 600)
        return (; svg, tiff)
    else
        svg = quarto_gost_save_svg(fig, stem; dir = dir)
        return (; svg)
    end
end

end
