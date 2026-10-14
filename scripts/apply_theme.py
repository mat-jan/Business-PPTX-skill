"""
apply_theme.py

Helper functions for building on-brand .pptx decks from a theme.json config.
This is a starting point meant to be adapted per deck, not a black-box CLI.

Requires: python-pptx  (pip install python-pptx --break-system-packages)

Typical usage from a generation script:

    import json
    from apply_theme import load_theme, new_presentation, add_title_slide, \
        add_section_slide, add_content_slide, add_chart_slide

    theme = load_theme("theme.json")
    prs = new_presentation(theme)

    add_title_slide(prs, theme, "Q3 Business Review", "Acme Corp — October 2026")
    add_section_slide(prs, theme, "Performance")
    add_content_slide(prs, theme, "Revenue grew 34% YoY, driven by EMEA",
                       bullets=["EMEA up 61%", "APAC flat", "Churn down 2pts"])
    add_chart_slide(prs, theme, "Monthly active users trending up",
                     categories=["Jan","Feb","Mar","Apr"],
                     series={"MAU": [120, 135, 150, 171]})

    prs.save("/mnt/user-data/outputs/q3_review.pptx")
"""

import json
import os
import re
import warnings
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE


def load_theme(path):
    """Load theme.json and resolve its asset paths (template_pptx, logo.path)
    relative to the theme file's own directory — not the process's current
    working directory. Absolute paths are left untouched."""
    theme_dir = os.path.dirname(os.path.abspath(path))

    with open(path, "r", encoding="utf-8") as f:
        theme = json.load(f)

    def _resolve(rel_path):
        if not rel_path or os.path.isabs(rel_path):
            return rel_path
        return os.path.normpath(os.path.join(theme_dir, rel_path))

    if theme.get("template_pptx"):
        theme["template_pptx"] = _resolve(theme["template_pptx"])

    logo = theme.get("logo")
    if logo and logo.get("path"):
        logo["path"] = _resolve(logo["path"])

    _validate_theme(theme, path)
    return theme


_HEX_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _validate_theme(theme, source_path="theme.json"):
    """Fail fast with a clear message rather than mid-render. Checks are
    deliberately non-exhaustive — this catches the mistakes that are both
    common and silent (bad hex, empty chart_sequence, missing files) without
    being a full JSON-schema validator."""
    errors = []

    colors = theme.get("colors", {})
    if not colors:
        errors.append("theme.colors is missing or empty.")
    else:
        for key in ("primary", "secondary", "accent", "background", "text_primary"):
            val = colors.get(key)
            if val is not None and not _HEX_RE.match(val):
                errors.append(f"colors.{key} = {val!r} is not a valid 6-digit hex color (e.g. '#0B3D91').")
        chart_sequence = colors.get("chart_sequence")
        if chart_sequence is not None:
            if not isinstance(chart_sequence, list) or len(chart_sequence) == 0:
                errors.append("colors.chart_sequence must be a non-empty list of hex colors.")
            else:
                for c in chart_sequence:
                    if not _HEX_RE.match(c):
                        errors.append(f"colors.chart_sequence contains an invalid hex color: {c!r}.")

    fonts = theme.get("fonts", {})
    if not fonts.get("heading") or not fonts.get("body"):
        errors.append("fonts.heading and fonts.body should both be set (fall back to a common font like Arial if unsure).")

    # Missing template/logo files are a deliberate graceful-fallback case
    # elsewhere in this module (new_presentation / _add_logo skip them
    # silently), so flag them as a warning — loud enough to notice, but not
    # a hard failure, since "no logo yet" or "no template yet" is valid.
    template_path = theme.get("template_pptx")
    if template_path and not os.path.exists(template_path):
        warnings.warn(
            f"theme.template_pptx points to '{template_path}', which does not "
            f"exist (resolved from {source_path}) — falling back to a blank "
            f"presentation.",
            stacklevel=3,
        )

    logo = theme.get("logo")
    if logo and logo.get("path") and not os.path.exists(logo["path"]):
        warnings.warn(
            f"theme.logo.path points to '{logo['path']}', which does not "
            f"exist (resolved from {source_path}) — no logo will be rendered.",
            stacklevel=3,
        )

    if errors:
        raise ValueError(
            "Invalid theme.json (" + source_path + "):\n- " + "\n- ".join(errors)
        )


def _rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    return RGBColor.from_string(hex_str.upper())


def new_presentation(theme, preserve_template_slides=False):
    """Start from the company's own template if available (preserves real
    master slides/layouts), otherwise start blank.

    preserve_template_slides: if the template_pptx already contains slides
    (sample content, a cover the user wants kept, slides that other shapes
    depend on, etc.), pass True to keep them and simply append new slides
    after them. Defaults to False (strip them) because most template files
    used with this skill are pure "layout libraries" — but this is now an
    explicit, deliberate choice rather than an unconditional wipe."""
    template_path = theme.get("template_pptx")
    if template_path and os.path.exists(template_path):
        prs = Presentation(template_path)
        if not preserve_template_slides:
            xml_slides = prs.slides._sldIdLst
            for s in list(xml_slides):
                xml_slides.remove(s)
        # Preserve the template's own slide geometry (e.g. 4:3 or a custom
        # size) — do NOT override it with the default 16:9 dimensions.
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
    return prs


def _blank_layout(prs):
    """Find the closest thing to a 'blank' layout, rather than assuming it
    lives at index 6 — that's only true for PowerPoint's default template,
    and custom template_pptx files can order/name layouts differently.

    Note even PowerPoint's own default 'Blank' layout carries a few
    inherited (invisible until used) placeholders such as date/footer/slide
    number, so we can't require exactly zero — we prefer a layout literally
    named "blank", and otherwise fall back to whichever layout has the
    fewest placeholders."""
    named_candidates = [
        layout for layout in prs.slide_layouts
        if "blank" in layout.name.strip().lower()
    ]
    if named_candidates:
        return named_candidates[0]

    if len(prs.slide_layouts) == 0:
        raise ValueError(
            "This template_pptx has no slide layouts at all, so there is "
            "nothing to build slides from. Use a different template or "
            "fall back to a blank Presentation()."
        )

    return min(prs.slide_layouts, key=lambda layout: len(layout.placeholders))


def _logo_dimensions(logo_path, max_w):
    """Return (width, height) EMU for the logo, preserving its real aspect
    ratio so non-square logos (most logos) don't get stretched. Falls back
    to a square box if the image's intrinsic size can't be read."""
    try:
        from PIL import Image
        with Image.open(logo_path) as img:
            px_w, px_h = img.size
        if px_w > 0 and px_h > 0:
            height = Emu(int(max_w * px_h / px_w))
            return max_w, height
    except Exception:
        pass
    return max_w, max_w


def _add_logo(slide, prs, theme):
    """Place the logo per theme.logo.placement, sized from its real aspect
    ratio. Shared by every slide type so placement is consistent (title
    slides previously ignored this and always rendered top-left)."""
    logo = theme.get("logo", {})
    logo_path = logo.get("path")
    if not (logo_path and os.path.exists(logo_path)):
        return
    max_w = Inches(logo.get("max_width_in", 0.9))
    logo_w, logo_h = _logo_dimensions(logo_path, max_w)
    placement = logo.get("placement", "bottom-right")
    if placement == "bottom-right":
        left = prs.slide_width - logo_w - Inches(0.3)
        top = prs.slide_height - logo_h - Inches(0.2)
    elif placement == "top-right":
        left = prs.slide_width - logo_w - Inches(0.3)
        top = Inches(0.2)
    elif placement == "bottom-left":
        left, top = Inches(0.3), prs.slide_height - logo_h - Inches(0.2)
    else:  # "top-left" and any unrecognized value fall back here
        left, top = Inches(0.3), Inches(0.2)
    slide.shapes.add_picture(logo_path, left, top, width=logo_w, height=logo_h)


def _add_footer_and_logo(slide, prs, theme, page_number=None):
    rules = theme.get("layout_rules", {})
    colors = theme.get("colors", {})

    if rules.get("footer_text"):
        left, top, width, height = Inches(0.6), prs.slide_height - Inches(0.4), Inches(6), Inches(0.3)
        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.text = rules["footer_text"]
        tf.paragraphs[0].font.size = Pt(theme.get("type_scale", {}).get("footnote_pt", 10))
        tf.paragraphs[0].font.color.rgb = _rgb(colors.get("text_primary", "#808080"))

    if rules.get("show_page_numbers"):
        # Honor the theme's show_page_numbers setting even if the caller
        # didn't pass an explicit page_number — derive it from the slide's
        # actual position in the deck so far.
        resolved_number = page_number
        if resolved_number is None:
            resolved_number = len(prs.slides._sldIdLst)
        left = prs.slide_width - Inches(1)
        top = prs.slide_height - Inches(0.4)
        box = slide.shapes.add_textbox(left, top, Inches(0.6), Inches(0.3))
        tf = box.text_frame
        tf.text = str(resolved_number)
        tf.paragraphs[0].font.size = Pt(theme.get("type_scale", {}).get("footnote_pt", 10))
        tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

    _add_logo(slide, prs, theme)


def add_title_slide(prs, theme, title, subtitle=None):
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})
    type_scale = theme.get("type_scale", {})

    slide = prs.slides.add_slide(_blank_layout(prs))
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb(colors.get("primary", "#0B3D91"))

    box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), prs.slide_width - Inches(1.6), Inches(1.5))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = title
    p = tf.paragraphs[0]
    p.font.size = Pt(type_scale.get("title_slide_pt", 40))
    p.font.bold = True
    p.font.name = fonts.get("heading", "Arial")
    p.font.color.rgb = _rgb(colors.get("text_on_dark", "#FFFFFF"))

    if subtitle:
        sbox = slide.shapes.add_textbox(Inches(0.8), Inches(4.1), prs.slide_width - Inches(1.6), Inches(0.8))
        stf = sbox.text_frame
        stf.text = subtitle
        stf.paragraphs[0].font.size = Pt(18)
        stf.paragraphs[0].font.name = fonts.get("body", "Arial")
        stf.paragraphs[0].font.color.rgb = _rgb(colors.get("text_on_dark", "#FFFFFF"))

    _add_logo(slide, prs, theme)

    return slide


def add_section_slide(prs, theme, section_title):
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})

    slide = prs.slides.add_slide(_blank_layout(prs))
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = _rgb(colors.get("secondary", colors.get("primary", "#0B3D91")))

    box = slide.shapes.add_textbox(Inches(0.8), prs.slide_height / 2 - Inches(0.6), prs.slide_width - Inches(1.6), Inches(1.2))
    tf = box.text_frame
    tf.text = section_title
    p = tf.paragraphs[0]
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.name = fonts.get("heading", "Arial")
    p.font.color.rgb = _rgb(colors.get("text_on_dark", "#FFFFFF"))
    return slide


def add_content_slide(prs, theme, title, bullets=None, page_number=None):
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})
    type_scale = theme.get("type_scale", {})
    margins = theme.get("layout_rules", {}).get("margins_in", {"top": 0.5, "left": 0.6, "right": 0.6})

    slide = prs.slides.add_slide(_blank_layout(prs))
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(colors.get("background", "#FFFFFF"))

    tbox = slide.shapes.add_textbox(Inches(margins.get("left", 0.6)), Inches(margins.get("top", 0.5)),
                                     prs.slide_width - Inches(margins.get("left", 0.6) + margins.get("right", 0.6)),
                                     Inches(1))
    ttf = tbox.text_frame
    ttf.word_wrap = True
    ttf.text = title
    ttf.paragraphs[0].font.size = Pt(type_scale.get("slide_title_pt", 28))
    ttf.paragraphs[0].font.bold = True
    ttf.paragraphs[0].font.name = fonts.get("heading", "Arial")
    ttf.paragraphs[0].font.color.rgb = _rgb(colors.get("primary", "#0B3D91"))

    if bullets:
        if len(bullets) > 6:
            warnings.warn(
                f"add_content_slide('{title}'): {len(bullets)} bullets given, "
                f"only the first 6 will be rendered (the 6x6 guideline is a "
                f"content ceiling, not a silent-truncation rule — split this "
                f"into two slides or trim the content instead).",
                stacklevel=2,
            )
        bbox = slide.shapes.add_textbox(Inches(margins.get("left", 0.6)), Inches(1.7),
                                         prs.slide_width - Inches(margins.get("left", 0.6) + margins.get("right", 0.6)),
                                         Inches(4.5))
        btf = bbox.text_frame
        btf.word_wrap = True
        for i, bullet in enumerate(bullets[:6]):  # enforce 6x6-ish ceiling
            p = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
            p.text = f"•  {bullet}"
            p.font.size = Pt(type_scale.get("body_pt", 16))
            p.font.name = fonts.get("body", "Arial")
            p.font.color.rgb = _rgb(colors.get("text_primary", "#1A1A1A"))
            p.space_after = Pt(10)

    _add_footer_and_logo(slide, prs, theme, page_number=page_number)
    return slide


def add_chart_slide(prs, theme, title, categories, series, chart_type="bar", page_number=None):
    """series: dict of {series_name: [values]} aligned with categories."""
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})
    type_scale = theme.get("type_scale", {})

    slide = prs.slides.add_slide(_blank_layout(prs))
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(colors.get("background", "#FFFFFF"))

    tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.5), prs.slide_width - Inches(1.2), Inches(1))
    ttf = tbox.text_frame
    ttf.word_wrap = True
    ttf.text = title
    ttf.paragraphs[0].font.size = Pt(type_scale.get("slide_title_pt", 28))
    ttf.paragraphs[0].font.bold = True
    ttf.paragraphs[0].font.name = fonts.get("heading", "Arial")
    ttf.paragraphs[0].font.color.rgb = _rgb(colors.get("primary", "#0B3D91"))

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series.items():
        chart_data.add_series(name, values)

    xl_type = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)

    x, y, cx, cy = Inches(0.8), Inches(1.7), prs.slide_width - Inches(1.6), Inches(4.8)
    gframe = slide.shapes.add_chart(xl_type, x, y, cx, cy, chart_data)
    chart = gframe.chart

    chart_sequence = colors.get("chart_sequence", [colors.get("primary", "#0B3D91")])

    if chart_type == "line":
        # A series fill doesn't affect a line's stroke — set the line color
        # (and marker fill, so markers match) per series instead.
        for i, plot_series in enumerate(chart.plots[0].series):
            color = chart_sequence[i % len(chart_sequence)]
            plot_series.format.line.color.rgb = _rgb(color)
            plot_series.format.line.width = Pt(2.25)
            try:
                plot_series.marker.format.fill.solid()
                plot_series.marker.format.fill.fore_color.rgb = _rgb(color)
            except (AttributeError, TypeError):
                pass  # not all chart engines expose marker formatting the same way
    elif chart_type == "pie":
        # A series-level fill would make every slice the same color — color
        # each point (slice) individually from chart_sequence instead.
        plot_series = chart.plots[0].series[0]
        for i, point in enumerate(plot_series.points):
            color = chart_sequence[i % len(chart_sequence)]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = _rgb(color)
    else:
        for i, plot_series in enumerate(chart.plots[0].series):
            color = chart_sequence[i % len(chart_sequence)]
            plot_series.format.fill.solid()
            plot_series.format.fill.fore_color.rgb = _rgb(color)

    _add_footer_and_logo(slide, prs, theme, page_number=page_number)
    return slide