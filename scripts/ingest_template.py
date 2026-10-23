"""
ingest_template.py

Turn a company PowerPoint template (.pptx / .potx) into a ready-to-use theme
folder:

    themes/<name>/
    ├── template.pptx     copy of the uploaded template
    ├── assets/logo.*     first picture found on the slide master / layouts
    └── theme.json        colors, fonts, layout mapping (a DRAFT to confirm)

Usage:
    python ingest_template.py path/to/company.pptx themes/acme [--name "Acme Corp"]

Requires: python-pptx
The result is a draft: show the printed summary to the user and let them
correct colors, fonts and the layout mapping before relying on it.
"""

import argparse
import json
import os
import shutil
import sys
from xml.etree import ElementTree as ET

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

_NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}


def read_theme_xml(prs):
    """Return the ElementTree root of the theme the first slide master uses."""
    theme_part = prs.slide_masters[0].part.part_related_by(RT.THEME)
    return ET.fromstring(theme_part.blob)


def extract_colors(root):
    """Map the <a:clrScheme> slots (dk1, accent1, ...) to '#RRGGBB'."""
    scheme = root.find(".//a:clrScheme", _NS)
    if scheme is None:
        raise ValueError("No <a:clrScheme> found in the template's theme XML.")
    out = {}
    for child in scheme:
        tag = child.tag.split("}")[1]
        srgb = child.find("a:srgbClr", _NS)
        sys_clr = child.find("a:sysClr", _NS)
        val = srgb.get("val") if srgb is not None else (
            sys_clr.get("lastClr") if sys_clr is not None else None)
        if val:
            out[tag] = "#" + val.upper()
    return out


def extract_fonts(root):
    """Return (heading, body) typefaces from <a:fontScheme>."""
    def face(kind):
        el = root.find(f".//a:fontScheme/a:{kind}/a:latin", _NS)
        return el.get("typeface") if el is not None else None
    return face("majorFont"), face("minorFont")


def extract_logo(prs, assets_dir):
    """Save the first picture on the master or a layout (logos live there).
    Returns the file name, or None if no picture was found."""
    sources = [prs.slide_masters[0]] + list(prs.slide_layouts)
    for src in sources:
        for shape in src.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = shape.image
                os.makedirs(assets_dir, exist_ok=True)
                name = f"logo.{image.ext}"
                with open(os.path.join(assets_dir, name), "wb") as f:
                    f.write(image.blob)
                return name
    return None


def guess_layouts(prs):
    """Guess which layout to use for each slide kind from names/placeholders.
    Returns only the kinds it could find."""
    def types(layout):
        return {ph.placeholder_format.type for ph in layout.placeholders}

    title_t = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
    body_t = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT}
    layouts = list(prs.slide_layouts)

    def find(pred):
        for layout in layouts:
            if pred(layout, layout.name.strip().lower(), types(layout)):
                return layout.name
        return None

    guesses = {
        "title": find(lambda l, n, t: n in ("title slide", "title", "cover")
                      or PP_PLACEHOLDER.CENTER_TITLE in t)
                 or find(lambda l, n, t: "title" in n and "only" not in n and "content" not in n),
        "section": find(lambda l, n, t: "section" in n or "divider" in n),
        "content": find(lambda l, n, t: n in ("title and content", "content")
                        or (bool(t & title_t) and bool(t & body_t)
                            and "section" not in n and "title slide" not in n)),
        "chart": find(lambda l, n, t: n == "title only" or ("only" in n and "title" in n))
                 or find(lambda l, n, t: n in ("title and content", "content")),
    }
    return {k: v for k, v in guesses.items() if v}


def build_theme(prs, name, template_file, logo_file):
    root = read_theme_xml(prs)
    c = extract_colors(root)
    heading, body = extract_fonts(root)
    accents = [c[k] for k in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6") if k in c]

    theme = {
        "company_name": name,
        "template_pptx": template_file,
        "colors": {
            "primary": c.get("accent1", "#0B3D91"),
            "secondary": c.get("accent2", "#1B998B"),
            "accent": c.get("accent3", "#F4A300"),
            "background": c.get("lt1", "#FFFFFF"),
            "background_dark": c.get("dk2", "#0B1220"),
            "text_primary": c.get("dk1", "#1A1A1A"),
            "text_on_dark": c.get("lt1", "#FFFFFF"),
            "chart_sequence": accents or [c.get("accent1", "#0B3D91")],
        },
        "fonts": {
            "heading": heading or "Arial",
            "body": body or "Arial",
            "heading_fallback": "Arial",
            "body_fallback": "Arial",
        },
        "type_scale": {"title_slide_pt": 40, "slide_title_pt": 28, "body_pt": 16, "footnote_pt": 10},
        "layout_rules": {
            "footer_text": "",
            "show_page_numbers": False,
            "margins_in": {"top": 0.5, "bottom": 0.5, "left": 0.6, "right": 0.6},
        },
        "voice": {"tone": "confident, concise, data-driven", "language": "en", "avoid": []},
    }
    layouts = guess_layouts(prs)
    if layouts:
        theme["layouts"] = layouts
    if logo_file:
        theme["logo"] = {"path": logo_file, "placement": "bottom-right", "max_width_in": 0.9}
    return theme


def ingest(pptx_path, out_dir, name=None):
    prs = Presentation(pptx_path)
    os.makedirs(out_dir, exist_ok=True)
    template_file = "template.pptx"
    # .potx is not openable by python-pptx as a deck on save, so only the
    # extracted values matter there; always store the copy as template.pptx.
    shutil.copyfile(pptx_path, os.path.join(out_dir, template_file))

    logo = extract_logo(prs, os.path.join(out_dir, "assets"))
    logo_rel = f"assets/{logo}" if logo else None
    company = name or os.path.splitext(os.path.basename(out_dir.rstrip("/\\")))[0].replace("-", " ").title()
    theme = build_theme(prs, company, template_file, logo_rel)

    theme_path = os.path.join(out_dir, "theme.json")
    with open(theme_path, "w", encoding="utf-8") as f:
        json.dump(theme, f, indent=2, ensure_ascii=False)
        f.write("\n")
    return theme_path, theme, [l.name for l in prs.slide_layouts]


def main(argv=None):
    ap = argparse.ArgumentParser(description="Create theme.json from a company .pptx template.")
    ap.add_argument("pptx", help="path to the company template (.pptx)")
    ap.add_argument("out_dir", help="theme folder to create, e.g. themes/acme")
    ap.add_argument("--name", help="company name (default: derived from the folder name)")
    args = ap.parse_args(argv)

    theme_path, theme, layout_names = ingest(args.pptx, args.out_dir, args.name)
    print(f"Wrote {theme_path}  (DRAFT — confirm with the user)")
    print("Colors :", json.dumps(theme["colors"]))
    print("Fonts  :", theme["fonts"]["heading"], "/", theme["fonts"]["body"])
    print("Logo   :", theme.get("logo", {}).get("path", "none found on master/layouts"))
    print("Layouts available:", layout_names)
    print("Layout mapping   :", theme.get("layouts", "none guessed — set theme.layouts by hand"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
