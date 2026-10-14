# theme.json schema

This file captures everything needed to render on-brand decks for one company. One `theme.json` per company/brand. Store it next to the output decks or in a shared `themes/` folder if the user manages multiple brands (e.g. an agency building decks for several clients).

```json
{
  "company_name": "Acme Corp",
  "template_pptx": "assets/acme_template.pptx",
  "logo": {
    "path": "assets/acme_logo.png",
    "placement": "bottom-right",
    "max_width_in": 0.9
  },
  "colors": {
    "primary": "#0B3D91",
    "secondary": "#1B998B",
    "accent": "#F4A300",
    "background": "#FFFFFF",
    "background_dark": "#0B1220",
    "text_primary": "#1A1A1A",
    "text_on_dark": "#FFFFFF",
    "chart_sequence": ["#0B3D91", "#1B998B", "#F4A300", "#C14953", "#6C5B7B", "#355C7D"]
  },
  "fonts": {
    "heading": "Montserrat",
    "body": "Source Sans Pro",
    "heading_fallback": "Arial",
    "body_fallback": "Arial"
  },
  "type_scale": {
    "title_slide_pt": 40,
    "slide_title_pt": 28,
    "body_pt": 16,
    "footnote_pt": 10
  },
  "layout_rules": {
    "footer_text": "Acme Corp — Confidential",
    "show_page_numbers": true,
    "margins_in": { "top": 0.5, "bottom": 0.5, "left": 0.6, "right": 0.6 }
  },
  "voice": {
    "tone": "confident, concise, data-driven, no corporate jargon",
    "language": "pl",
    "avoid": ["buzzwords like 'synergy' or 'leverage'", "exclamation marks"]
  }
}
```

## Field notes

- **template_pptx**: if present, ALWAYS start new decks from this file's master slides/layouts rather than building layouts from scratch in python-pptx. This is the single biggest driver of a deck actually looking "real" instead of generic. If absent, build one once from the other fields and offer to save it back into the theme for next time.
- **colors.chart_sequence**: use this instead of python-pptx/PowerPoint default chart colors. Order matters — first color for the primary series.
- **fonts**: PowerPoint will substitute a fallback if the exact font isn't installed on the viewer's machine; always set `*_fallback` to something common (Arial, Calibri) so the deck doesn't break on another computer.
- **voice.language**: generate all slide text (titles, bullets, chart labels) in this language unless the user explicitly asks for another one for a specific deck.
- **layout_rules.footer_text**: rendered on every content slide (not usually on title/section dividers).

## Multiple brands / clients

If the skill is being used by an agency or consultant across several clients, keep a `themes/` directory:
```
themes/
├── acme/theme.json
├── acme/assets/...
├── globex/theme.json
└── globex/assets/...
```
Ask which theme to use if more than one exists and it's ambiguous from context.
