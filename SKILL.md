---
name: business-presentation-builder
description: Build professional, on-brand business presentations (.pptx) from raw data, notes, or bullet points — using a company's own theme (colors, fonts, logo, slide layouts) defined in a config file. Use this skill whenever the user asks to create a business presentation, pitch deck, quarterly report, board deck, sales deck, or any PowerPoint deliverable for a company, AND especially when the user mentions "firmowy motyw", "nasz branding", "company theme", "brand guidelines", or wants consistent, reusable corporate styling across presentations — even if they don't explicitly say "use the skill." Also use this skill to set up a NEW company theme config the first time a user wants to brand their presentations. Do not use for one-off casual slides with no branding requirement, or for non-PowerPoint documents (use the docx/pdf skill instead).
---

# Business Presentation Builder

A universal, brand-aware skill for generating business presentations as `.pptx` files. It separates **content** (what the user wants to say) from **theme** (how the company's decks should look), so the same skill works for any company — the only thing that changes is a small config file.

This skill builds on top of the `pptx` skill for the low-level mechanics of creating/editing PowerPoint files (python-pptx). Read that skill first if it's available in this environment; this skill adds the branding layer, data-to-slide mapping, and narrative structure on top of it.

## How this skill works, in one sentence

Read the company's `theme.json` (or create one if missing) → turn the user's content/data into a slide outline following proven deck narrative patterns → render each slide using python-pptx, applying the theme's colors/fonts/logo/layouts consistently → save and present the `.pptx`.

## Step 1 — Get or create the theme

Look for a theme file the user points to, or ask where it lives. Common locations to check first: a file named `theme.json`, `brand.json`, or similar in the working directory or an uploaded file.

- **If a theme file exists**: load it. Its schema is documented in `references/theme-schema.md` (copy `assets/theme.example.json` as a starting reference for the fields).
- **If the user provides a company template (`.pptx`/`.potx`)** and no theme file: run `python scripts/ingest_template.py <template> themes/<company>` to generate a draft `theme.json` (colors, fonts, logo, layout mapping), show the extracted values and layout mapping to the user, apply their corrections, then continue. With a layout mapping the deck is built from the template's own layouts and placeholders.
- **If no theme file exists**: this is a branding session, not just a deck. Ask the user (briefly, batched into one question) for:
  1. Primary + secondary brand color (hex codes) and an accent color
  2. Font for headings and font for body text (must be a font available in PowerPoint/Office, or embeddable)
  3. Logo file (PNG/SVG with transparent background, ideally)
  4. Any existing template `.pptx` they already use (if they have one, prefer extracting the theme FROM it rather than building from scratch — see `references/extract-theme-from-pptx.md`)

  Then generate a `theme.json` per the schema and save it alongside the output so it can be reused next time. Tell the user explicitly, in their own language, that the theme was saved to `theme.json` and can simply be referenced next time without re-entering these details.

Never invent brand colors/fonts/logo on your own if the user hasn't provided them and no theme file exists — ask first. Generic "corporate blue" defaults are a last resort only if the user explicitly says they don't care about branding for this one-off deck.

## Step 2 — Understand the content and audience

Before building slides, clarify (ask only what's missing, don't re-ask what's already given):

- **Purpose**: internal update, investor pitch, sales deck, board report, conference talk?
- **Audience**: executives (terse, high-level, decision-focused) vs. technical team (more detail) vs. external clients (more polish, less internal jargon)
- **Source material**: pasted text/bullets, a spreadsheet/CSV of data, an existing document to condense, or raw ideas to structure
- **Length**: rough slide count or time limit (1 min/slide is a reasonable rule of thumb for spoken presentations)

## Step 3 — Build the narrative outline first, as text

Do NOT jump straight to building slides. First draft a slide-by-slide outline (title + 1-line purpose per slide) and show it to the user as a plain list, unless they've explicitly said to skip review. This is cheap to correct in text; it's expensive to correct after rendering 20 slides.

Use `references/deck-structures.md` for proven narrative patterns (investor pitch, QBR, project update, sales deck, etc.) — pick the one that matches the purpose from Step 2, or blend them.

General slide-writing rules regardless of structure:
- One idea per slide. If a slide needs "and" in its title, it's probably two slides.
- Titles are a claim/insight, not a topic label. Prefer "Revenue grew 34% YoY driven by EMEA" over "Revenue Overview."
- 6×6 guideline as a ceiling, not a target: max ~6 bullets, ~6 words each. If there's more to say, that's speaker notes, not slide text.
- Every data slide should have ONE visual takeaway, not a raw data dump. Pull the insight into the title.
- End with a clear next step / ask / decision needed — decks that just "present information" without a call to action are weaker.

## Step 4 — Map data to slide types

If the user supplied data (CSV, spreadsheet, numbers in chat), map it to appropriate visual slide types rather than pasting tables:

| Data shape | Slide type |
|---|---|
| Trend over time | Line chart |
| Comparison across categories (≤7) | Bar chart |
| Part-to-whole | Stacked bar or simple donut (avoid 3D, avoid >5 slices) |
| Single important number | Big number / "stat" slide |
| Process/sequence | Numbered step or horizontal flow diagram |
| Before/after or options | Two/three-column comparison |
| Timeline/roadmap | Horizontal timeline |
| Org/structure | Simple box hierarchy |

Keep chart styling theme-consistent (use the theme's chart color sequence from `theme.json`, not PowerPoint's default blue/orange/gray palette).

## Step 5 — Render the deck

Use python-pptx (see the `pptx` skill for setup details and gotchas) with `scripts/apply_theme.py` as a starting point (add `scripts/` to `sys.path` to `import apply_theme`) — it reads `theme.json` and gives you helper functions for: title slide, section divider, content slide, chart slide, and a consistent footer/logo placement across all slides. Adapt/extend these helpers rather than writing slide XML from scratch.

Concretely:
1. Start from the theme's template `.pptx` if one exists (preserves the real master, theme and slide size — far more reliable than reconstructing them from scratch). If no template exists, the helpers fall back to a blank 16:9 presentation styled from the theme fields.
2. If `theme.json` has a `layouts` mapping, each slide is created from the mapped template layout and its title/body placeholders are filled (fonts, logo and footer come from the template master). Without a mapping, the helpers add slides on the template's "Blank" layout and draw titles, bullets, charts, footer and logo as themed shapes.
3. Keep consistent margins, logo placement (usually bottom-right or footer), and page numbers per the theme config.
4. For charts, generate them natively in pptx (editable) rather than as pasted images, unless the theme config specifies otherwise.

## Step 6 — Review pass before delivering

Before presenting the file, do a self-check:
- Does every slide follow the theme's color/font rules (no stray default PowerPoint styling)?
- Any slide with >6 bullet points or dense paragraphs that should be split or trimmed?
- Logo and page numbers present and consistent across slides?
- Title slide and closing/next-steps slide present?
- Speaker notes added for slides that need verbal context (optional, ask if wanted)?

Then save to the environment's output location (`/mnt/user-data/outputs/` in Claude.ai, otherwise the working directory or a path the user gives), and present the file to the user along with the theme.json if it was newly created.

## Reference files

- `references/theme-schema.md` — full schema for `theme.json`, with field explanations
- `references/deck-structures.md` — narrative templates (investor pitch, QBR, sales deck, project update, board report, conference talk)
- `references/extract-theme-from-pptx.md` — how to reverse-engineer a theme.json from an existing company .pptx template
- `assets/theme.example.json` — a filled-out example theme to copy and edit
- `scripts/apply_theme.py` — python-pptx helper functions keyed off theme.json (title slide, section slide, content slide, chart slide, footer/logo)
- `scripts/ingest_template.py` — builds a draft theme folder (`theme.json`, logo, template copy, layout mapping) from a company `.pptx`

## Common mistakes to avoid

- Building slides before agreeing on the outline (expensive to fix)
- Treating this as "one deck" when the user actually wants a reusable theme — always offer to save theme.json
- Dumping raw data tables onto slides instead of converting to a visual + takeaway
- Ignoring the audience (a board deck and a team standup update should look and read very differently even with the same data)
- Overfilling slides because "more information feels more thorough" — push detail to speaker notes or an appendix section instead
