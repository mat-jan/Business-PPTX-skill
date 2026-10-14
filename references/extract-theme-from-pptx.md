# Extracting a theme.json from an existing company .pptx

If the user has an existing branded deck or template, prefer reverse-engineering `theme.json` from it over asking them to re-specify colors/fonts from scratch — it's more accurate and less work for them.

## Steps

1. Load the `.pptx` with python-pptx:
   ```python
   from pptx import Presentation
   prs = Presentation("existing_deck.pptx")
   ```
2. **Colors**: inspect `prs.slide_masters[0]` and its `theme` XML (python-pptx doesn't expose theme colors directly — read `ppt/theme/theme1.xml` inside the pptx zip archive for the `<a:clrScheme>` block, which lists dk1/lt1/dk2/lt2/accent1-6 as hex values). Map accent1 → primary, accent2 → secondary, etc., then confirm with the user rather than guessing blindly.
3. **Fonts**: same theme XML has `<a:fontScheme>` with majorFont (headings) and minorFont (body).
4. **Logo**: check slide masters/layouts for a picture placeholder that repeats across slides — extract it via the pptx's `ppt/media/` folder and save it to `assets/`.
5. **Layouts**: keep the original file as `template_pptx` in theme.json rather than trying to rebuild its layouts manually — duplicating slides from it is far more reliable than recreating them.
6. Always show the user the extracted values and ask them to confirm/correct before saving `theme.json` — automated extraction can pick up incidental colors that aren't actually brand colors.

## Quick extraction snippet

```python
import zipfile
from xml.etree import ElementTree as ET

def extract_theme_colors(pptx_path):
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    with zipfile.ZipFile(pptx_path) as z:
        xml = z.read('ppt/theme/theme1.xml')
    root = ET.fromstring(xml)
    scheme = root.find('.//a:clrScheme', ns)
    colors = {}
    for child in scheme:
        tag = child.tag.split('}')[1]
        srgb = child.find('a:srgbClr', ns)
        sys = child.find('a:sysClr', ns)
        val = srgb.get('val') if srgb is not None else (sys.get('lastClr') if sys is not None else None)
        colors[tag] = f"#{val}" if val else None
    return colors
```

Treat the output as a draft, not a final answer — confirm with the user.
