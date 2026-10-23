# Extracting a theme.json from an existing company .pptx

If the user has an existing branded deck or template, prefer reverse-engineering `theme.json` from it over asking them to re-specify colors/fonts from scratch — it's more accurate and less work for them.

## Steps

1. Load the `.pptx` with python-pptx:
   ```python
   from pptx import Presentation
   prs = Presentation("existing_deck.pptx")
   ```
2. **Colors**: python-pptx doesn't expose theme colors, so read the theme XML from the pptx zip archive (usually `ppt/theme/theme1.xml`; the exact part is the one the slide master's relationships point to). Its `<a:clrScheme>` block lists dk1/lt1/dk2/lt2/accent1–6 as hex values. Propose accent1 → primary, accent2 → secondary, accent3 → accent (dk2 is often a brand color too), then confirm with the user rather than guessing blindly.
3. **Fonts**: the same theme XML has `<a:fontScheme>`; read the `typeface` attribute of `<a:latin>` inside `<a:majorFont>` (headings) and `<a:minorFont>` (body).
4. **Logo**: look for a picture shape on the slide master/layouts that repeats across slides — extract it from the pptx's `ppt/media/` folder and save it to `assets/`.
5. **Layouts**: keep the original file as `template_pptx` in theme.json rather than rebuilding its master and slide size manually. `apply_theme.py` starts from this file and adds new slides on its Blank layout (python-pptx cannot duplicate existing slides).
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
    if scheme is None:
        raise ValueError('No <a:clrScheme> found in the theme XML')
    colors = {}
    for child in scheme:
        tag = child.tag.split('}')[1]
        srgb = child.find('a:srgbClr', ns)
        sys_clr = child.find('a:sysClr', ns)
        val = srgb.get('val') if srgb is not None else (sys_clr.get('lastClr') if sys_clr is not None else None)
        colors[tag] = f"#{val}" if val else None
    return colors
```

Treat the output as a draft: confirm it with the user (step 6) before saving `theme.json`.
