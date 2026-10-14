# Roadmap

Status aktualny: `apply_theme.py` ma działający silnik renderujący i przeszedł dwie rundy code review (Copilot + ręczny review), z 10 zgłoszonych problemów naprawione zostały 6 bezpośrednio w kodzie. Poniżej lista tego, co zostało — w przybliżonej kolejności wartości/wysiłku.

## P0 — największa wartość, największy zakres

### Render-to-PNG/PDF + wizualna kontrola jakości
Obecnie "review pass" w `SKILL.md` to tylko lista pytań kontrolnych dla Claude ("czy slajd ma >6 bulletów" itd.) — nie ma żadnej automatycznej weryfikacji, czy tekst faktycznie mieści się na slajdzie, czy nic się nie nakłada.

Plan:
- Funkcja `render_deck_preview(pptx_path, out_dir)` w nowym `scripts/render_preview.py`, oparta o LibreOffice headless (`soffice --headless --convert-to pdf`) + `pdf2image` do rozbicia na PNG per slajd
- Prosty detektor overflow tekstu: sprawdzenie czy wysokość renderowanego text frame przekracza wysokość shape'a (python-pptx nie daje tego wprost — trzeba albo estymować po długości tekstu i foncie, albo faktycznie renderować i mierzyć)
- Dodanie tego jako obowiązkowy krok w pipeline opisanym w `SKILL.md`, nie opcjonalny

### Pełny pipeline: outline → generation → render → QA → revision → final
Rozszerzenie obecnego "Step 6 — Review pass" w `SKILL.md` o konkretne, wykonywalne kroki (nie tylko checklistę), wiążące się z powyższym punktem.

## P1 — podnosi jakość wykresów i configu

### Pełne stylowanie wykresów
`add_chart_slide` dziś ustawia tylko kolory serii. Brakuje:
- kontroli osi (min/max, format liczb, widoczność)
- stylu legendy (pozycja, font, włącz/wyłącz)
- etykiet danych na słupkach/punktach
- gridlines (kolor, czy w ogóle pokazywać)
- fontu i rozmiaru tekstu wewnątrz wykresu (dziś dziedziczy domyślny PowerPoint)

Efekt: dziś wykres może nadal "wyglądać jak domyślny PowerPoint" poza samym kolorem serii.

### Rozszerzenie walidacji theme.json
`_validate_theme()` dodana w ostatnim PR łapie złe hex-y i puste `chart_sequence`. Do dodania:
- walidacja że fonty są stringami niepustymi (już jest częściowo)
- ostrzeżenie gdy `type_scale` ma nielogiczne wartości (np. `body_pt > slide_title_pt`)
- walidacja `margins_in` (czy nie sumują się do więcej niż szerokość slajdu)

## P2 — nowe możliwości contentowe

### Speaker notes, appendix, źródła danych
Brakujące elementy typowe dla decków biznesowych:
- `add_speaker_notes(slide, text)` — prosty helper, już częściowo wspominany w `SKILL.md` jako "opcjonalnie", ale nieobsługiwany w kodzie
- Sekcja "Appendix" jako osobny tryb w `deck-structures.md` (wspomniana w review #10 jako "opcjonalny element")
- Przypisy/źródła danych na slajdach z wykresami (mały tekst w rogu: "Źródło: ...")

### Import danych z Excel/CSV
`scripts/import_data.py` z funkcją mapującą arkusz na `categories`/`series` zgodnie z tabelą z `SKILL.md` (Step 4 — Map data to slide types), żeby nie trzeba było ręcznie przepisywać liczb do Pythona.

### `build_from_outline()`
Funkcja przyjmująca listę/dict tematów (zob. przykład w dyskusji) i budująca cały deck jednym wywołaniem — przydatne przy dużych, wielotematycznych deckach.

## P3 — dostępność i jakość kodu

### Accessibility
- Minimalna kontrast-checker funkcja (WCAG AA) dla kombinacji `text_primary`/`background` z theme.json — ostrzeżenie jeśli kontrast za niski
- Minimalny rozmiar tekstu (np. ostrzeżenie jeśli `body_pt < 14` na slajdzie, który może być rzutowany)
- Zasada "kolor nie jest jedynym nośnikiem informacji" — np. przy wykresach z wieloma seriami, dodanie też wzorów/etykiet, nie tylko koloru

### Refaktor `apply_theme.py` na moduły
Obecnie jeden plik ~350 linii. Docelowy podział:
```
scripts/
├── theme_loader.py      # load_theme, _validate_theme
├── slides_core.py        # new_presentation, _blank_layout, add_title_slide, add_section_slide, add_content_slide
├── charts.py              # add_chart_slide i stylowanie wykresów
├── branding.py            # _add_logo, _add_footer_and_logo, _logo_dimensions
└── apply_theme.py         # cienki plik łączący powyższe, zachowuje dotychczasowe importy
```
Zrobić dopiero gdy funkcjonalność się ustabilizuje — przedwczesny refaktor przy wciąż zmieniającym się kodzie tylko utrudnia review.

### Testy automatyczne
`tests/test_apply_theme.py` (pytest) pokrywający:
- każdy z 10 punktów z code review jako osobny test regresyjny
- walidację theme.json (dobry/zły przypadek)
- generowanie pełnego decku end-to-end bez wyjątków

Priorytet wysoki mimo że w kategorii P3 — bez testów każdy kolejny PR ryzykuje cichym zepsuciem wcześniejszych napraw.

## Drobne porządki (dowolny moment)
- Docstring w `apply_theme.py` używa `/mnt/user-data/outputs/` jako przykładowej ścieżki zapisu — to artefakt środowiska, w którym plik powstał; zmienić na neutralny placeholder (np. `"output/deck.pptx"`), żeby nie sugerować konkretnego środowiska
- Doprecyzować w `SKILL.md`, że automatyczny zapis `theme.json` po ekstrakcji z istniejącego pliku `.pptx` następuje DOPIERO po potwierdzeniu wartości przez użytkownika (dziś kolejność jest nieco niejasna między sekcją ekstrakcji a głównym flow)
- `references/deck-structures.md`: dodać wariant "reading deck / async deck" jako osobną strukturę (dziś jest tylko wzmianka o gęstszym tekście w notatce na dole pliku)