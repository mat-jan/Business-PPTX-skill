# Business-PPTX-Skill

Uniwersalny skill dla Claude (Claude Code / Claude.ai), który buduje profesjonalne, spójne z marką prezentacje biznesowe (`.pptx`) z danych, notatek lub luźnych myśli — na podstawie konfigurowalnego motywu firmowego (`theme.json`), a nie na sztywno wpisanej jednej marce.

## Instalacja

1. Skopiuj repozytorium do **podkatalogu** w katalogu ze skillami Claude, tak aby `SKILL.md` leżał bezpośrednio w nim — np. `git clone https://github.com/mat-jan/Business-PPTX-skill ~/.claude/skills/business-presentation-builder` w Claude Code (nazwa folderu powinna odpowiadać polu `name` w `SKILL.md`). W Claude.ai/Cowork zaimportuj repozytorium jako skill.
2. Gotowe — Claude zacznie automatycznie korzystać z tego skilla, gdy poprosisz o przygotowanie prezentacji biznesowej.

## Jak to działa

- `SKILL.md` — główne instrukcje: kiedy ten skill ma się aktywować i jak krok po kroku budować deck
- `references/theme-schema.md` — pełny opis struktury `theme.json` (kolory, fonty, logo, głos marki)
- `references/deck-structures.md` — sprawdzone szkielety narracyjne (pitch inwestorski, QBR, deck sprzedażowy, aktualizacja projektu, raport do zarządu, prelekcja)
- `references/extract-theme-from-pptx.md` — jak wyciągnąć motyw z już istniejącego firmowego szablonu PowerPoint
- `assets/theme.example.json` — przykładowy, wypełniony plik motywu do skopiowania i edycji (wskazane w nim pliki `example_template.pptx` i `example_logo.png` są tylko przykładowe — podmień je na własne; brakujące pliki powodują ostrzeżenie i render bez logo/szablonu)
- `scripts/apply_theme.py` — funkcje pomocnicze python-pptx, które czytają `theme.json` i renderują slajdy zgodnie z motywem (z mapowaniem `layouts` używa prawdziwych layoutów i placeholderów szablonu)
- `scripts/ingest_template.py` — z firmowego `.pptx` tworzy szkic folderu motywu: `theme.json`, logo, kopię szablonu i mapowanie layoutów

## Własny szablon firmowy

```bash
python scripts/ingest_template.py firma.pptx themes/firma --name "Firma S.A."
```

Skrypt kopiuje szablon do `themes/firma/template.pptx`, wyciąga logo, kolory i fonty z motywu PowerPoint, zgaduje które layouty odpowiadają slajdom tytułowemu / sekcji / treści / wykresowi i zapisuje szkic `theme.json`. Claude pokaże wyciągnięte wartości do potwierdzenia, a potem buduje prezentacje na prawdziwych layoutach Twojego szablonu.

## Pierwsze użycie

Przy pierwszym uruchomieniu bez istniejącego `theme.json`, Claude zapyta o:
- kolory marki (hex),
- fonty nagłówków/treści,
- logo,
- (opcjonalnie) istniejący szablon `.pptx`, z którego można wyciągnąć motyw automatycznie.

Zapisze to jako `theme.json`, który można potem wielokrotnie reużywać — także dla wielu różnych marek/klientów (zobacz sekcję "Multiple brands" w `references/theme-schema.md`).

## Wymagania

- `python-pptx` i `Pillow` (`pip install python-pptx pillow`, najlepiej w wirtualnym środowisku; flaga `--break-system-packages` tylko jako ostateczność w środowiskach typu sandbox)
- Claude z dostępem do wykonywania kodu / plików (Claude Code, Claude.ai z Code Execution, lub Cowork)

## Licencja

MIT — zobacz [LICENSE](LICENSE).