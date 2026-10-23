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
- `scripts/apply_theme.py` — funkcje pomocnicze python-pptx, które czytają `theme.json` i renderują slajdy zgodnie z motywem

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