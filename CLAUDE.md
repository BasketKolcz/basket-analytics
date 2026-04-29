# BasketKołcz — Analize Platform · CLAUDE.md

Plik dokumentacji projektu wczytywany automatycznie przez Claude na początku każdej sesji.

---

## Stack techniczny

- **Backend:** Python / Flask, jednoplikowy `app.py` (~16 900 linii)
- **Baza danych:** PostgreSQL (psycopg2), Docker Compose + Gunicorn `--reload`
- **Frontend:** Bootstrap 5, Chart.js 4.4, vanilla JS inline w f-stringach Pythona
- **UI:** język polski, wszystkie szablony HTML generowane jako f-stringi w app.py
- **Uruchomienie:** `docker compose up` → Flask na `localhost:5000`
- **Podgląd Claude:** Python http.server port 8080 → `.claude/serve.bat`

---

## Struktura pliku app.py (linie orientacyjne)

| Linia | Sekcja |
|-------|--------|
| 1–57 | Importy, konfiguracja Flask, baza danych |
| 58–1349 | Court preview, static files, login/logout |
| 1350–1838 | Globalny CSS programu + funkcja `nav()` |
| 1839–2354 | Strona główna (`/`) — dashboard |
| 2355–3138 | Upload meczów (`/upload`) |
| 3140–3762 | Walidacja, admin routes |
| 3764–4079 | Historia meczów (`/historia`) |
| 4080–5928 | Raport meczu (`/mecz/<id>`) |
| 5942–6757 | Raport trenerski PDF, shooting chart |
| 6758–7595 | Sezon — tabela wyników (`/sezon`) |
| 7596–8066 | Program → Zawodnicy (`/zawodnicy`) |
| 8067–8764 | Program → Zawodnik profil (`/zawodnik/<id>`) |
| 8765–9985 | Ustawienia, drużyny, roster |
| 9986–10716 | Export XLSX/PDF meczu |
| 10717–11799 | Szablony składu, roster import/edit |
| 11800–12184 | Edycja meczu, reparse |
| 12185–13571 | Porównaj zbiorcze / porównaj zawodników |
| 13572–13931 | Portal CSS (`_PORTAL_CSS`), funkcje pomocnicze portalu |
| 13932–14886 | Portal → strona główna (`/portal`) — Mecze + Statystyki |
| 14887–16119 | Portal → Mecz (`/portal/mecz/<id>`) |
| 16120–16293 | Portal login/logout, preview selectors |
| 16294–16888 | Portal → Zawodnik (`/portal/zawodnik/<pid>`) |

---

## Program → Zawodnicy (`/zawodnicy`) — szczegóły

### Kolumny thead (0-based index dla sortZaw)
```
0: Zawodnik  1: MIN  2: G  3: PTS
4: 2PT-M  5: 2PT-A  (brak %)  6: 3PT-M  7: 3PT-A  8: FT-M  9: FT-A
10: ZB-A (oreb)  11: ZB-O (dreb)  12: ZB-S (suma)
13: AST  14: TO  15: STL  16: BLK  17: FD
18: eFG%  19: TS%  20: USG%
```
Kolumna **FIN usunięta**. ZB ma 3 podkolumny: A | O | S.

### JS — kluczowe funkcje
```javascript
setPer(mode)        // 'game'|'36'|'40'|'100' — przelicza per-min stats
setStatMode(mode)   // 'avg'|'sum' — przełącznik Średnie/Sumaryczne
sortZaw(col)        // sortowanie po indeksie komórki (0-based)
```

### Header zawodnicy
```html
<div class="d-flex justify-content-between align-items-center mb-3 flex-wrap gap-2">
  <!-- Lewa: border-left title "Statystyki indywidualne" + subtitle id="per36-desc" -->
  <!-- Prawa: toggle Średnie/Sumaryczne + select#perSelect (Per Mecz/36/40/100) -->
```

### data-* atrybuty na wierszach tbody
`data-n data-min data-pts data-p2m data-p2a data-p3m data-p3a data-ftm data-fta data-oreb data-dreb data-ast data-to data-stl data-blk data-fd`
(brak data-fin)

---

## Portal → Statystyki indywidualne — szczegóły

### Kolumny thead_p (0-based dla sortP)
Identyczna kolejność jak Program → Zawodnicy powyżej.

### JS — kluczowe funkcje portalu
```javascript
setPer_p(mode)      // odpowiednik setPer dla portalu
setStatMode_p(mode) // odpowiednik setStatMode dla portalu
sortP(col)          // sortowanie tabeli portalu
```

### Minuty w portalu
```python
_total_min_pl = _pt_total_pl / 60   # sekundy → minuty dziesiętne
```

---

## Portal → Zawodnik (`/portal/zawodnik/<pid>`) — wykresy

### switchMetric(metric) — obsługiwane metryki
`'PTS'|'AST'|'REB'|'OREB'|'DREB'|'STL'|'BLK'|'TO'|'FIN'|'EFG'|'TS'|'USG'`

### Nazwy wykresów
```javascript
var titles = {
  'PTS':'Punkty', 'AST':'Asysty', 'REB':'Rebounds',
  'OREB':'Off. Reb', 'DREB':'Def. Reb',
  'STL':'Przechwyty', 'BLK':'Bloki', 'TO':'Straty', 'FIN':'Wykończenia',
  'EFG':'eFG%', 'TS':'TS%', 'USG':'USG%'
}
```

### Python — zmienne do wykresów (mecz po meczu)
```python
pts_js / ast_js / reb_js / oreb_js / dreb_js / stl_js / blk_js / to_js / fin_js / efg_js / ts_js / usg_js
pts_season_js / ast_season_js / oreb_season_js / dreb_season_js / ...
```

---

## Program → Zawodnik profil (`/zawodnik/<roster_id>`) — wykresy

Taki sam system `switchMetric` jak portal, te same metryki i nazwy.

---

## Portal → Mecze (`/portal` tab Mecze)

### Styl wierszy (bez kolorowego tła)
```python
badge_bg  = "#e8ecf3"          # neutralny dla W i P
badge_tc  = "#1a6b3c" if is_win else "#8b1a1a"   # kolor litery
row_bg    = "#fff"
row_brd   = "#e8ecf3"
```
Litera W/P ma kolor (zielony/czerwony), tło badge jest szare.

---

## Baza danych — kluczowe tabele

```sql
matches       — id, data_meczu, przeciwnik, wynik_gtk, wynik_opp, sezon,
                team_name_a, team_name_b, file_path, parser_version
match_stats   — match_id, druzyna, kwarta, pts, poss, p2m, p2a, p3m, p3a,
                ftm, fta, br, fd, ast, oreb, dreb, stl, blk, d2m, d2a, przerw
player_stats  — match_id, druzyna, nr, pts, p2m, p2a, p3m, p3a, ftm, fta,
                ast, oreb, dreb, stl, blk, br, fd, finishes
timing_stats  — match_id, druzyna, bucket, kwarta, made2, att2, made3, att3, br, ftm
lineup_stats  — match_id, druzyna, lineup, poss, pts, ...
roster        — id, klub_id, nr, imie, nazwisko, pozycja, aktywny, sezon
```

---

## Format arkusza Excel (szablon meczu)

```
Arkusze: META | KODY | GTK Gliwice | UKS... | LEGENDA
Kolumny (GTK/OPP):
  A: Kwarta   B: Czas(sek)   C: Kod akcji   D: Strefa
  E-I: Skład piątki (wypełnione tylko przy zmianach)
  J: Timeout  K: Kończy akcję  L: Asysta  M: Zbiórka OFF
  N: Zbiórka DEF  O: Przechwyt  P: Blok
```

---

## Wzorce i konwencje

### HTML generowany w Pythonie
Cały HTML jako f-stringi. Żadnych osobnych plików szablonów.

### Responsywność (mobile)
- Program: sidebar → hamburger < 768px, media queries w globalnym CSS (~linia 1558)
- Portal: sidebar → poziomy pasek < 900px, dodatkowe media queries w `_PORTAL_CSS`
- Tabele: `overflow-x:auto` + `min-width` — przewijane poziomo na mobile

### Pomocnicze funkcje Python
```python
_avg(v)     # dzieli przez n (liczbę meczów) — używane w tbody zawodników
_a(v)       # odpowiednik dla portalu
f1(v)       # format 1 miejsce po przecinku
```

### Portal CSS
Współdzielony blok `_PORTAL_CSS` (~linia 13565) używany przez:
- `/portal` (strona główna)
- `/portal/mecz/<id>`
- `/portal/zawodnik/<pid>`

---

## Znane ograniczenia / TODO

- `app.py` to monolit ~17 000 linii — edytuj precyzyjnie z offsetem
- Przy dodawaniu/usuwaniu kolumn zawsze przelicz indeksy `sortZaw(N)` i `sortP(N)`
- Portal nie ma systemu uprawnień per-drużyna — jeden login dla wszystkich
- Czas gry (szac.) zaimplementowany w `/zawodnicy` i `/zawodnik`, NIE w portalu
