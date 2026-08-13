"""
Platforma Edukacyjna — Seed curriculum wiekowe wg Żalgirisu Kowno
=================================================================
Uruchomienie (z katalogu projektu lub z Dockera):
    python knowledge_base/ak_curriculum_seed.py

Lub przez Docker:
    docker exec -i 5basketkolcz-analizeplatform-web-1 python /app/knowledge_base/ak_curriculum_seed.py

Tabele:
    player_lessons       — lekcje (title, category, content, position_focus=wiek, order_num, duration_min)
    player_quiz_questions — pytania ABCD (lesson_id, question, options_json, correct_idx, explanation)

Konwencja position_focus:
    'ALL'  — istniejące lekcje (ogólne, domyślnie U14)
    'U8'   — kategoria U6–U8   (fundament ruchowy)
    'U10'  — kategoria U9–U10  (słabsza ręka + taktyczne podstawy)
    'U13'  — kategoria U11–U13 (zaawansowane umiejętności + podstawy teamowe)
    'U15'  — kategoria U14–U15 (P&R + 5x5 + statystyki)
    'U18'  — kategoria U16–U18 (strefa + presing + zaawansowane)

order_num:
    1–99   — istniejące lekcje (ALL)
    100–199 — U8
    200–299 — U10
    300–399 — U13
    400–499 — U15
    500–599 — U18
"""

import json
import os
import sys

try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    print("Brak psycopg2. Instaluj: pip install psycopg2-binary")
    sys.exit(1)

DB_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/basketkolcz"
)

# ──────────────────────────────────────────────────────────────────────────────
# DANE CURRICULUM
# Format: (title, category, duration_min, position_focus, order_num, content, quizzes)
# quizzes: lista (question, [opt0, opt1, opt2, opt3], correct_idx_0based, explanation)
# ──────────────────────────────────────────────────────────────────────────────

CURRICULUM = [

    # ══════════════════════════════════════════════════════════
    # U6–U8 | Fundament ruchowy. Zero taktyki zespołowej.
    # Żalgiris: bieg, zmiana kierunku, ball handling, step/jump stop,
    #           pivot, jab step, layup MOCNĄ ręką, podstawy rzutu.
    # ══════════════════════════════════════════════════════════

    (
        "Jak trzymać piłkę i kozłować",
        "taktyka", 4, "U8", 101,
        """Podstawa koszykówki: dobry chwyt i kozłowanie.

Chwyt piłki:
• Trzymaj piłkę opuszkami palców, NIE całą dłonią
• Kciuki skierowane ku sobie, pozostałe palce rozłożone
• Między piłką a dłonią masz „okienko" — widać światło

Kozłowanie:
• Uderzasz piłkę jedną ręką, nie chwytasz jej
• Patrz NA BOISKO i kolegów — NIE na piłkę
• Kolana lekko ugięte, ciało pochylone do przodu

W kategorii U6–U8 ćwiczymy TYLKO mocną ręką.
Słabsza ręka dołącza w U9-U10 — taka jest progresja Żalgirisu!""",
        [
            ("Którą ręką najpierw uczymy się kozłować?",
             ["Słabszą", "Mocniejszą", "Obiema naraz", "To nie ma znaczenia"],
             1,
             "W kategorii U6–U8 ćwiczycie TYLKO mocną ręką. Słabsza dołącza dopiero w U9–U10 — "
             "tak zaprojektowano cały plan Żalgirisu. Dwie ręce naraz byłoby zbyt trudne!"),
            ("Podczas kozłowania powinieneś patrzeć na...",
             ["Piłkę", "Podłogę", "Boisko i kolegów", "Swoje buty"],
             2,
             "Dobry koszykarz NIE patrzy na piłkę — 'czuje' ją przez palce. "
             "Patrzy na boisko, żeby widzieć kolegów i rywali. Trenuj bez patrzenia na piłkę!"),
            ("Jak trzymasz piłkę przy chwycie?",
             ["Całą dłonią jak przy ściskaniu", "Opuszkami palców z 'okienkiem' pod piłką",
              "Tylko kciukiem i palcem wskazującym", "Między oba łokcie"],
             1,
             "Piłka leży na OPUSZKACH palców — nie na całej dłoni. "
             "Między piłką a środkiem dłoni jest 'okienko' (widać między nimi światło). "
             "Taki chwyt daje ci kontrolę i siłę przy rzucie!"),
        ]
    ),

    (
        "Zatrzymanie — step stop i jump stop",
        "taktyka", 4, "U8", 102,
        """Po kozłowaniu musisz się zatrzymać — i zrobić to prawidłowo!

JUMP STOP (zatrzymanie jednoczesne):
• Skaczysz ilądujesz na DWA stopy jednocześnie
• Teraz masz dwie stopy „obrotowe" — możesz użyć KAŻDEJ jako pivot

STEP STOP (zatrzymanie krokami):
• Najpierw jedna noga, potem druga
• Pierwsza noga ląduje = ta jest Twoją stopą obrotową (pivot foot)

Po zatrzymaniu możesz:
✓ Obrócić się (pivot)
✓ Podać piłkę
✓ Rzucić do kosza
✗ NIE możesz już kozłować ani biec!

Krok za krokiem z Żalgirisu: step stop i jump stop ćwiczycie już od U6–U8.""",
        [
            ("Jump stop to lądowanie na...",
             ["Jedną nogę", "Dwie nogi jednocześnie", "Kolana", "Palce u stóp"],
             1,
             "Jump STOP = skaczysz ilądujesz na OBA stopy RÓWNOCZEŚNIE. "
             "To super zatrzymanie, bo możesz wybrać dowolną stopę jako obrotową!"),
            ("Po zatrzymaniu (stop) możesz dalej...",
             ["Kozłować i biec", "Tylko stać w miejscu", "Obrócić się, podać lub rzucić", "Zrobić trzy kroki"],
             2,
             "Stop = koniec ruchu z piłką! Możesz pivot (obrót), podać lub rzucić. "
             "Ale NIE możesz znowu kozłować — to byłby błąd (double dribble)!"),
            ("Przy step stop, która noga jest Twoją stopą obrotową?",
             ["Prawa, zawsze", "Lewa, zawsze", "Ta, która ląduje jako pierwsza",
              "Ta, która ląduje jako ostatnia"],
             2,
             "Przy step stop: noga która ląduje PIERWSZA = Twoja stopa obrotowa (pivot foot). "
             "Nie możesz jej podnosić zanim nie zwolnisz piłki!"),
        ]
    ),

    (
        "Rzut layup — mocna ręka",
        "taktyka", 5, "U8", 103,
        """Layup to najważniejszy rzut w koszykówce. Ćwiczymy go już od U6–U8!

Kolejność kroków (prawa ręka):
1. Kozłujesz w stronę kosza
2. Łapiesz piłkę lewą — prawą jednocześnie (lub tylko dwie ręce)
3. Krok LEWĄ nogą (długi krok)
4. Krok PRAWĄ nogą (krótki, odbijasz się)
5. Piłka leci W TABLICĘ — w kwadrat powyżej obręczy
6. Łagodne lądowanie na dwie nogi

Dla lewej ręki: wszystko odwrotnie!

Cel rzutu: w kwadrat na tablicy — piłka się od niego odbija.
Żalgiris: w U6–U8 ćwiczymy layup TYLKO mocną ręką.""",
        [
            ("Przy layupie PRAWĄ ręką, odbijasz się z...",
             ["Prawej nogi", "Lewej nogi", "Obu nóg jednocześnie", "To nie ma znaczenia"],
             1,
             "Praworęczny layup: krok lewą → ODBIJ z prawej → lewa ręka zbiera, prawa rzuca! "
             "Dla leworęcznego jest odwrotnie. Ta zasada dotyczy każdego layupu."),
            ("W co celujesz przy layupie?",
             ["Prosto w obręcz od góry", "W kwadrat na tablicy", "W krawędź obręczy",
              "W sznurki kosza"],
             1,
             "Celuj W KWADRAT namalowany na tablicy! Piłka odbija się od niego wprost do kosza. "
             "To dlatego layup jest tak skuteczny — używasz tablicy jako pomocy."),
            ("Ile kroków możesz zrobić po ostatnim kozłowaniu przy layupie?",
             ["Jeden krok", "Dwa kroki", "Trzy kroki", "Dowolnie dużo"],
             1,
             "FIBA zezwala na DWA KROKI po ostatnim kozłowaniu (lub po złapaniu piłki). "
             "Trzeci krok = błąd kroków! Dotyczy to każdego zawodnika, od U6 do NBA."),
            ("Dlaczego w U6–U8 ćwiczymy layup TYLKO mocną ręką?",
             ["Zasady zabraniają layupu słabszą ręką", "Żeby najpierw dobrze opanować jedną rękę",
              "Słabsza ręka służy tylko do obrony", "Tak jest szybciej"],
             1,
             "Progresja Żalgirisu: najpierw mocna ręka perfekcyjnie, potem (w U9–U10) dokładamy słabszą. "
             "Nauka dwóch rzeczy naraz sprawia, że obie wychodzą słabiej. Jedno po drugim!"),
        ]
    ),

    (
        "Pivot — obrót z piłką",
        "zasady", 4, "U8", 104,
        """Pivot = obrót na jednej nodze z piłką. To legalne!

Jak to działa:
• Masz piłkę i stoisz zatrzymany
• Jedna noga = stopa obrotowa (pivot foot) — NIGDY jej nie podnosisz
• Drugą nogą możesz chodzić dookoła jak kompasem

Dwa rodzaje:
FORWARD PIVOT — obracasz się do przodu
REVERSE PIVOT — obracasz się do tyłu

Kiedy używasz?
• Gdy obrońca za blisko — odwróć się tyłem (reverse)
• Gdy szukasz miejsca na podanie — obróć się i przejrzyj boisko
• Po jump stopie — wybierz dowolną stopę jako pivot foot!

Błąd: Jeśli podniesiesz pivot foot i postawisz ją z powrotem — błąd kroków!""",
        [
            ("Pivot to...",
             ["Kozłowanie tyłem", "Obrót na jednej nodze trzymając piłkę",
              "Podanie za plecami", "Rzut z obrotu"],
             1,
             "Pivot = obrót, gdzie JEDNA noga stoi w miejscu (stopa obrotowa) "
             "a druga chodzi wokół niej. Piłka jest w rękach. To legalny ruch!"),
            ("Stopa obrotowa (pivot foot) podczas obrotu...",
             ["Może się przesuwać po boisku", "Musi stać w miejscu cały czas",
              "Musi być prawa noga", "Musi być lewa noga"],
             1,
             "Stopa obrotowa = Twój kotwica! Nie możesz jej przenosić. "
             "Możesz ją OBRACAĆ (na piętce lub na przodostopiu), ale nie podnosić i stawiać gdzie indziej."),
            ("Po jump stopie Twoja stopa obrotowa to...",
             ["Zawsze prawa noga", "Zawsze lewa noga",
              "Noga która ląduje pierwsza", "Dowolna noga — sam decydujesz"],
             3,
             "To NAJLEPSZA rzecz w jump stopie! Gdy lądujemy na obie nogi jednocześnie, "
             "SAMI wybieramy która będzie pivot foot. Daje to dużo możliwości!"),
        ]
    ),

    (
        "Jab step — zwód ciałem przed atakiem",
        "taktyka", 4, "U8", 105,
        """Jab step to mały krok, który myli obrońcę — zanim wykonasz ruch!

Jak to działa:
1. Stoisz z piłką (po zatrzymaniu lub po złapaniu)
2. Robisz MAŁY SZYBKI krok w jedną stronę (to jest jab step)
3. Patrzysz na obrońcę — jak zareaguje?
4. Jeśli cofnął się → STRZELASZ z rzutu
5. Jeśli nie cofnął się → ATAKUJESZ kozłem w tę stronę

Kluczowe: Jab foot to NIE jest Twoja pivot foot — możesz ją cofnąć!

To podstawa gry 1 na 1 z Żalgirisu — uczą jej już w U6–U8!""",
        [
            ("Co to jest jab step?",
             ["Szybkie kozłowanie w prawo", "Mały szybki krok, który myli obrońcę",
              "Rzut z jednej nogi", "Bieg tyłem"],
             1,
             "Jab step = mały, szybki krok w jedną stronę. To ZWÓD — sprawdzasz reakcję obrońcy. "
             "Ciało sugeruje atak w jedną stronę, ale możesz zmienić decyzję!"),
            ("Gdy po jab stepie obrońca cofnął się, powinieneś...",
             ["Atakować kozłem", "Strzelić z rzutu", "Cofnąć się i podać", "Stać w miejscu"],
             1,
             "Obrońca cofnął się = MA strach przed Twoim atakiem! "
             "To idealny moment na RZUT — masz przestrzeń. Nie trać tej szansy!"),
            ("Czy noga, którą robisz jab step, może wrócić do tyłu?",
             ["Nie, to błąd kroków", "Tak, bo to NIE jest pivot foot",
              "Tylko jeśli sędzia nie widzi", "Tak, ale tylko przy jab w lewo"],
             1,
             "Przy zatrzymaniu: jedna noga to pivot (nie możesz jej ruszyć), "
             "DRUGA noga to Twoja 'swobodna' noga — możesz nią robić jab step i cofać ją. "
             "To nie jest błąd!"),
        ]
    ),

    # ══════════════════════════════════════════════════════════
    # U9–U10 | Słabsza ręka + PIERWSZE TACTICAL SKILLS
    # Żalgiris: zwody kozłem, layup słabszą ręką, power layup,
    #           podania słabszą ręką, rzut z półdystansu po koźle,
    #           obrona 1x1, box-out.
    #           Taktyka: 3x0, Read & React, Give & Go, Spacing.
    # ══════════════════════════════════════════════════════════

    (
        "Spacing — jak rozciągać boisko",
        "taktyka", 5, "U10", 201,
        """Spacing = rozstawienie zawodników na całym boisku. Pierwsza zasada taktyki!

Dlaczego to ważne?
Gdy 5 zawodników stoi blisko siebie → obrona kryje wszystkich małą strefą.
Gdy zawodnicy rozstawieni są szeroko → obrona musi się 'rozciągnąć'.

Podstawowe ustawienie: 4 out — 1 in
• 4 zawodników na linii 3PT lub dalej (OUT)
• 1 zawodnik w farbie lub na poście (IN)

Zasady spacingu Żalgirisu:
• Nie stój w miejscu gdzie masz kozłującego kolegę → blookujesz jego ścieżkę
• Gdy kolega jedzie do kosza → TY przesuń się na wolne miejsce
• Zawodnicy NIE biegają w groupach → max 2 osoby przy jednej stronie boiska

Żalgiris uczy spacingu już od U9–U10!""",
        [
            ("Spacing oznacza...",
             ["Szybkie bieganie po całym boisku",
              "Rozstawienie zawodników szeroko na boisku",
              "Zbieranie się wszyscy przy piłce",
              "Kozłowanie w środku boiska"],
             1,
             "Spacing = rozstawienie! Gdy zawodnicy stoją daleko od siebie, "
             "obrońcy muszą pokryć większą powierzchnię. To tworzy wolne miejsca i przewagi."),
            ("Gdy kolega z piłką atakuje kozłem do kosza, Ty powinieneś...",
             ["Biec do niego by mu pomóc", "Stać w miejscu i czekać",
              "Przesunąć się na wolne miejsce (relocate)", "Biec pod kosz po zbiórce"],
             2,
             "Gdy kolega atakuje → TY relocatujesz (przesuwasz się na lepsze miejsce)! "
             "Stanie w miejscu blokuje jego ścieżkę. Ruch jest zawsze lepszy od stania."),
            ("Co oznacza '4 out 1 in'?",
             ["4 graczy atakuje, 1 broni",
              "4 graczy stoi za linią 3PT, 1 w farbie",
              "4 graczy na jednej stronie boiska",
              "4 graczy pod koszem, 1 daleko"],
             1,
             "'Out' = poza strefą (za linią 3PT lub dalej). 'In' = w farbie/poście. "
             "To podstawowe ustawienie Żalgirisu od U9 wzwyż!"),
            ("Co się dzieje gdy wszyscy zbiorą się w jednym miejscu?",
             ["Łatwiej strzelić do kosza",
              "Obrona ma łatwiej — kryje małą strefę",
              "Tworzymy silną formację ataku",
              "Nic szczególnego się nie dzieje"],
             1,
             "Stłoczenie = raj dla obrony! Mogą kryć wszystkich bez trudu. "
             "Dobry spacing 'rozkłada' obronę i tworzy wolnych zawodników."),
        ]
    ),

    (
        "Give & Go — podaj i wbiegaj",
        "taktyka", 5, "U10", 202,
        """Give & Go = podaj piłkę i natychmiast wbiegaj do kosza po powrót!

Dlaczego to działa?
Obrońca często 'odpoczywa' po tym jak jego zawodnik odda piłkę.
To Twoja szansa — zanim obrońca zareaguje, jesteś już przy koszu!

Jak wykonać:
1. Masz piłkę, Twój obrońca blisko
2. Podajesz piłkę do kolegi
3. NATYCHMIAST wbiegasz w stronę kosza (bez zatrzymywania!)
4. Kolega oddaje Ci piłkę przy koszu
5. Kończysz layupem lub krótkm rzutem

Kiedy szczególnie skuteczne:
• Gdy obrońca 'idzie' za piłką (head turns = zwrot głowy)
• Gdy obrońca stoi między Tobą a koszem

Give & Go to 'warstwa 1' systemu Read & React — uczy się jej w U9–U10.""",
        [
            ("Give & Go oznacza...",
             ["Podaj piłkę i zostań w miejscu",
              "Podaj piłkę i natychmiast wbiegaj do kosza",
              "Podaj piłkę i wróć do obrony",
              "Podaj i cofnij się za linię 3PT"],
             1,
             "'Give' = daj (podaj). 'Go' = idź (wbiegaj)! "
             "Dajesz piłkę i od razu biegniesz do kosza. "
             "Twój obrońca często 'odpuści' po tym jak oddasz piłkę — to Twój moment!"),
            ("Kiedy Give & Go jest najskuteczniejsze?",
             ["Gdy jesteś bardzo daleko od kosza",
              "Gdy obrońca odwraca głowę za piłką",
              "Gdy grasz strefę",
              "Gdy masz zmęczone nogi"],
             1,
             "Obrońca często 'idzie' za piłką wzrokiem gdy ją podajesz (głowa się obraca). "
             "W tym momencie jesteś 'niewidzialny'! Wbiegasz za jego plecami."),
            ("Po podaniu w Give & Go, wbiegasz...",
             ["Wolno, żeby nie pomylić kolegi",
              "W stronę kosza najszybciej jak potrafisz",
              "Z powrotem za linię 3PT",
              "W stronę narożnika"],
             1,
             "NATYCHMIAST i SZYBKO! Każda sekunda daje obrońcy czas na reakcję. "
             "Cel: być przy koszu zanim obrońca zorientuje się co robisz."),
        ]
    ),

    (
        "Layup słabszą ręką — dlaczego to ważne?",
        "taktyka", 5, "U10", 203,
        """W U9–U10 dokładamy drugą rękę — słabszą. To wielki krok naprzód!

Dlaczego słabsza ręka jest ważna?
• Gdy atakujesz z lewej strony → layup lewą ręką jest naturalny i trudniejszy do zablokowania
• Zawodnik który używa TYLKO mocnej ręki → obrońca wie gdzie trafi
• Dwuręczny zawodnik = nieprzewidywalny = trudniejszy do krycia

Power Layup (dochodzi w U9–U10):
• Lądowanie na DWA stopy jednocześnie (jump stop)
• Silniejszy kontakt z obrońcą — trudniejszy do wytrącenia
• Dobry przy jeździe do kosza na bardzo bliskim dystansie

Progresja Żalgirisu:
U6–U8: layup mocną ręką
U9–U10: layup słabszą ręką + power layup
U11–U13: eurostep, floater, reverse layup""",
        [
            ("Dlaczego ważne jest ćwiczenie layupu słabszą ręką?",
             ["Bo tak mówi trener",
              "Bo z lewej strony dochodzisz z lewą ręką — to naturalniejsze i trudniejsze do blokowania",
              "Bo mocna ręka i tak zawsze wychodzi lepiej",
              "Bo słabsza ręka działa szybciej"],
             1,
             "Gdy atakujesz z lewej → layup lewą ręką jest bliżej kosza i dalej od obrońcy. "
             "Zawodnik z jedną ręką = przewidywalny. Zawodnik z dwoma = koszmar obrony!"),
            ("Power layup różni się od zwykłego tym, że...",
             ["Rzucasz mocniej od dołu",
              "Lądowanie na dwa stopy jednocześnie (jump stop)",
              "Używasz obu rąk do rzutu",
              "Robisz trzy kroki zamiast dwóch"],
             1,
             "Power layup = kończysz JUMP STOPEM (obie stopy jednocześnie). "
             "Jesteś stabilniejszy, trudniejszy do powalenia. Dobry gdy masz obrońcę na sobie!"),
            ("Kiedy w Żalgirisu zaczyna się ćwiczyć layup słabszą ręką?",
             ["U6–U8", "U9–U10", "U11–U13", "U14–U15"],
             1,
             "Żalgiris: U6–U8 = tylko mocna ręka (najpierw perfekcja). "
             "U9–U10 = dokładamy słabszą ręką i power layup. Krok po kroku!"),
        ]
    ),

    (
        "Obrona 1 na 1 — pozycja i kroki dostawne",
        "taktyka", 5, "U10", 204,
        """Obrona zaczyna się od POZYCJI — zanim nastąpi jakikolwiek atak!

Prawidłowa pozycja obrońcy:
• Stopy szersze niż ramiona, ugięte kolana
• Ciężar na przodostopiu (gotowość do ruchu)
• Ręce aktywne — jedna przy piłce, druga blokuje lane
• Oczy na BRZUCH atakującego — nie daj się zmylić zwodami głową/stopami

Kroki dostawne (defensive slide):
• Krok w stronę ataku pierwszą nogą
• Przysuń drugą nogę — ale NIE krzyżuj nóg!
• Skrzydłowy kroczy → Ty slajdujesz, nie biegniesz
• NIE krzyżuj nóg — stracisz równowagę i zawodnik Cię minie

Błąd #1 początkujących: patrzysz na piłkę zamiast na brzuch. Piłka jest 'zwodnicą'!

Żalgiris uczy pozycji i kroków dostawnych od U9–U10.""",
        [
            ("W pozycji obronnej patrzysz na...",
             ["Piłkę — żeby ją przechwycić", "Twarz atakującego — żeby czytać jego miny",
              "Brzuch atakującego — bo to centrum masy", "Nogi atakującego — bo tam widać kierunek"],
             2,
             "Brzuch = centrum masy ciała. Możesz oszukać wzrok i głową i rękoma, "
             "ale BRZUCH zawsze idzie w tym samym kierunku co atak. Dobry obrońca patrzy na brzuch!"),
            ("Podczas kroków dostawnych (defensive slide) NIE powinieneś...",
             ["Uginać kolan", "Krzyżować nóg",
              "Trzymać rąk aktywnych", "Patrzeć na atakującego"],
             1,
             "Skrzyżowanie nóg = utrata balansu = łatwe minięcie! "
             "Krok w stronę → dosuń drugą → nigdy nie krzyżuj. "
             "To właśnie 'defensive slide' — ślizganie się, nie bieganie."),
            ("Gdzie masz ręce w pozycji obronnej?",
             ["Za plecami żeby nie faulować",
              "Obniżone wzdłuż ciała",
              "Jedna przy piłce, druga blokuje ścieżkę do kosza",
              "Obie przy piłce, żeby kraść"],
             2,
             "Aktywne ręce robią dużo: jedna 'zagraża' piłce (zmusza do trudniejszego podania), "
             "druga blokuje lane (ścieżkę) do kosza lub do kolegi. "
             "Ręce za plecami = brak obrony!"),
        ]
    ),

    (
        "Podanie — klucz do dobrego ataku",
        "taktyka", 5, "U10", 205,
        """Koszykówka to sport podań. W U9–U10 dokładamy słabszą rękę!

Rodzaje podań (podstawowe):
• PODANIE Z KLATKI (chest pass) — z wysokości klatki piersiowej, prosto do kolegi
• PODANIE ZAD GŁOWY (overhead pass) — nad głową, do wyżej stojącego kolegi
• PODANIE KOZŁEM (bounce pass) — odbija się od podłogi — dobre gdy obrońca ma ręce wysoko

Progresja Żalgirisu:
U6–U8: podanie mocną ręką z miejsca
U9–U10: podanie SŁABSZĄ ręką z miejsca i po koźle + 'transition off-pass'

Klucze dobrego podania:
• Cel: chwytna strona kolegi (prawa strona = prawa dłoń kolegi)
• Siła: nie za mocno, nie za słabo — kolega musi złapać bez trudu
• Timing: podaj zanim obrońca zaatakuje, nie po!

Transition off-pass: szybkie podanie od razu po zbiórce, zanim obrona wróci. To rozruch szybkiego ataku!""",
        [
            ("Kiedy używasz podania kozłem (bounce pass)?",
             ["Gdy kolega jest bardzo daleko",
              "Gdy obrońca ma ręce wysoko — przebijasz pod jego rękami",
              "Gdy chcesz zaskoczyć sędziego",
              "Zawsze — to najbezpieczniejsze podanie"],
             1,
             "Bounce pass = piłka odbywa się pod rękoma obrońcy! "
             "Szczególnie użyteczny w grze postu i przy obronie dobrze blokującej lane."),
            ("Transition off-pass to...",
             ["Podanie za linią końcową",
              "Szybkie podanie po zbiórce by ruszyć szybki atak zanim obrona wróci",
              "Podanie po wyjściu z boiska",
              "Podanie z przełożeniem za plecami"],
             1,
             "Szybki atak zaczyna się od SZYBKIEGO PODANIA po zbiórce! "
             "Każda sekunda zwłoki daje obronie czas na powrót. "
             "Transition off-pass = natychmiast oddaj piłkę po zbiórce i biegnij!"),
            ("W którym miejscu ciała kolegi celujesz przy podaniu?",
             ["Nad głową — żeby łatwiej złapał",
              "W stronę chwytnej ręki — żeby mógł złapać bez trudu",
              "W nogi — żeby się pochylił",
              "W tors — zawsze"],
             1,
             "Cel: chwytna strona kolegi. Jeśli kolega chce złapać prawą ręką → celujesz w jego prawą stronę. "
             "Dobre podanie to takie, które kolega może złapać i od razu coś z nim zrobić!"),
        ]
    ),

    # ══════════════════════════════════════════════════════════
    # U11–U13 | Zaawansowane umiejętności + podstawy teamowe
    # Żalgiris: eurostep, floater, reverse, handoff, box-out,
    #           closeout, deny, bump cuts.
    #           Taktyka: shell box, team OFF, 3x0/4x0, backdoor, szybki atak.
    # ══════════════════════════════════════════════════════════

    (
        "Eurostep, floater i reverse layup",
        "taktyka", 6, "U13", 301,
        """W U11–U13 wychodzisz poza standardowy layup. Czas na finishes!

EUROSTEP:
• Dwa kroki — każdy w inny kierunek
• Krok 1: prawa (obrońca reaguje)
• Krok 2: lewa (mijasz go!)
• Skończyć można dwiema rękami lub słabszą

FLOATER (łódka):
• Rzut jedną ręką z bliskiego dystansu — ale wyższy łuk niż layup
• Używasz gdy obrońca jest za blisko pod koszem (rim protector)
• Piłka 'pływa' ponad blokerem i opada do kosza

REVERSE LAYUP:
• Zamiast rzucać od frontu → obchodzisz kosz i rzucasz z DRUGIEJ STRONY tablicy
• Używasz gdy bloker stoi na Twojej ścieżce do kosza
• Kończysz od tablicy od tyłu

Żalgiris: te umiejętności pojawiają się w U11–U13 — po opanowaniu podstaw layupu.""",
        [
            ("W eurostepie kluczowe jest to, że...",
             ["Robisz 3 kroki zamiast 2",
              "Dwa kroki idą w różne kierunki — mija obrońcę",
              "Używasz tylko lewej ręki",
              "Skaczysz nad obrońcą"],
             1,
             "Eurostep = zmiana kierunku w trakcie dwóch kroków! "
             "Obrońca reaguje na pierwszy krok i 'pada' — mijasz go drugim krokiem. "
             "To legalne — FIBA zezwala na dwa kroki."),
            ("Kiedy używasz floatera?",
             ["Gdy masz dużo miejsca i czasu",
              "Gdy obrońca jest daleko pod koszem",
              "Gdy pod koszem stoi wysoki bloker — przerzucasz nad nim",
              "Gdy jesteś zmęczony i chcesz łatwiejszego rzutu"],
             2,
             "Floater = rzut 'z lotu' nad blokerem! Wysoki łuk sprawia, że piłka opada z góry — "
             "trudno ją zablokować. Ideał gdy pod koszem czeka rim protektor."),
            ("Reverse layup to...",
             ["Layup tyłem do kosza",
              "Layup z drugiej strony tablicy — obchodzisz kosz",
              "Layup lewą ręką zamiast prawą",
              "Layup kozłem między nogami"],
             1,
             "Reverse = obchodzisz kosz i finiszujesz z DRUGIEJ strony tablicy. "
             "Obrońca blokuje Twój standardowy layup? Idź dookoła! "
             "Piłka odbija się od tablicy od strony tylnej."),
        ]
    ),

    (
        "Shell Box — podstawy obrony zespołowej",
        "taktyka", 6, "U13", 302,
        """Shell Box to trening, który buduje podstawy każdej obrony drużynowej.

Podstawowe zasady Shell Box:
• Każdy obrońca ma 1 obowiązek: kryć swojego zawodnika LUB być w pozycji pomocy
• Obrońca przy piłce = NACISK (deny lub on-ball)
• Obrońca dalej od piłki = POMOC (help position, po stronie piłki)
• Gdy piłka się przesuwa → wszyscy rotują!

Cztery ustawienia w Shell:
1. Piłka na rogu — obrońca przy piłce, reszta po stronie piłki
2. Piłka na skrzydle — 'słaby' boczny obrońca schodzi do linii środkowej
3. Piłka na poście — podwojenie lub remain
4. Drive obrońcy — wszyscy rotują (help & recover)

Dlaczego Shell Box?
• Uczy gdzie być gdy NIE masz swojego zawodnika przy piłce
• Podstawa dla 'help & recover' z U14–U15
• Żalgiris zaczyna Shell Box w U11–U13""",
        [
            ("W Shell Box, obrońca DALEKO od piłki powinien...",
             ["Atakować piłkę i kraść",
              "Stać przy swoim zawodniku jakby piłka była przy nim",
              "Schodzić do pozycji pomocy po stronie piłki",
              "Cofać się pod kosz"],
             2,
             "Obrońca daleko od piłki = POMOC! Schodzi między swojego zawodnika a kosz, "
             "po stronie gdzie jest piłka. Gdy zawodnik z piłką przebija — TY jesteś hamulcem!"),
            ("Co się dzieje gdy piłka się przesuwa do innej strony boiska?",
             ["Obrońcy zostają na swoich miejscach",
              "Wszyscy obrońcy rotują razem z piłką",
              "Obrońca przy piłce biega za nią",
              "Zmiana obrońców co każde podanie"],
             1,
             "Rotacja to KLUCZ obrony! Gdy piłka przesuwa się w prawo → cała obrona przesuwa się w prawo. "
             "'Shell Box' to właśnie trening tej rotacji — zanim to wejdzie w meczową automatykę."),
            ("Obrońca 'help & recover' to...",
             ["Obrońca który fauluje i wraca do ławki",
              "Obrońca który wychodzi do pomocy i wraca do swojego zawodnika",
              "Obrońca który zbiera piłkę po rzucie",
              "Trener pomagający przy substytucjach"],
             1,
             "Help & recover: wychodzisz pomóc (gdy kolega jest miniony) → jego zawodnik jest wolny → "
             "Twój kolega wraca do krycia → Ty wracasz do swojego. Musi być szybkie i skoordynowane!"),
        ]
    ),

    (
        "Backdoor cut — kiedy obrońca za blisko",
        "taktyka", 5, "U13", 303,
        """Backdoor cut = atak ZA plecami obrońcy. Nagradzasz jego błąd!

Kiedy użyć backdoor cuta?
• Obrońca stoi MIĘDZY Tobą a piłką (overplays)
• Obrońca odwraca głowę za piłką
• Jesteś za linią 3PT i obrońca bardzo blisko

Jak wykonać:
1. Stoisz na skrzydle, obrońca Cię 'overplayuje' (blokuje drogę do piłki)
2. Robisz fałszywy krok DO PIŁKI (on reaguje!)
3. NATYCHMIAST zmieniasz kierunek — wbiegasz za jego plecami do kosza
4. Podający widzi Twój ruch → podaje natychmiast w miejsce przy koszu
5. Kończysz layupem

Backdoor to 'odpowiedź' na agresywną obronę. Żalgiris uczy cięcia backdoor w U11–U13, jako część ataku 3x0/4x0.""",
        [
            ("Backdoor cut NAJLEPIEJ działa gdy...",
             ["Obrońca stoi daleko od Ciebie",
              "Obrońca agresywnie blokuje Ci drogę do piłki (overplays)",
              "Jesteś przy koszu",
              "Masz otwartą linię 3PT"],
             1,
             "Overplay = błąd obrońcy! Stoi zbyt blisko by odciąć Ci podanie → "
             "plecami do kosza. Backdoor karze go za ten błąd."),
            ("Fałszywy krok przed backdoor cutem służy do...",
             ["Zwolnienia tempa gry",
              "Zmylenia obrońcy — reaguje na fałszywy kierunek",
              "Zdobycia lepszej pozycji do rzutu",
              "Wywołania faulu obrońcy"],
             1,
             "Fałszywy krok = 'set up'! Robisz krok jakbyś szedł po piłkę → obrońca reaguje → "
             "ODWRACASZ się i biegniesz w przeciwnym kierunku. On jest już za Tobą!"),
            ("Kto inicjuje backdoor cut?",
             ["Zawodnik z piłką — on decyduje",
              "Trener z ławki",
              "Zawodnik BEZ piłki — on widzi obrońcę i decyduje kiedy ciąć",
              "Sędzia"],
             2,
             "TY (bez piłki) widzisz, że obrońca Cię overplayuje — TY decydujesz kiedy ciąć! "
             "Zawodnik z piłką obserwuje i natychmiast podaje gdy zobaczy Twój backdoor. "
             "To komunikacja bez słów — 'Read & React'!"),
        ]
    ),

    (
        "Box out — zastawianie przy zbiórce",
        "taktyka", 5, "U13", 304,
        """Zbiórka zaczyna się PRZED rzutem — od box outa!

Co to box out?
• Ustawiasz ciałem między napastnikiem a koszem
• On nie może przejść obok Ciebie do piłki
• TY jesteś bliżej kosza — piłka Twoja!

Jak wykonać box out:
1. Widzisz że rzut poleciał
2. Znajdź swojego zawodnika (find your man)
3. Obróć się na niego tyłem — tył ciała blokuje go
4. Stopy szeroko, kolana ugięte, łokcie na zewnątrz
5. Utrzymuj kontakt plecami — gdy on chce obejść, Ty obracasz się z nim
6. Po złapaniu piłki → od razu szukaj podania (outlet pass)

Zasada Żalgirisu: 'Defense is not finished until the ball is secured'
(Obrona nie jest skończona dopóki nie złapiemy piłki).

Żalgiris: box out pojawia się oficjalnie w U9–U10 ale ćwiczycie go intensywnie w U11–U13!""",
        [
            ("Box out to ustawienie...",
             ["Przy linii końcowej przed rzutem wolnym",
              "Ciałem między napastnikiem a koszem po oddaniu rzutu",
              "Obrońcy przy koszu w ataku strefowym",
              "Dwóch zawodników razem blokujących jednego"],
             1,
             "Box out = zastawianie! Stawiasz ciało między atakującym a koszem. "
             "On musi 'obejść' Cię — dając Ci czas na złapanie piłki."),
            ("Kiedy wykonujesz box out?",
             ["Gdy piłka jest pod Twoim koszem",
              "Gdy tylko rzut poleci w kierunku kosza",
              "Gdy jesteś zmęczony",
              "Tylko w końcówce meczu"],
             1,
             "BOX OUT jest natychmiastową reakcją na KAŻDY rzut! "
             "Nie czekasz czy 'on rzuci' — gdy tylko widzisz ruch do rzutu, szukasz zawodnika do zastawienia."),
            ("Jak Żalgiris definiuje koniec obrony?",
             ["Po zakończonym rzucie",
              "Gdy sędzia gwizdnie",
              "Dopiero gdy złapiemy piłkę po rzucie",
              "Po 24 sekundach ataku"],
             2,
             "'Defense is not finished until the ball is secured!' "
             "Możesz perfekcyjnie bronić przez 23 sekundy, ale jeśli przeciwnik złapie ofensywną zbiórke — "
             "oni dostają nowe posiadanie. Zbiórka = ostatni element obrony!"),
        ]
    ),

    (
        "Szybki atak — pierwsze sekundy po zbiórce",
        "taktyka", 7, "U13", 305,
        """Szybki atak (fast break) = zdobycie punktów ZANIM obrona wróci na swoje miejsce!

Sekwencja po defensywnej zbiórce:
1. ZBIERZ (box out → złap piłkę)
2. PODAJ NATYCHMIAST (outlet pass) → rzucający kozłować NIE powinien
3. Biegną 3 tory: środkowy (PG z piłką) + dwa skrzydłowe
4. PG jedzie do środka, skrzydłowi szerocy
5. Gdy PG ma przewagę → kończy layupem
6. Gdy obrońca blokuje → podaje na skrzydło
7. Wszyscy ładują się — także 'czwarty i piąty' zawodnik na sekundy po nich

Zasada Żalgirisu: pierwsze 3 kroki są decydujące!
Jak szybko ruszysz po zbiórce = czy masz przewagę czy nie.

Typy szybkiego ataku:
• 2 na 1 — wielka przewaga, kończ layupem lub podaj pod kosz
• 3 na 2 — jedna dodatkowa osoba, atakuj środkowy pas""",
        [
            ("Co robisz PIERWSZEGO po defensywnej zbiórce?",
             ["Kozłujesz jak najszybciej do przodu",
              "Natychmiast oddajesz outlet pass do rozgrywającego",
              "Szukasz wolnego zawodnika przy linii 3PT",
              "Czekasz aż pozostali wrócą na pozycje"],
             1,
             "OUTLET PASS! Zbierający nie kozłuje — natychmiast oddaje do PG. "
             "Kozłowanie traci czas. Podanie pozwala uruchomić szybki atak zanim obrona wróci."),
            ("W szybkim ataku 3 na 2, skrzydłowi powinni biec...",
             ["W grupie z PG", "Szeroko przy liniach bocznych",
              "Pod kosz razem", "Czekają za linią 3PT"],
             1,
             "Skrzydłowi SZEROKO! To rozciąga 2 obrońców. "
             "Gdy obaj obrońcy skupią się na PG → skrzydłowy jest wolny. "
             "Gdy kryją skrzydłowego → PG finiszuje pod koszem."),
            ("Zasada 'pierwsze 3 kroki' w szybkim ataku oznacza...",
             ["Możesz zrobić tylko 3 kroki",
              "Szybkość pierwszych kroków po zbiórce decyduje czy zdobędziesz przewagę",
              "Podanie musi nastąpić po 3 krokach",
              "Trzeci zawodnik decyduje o ataku"],
             1,
             "Żalgiris: 'pierwsze 3 kroki są decydujące'! "
             "Jeśli pierwsze kroki po zbiórce są szybkie → jesteś przed obroną. "
             "Jeśli powolne → obrona wraca i nie ma ataku. Pierwsze sekundy to wygrany lub przegrany fast break."),
        ]
    ),

    # ══════════════════════════════════════════════════════════
    # U14–U15 | Pick & Roll wchodzi! + 5x5, statystyki
    # Żalgiris: P&R DEF (Flat/Side/Step Out/Trap/Under),
    #           P&R ATK, Shell Box, 5x5 DEF help&recover,
    #           5x5 OFF spacing, Transition, rzut za 3 z miejsca.
    # ══════════════════════════════════════════════════════════

    (
        "Pick & Roll — zasłona z kozłem (atak)",
        "taktyka", 7, "U15", 401,
        """P&R to najczęstszy element ataku w nowoczesnej koszykówce. Uczy się go w U14–U15!

Elementy P&R:
• BALL HANDLER (BH) — zawodnik z piłką
• SCREENER — zawodnik stawiający zasłonę (blokuje obrońcę BH)

Sekwencja ataku:
1. Screener ustawia zasłonę przy obrońcy BH
2. BH kozłuje 'po zasłonie' (używa jej by ominąć obrońcę)
3. Screener WYCHODZI (roll) — atakuje do kosza lub (pop) — cofa się na rzut
4. BH decyduje: sam atakuje lub podaje screenerowi

Opcje screener'a:
• ROLL — biega do kosza po zasłonie
• POP (fade) — cofa się za linię 3PT

Czytanie P&R (Read):
• Obrońca screener'a zostaje → BH atakuje sam
• Obrońca BH zostaje → BH podaje screener'owi
• Double team → BH szuka otwartego zawodnika""",
        [
            ("W P&R 'Roll' screener'a oznacza...",
             ["Screener stoi w miejscu po zasłonie",
              "Screener wraca za linię 3PT po zasłonie",
              "Screener wbiega do kosza po zasłonie",
              "Screener biegnie do narożnika"],
             2,
             "ROLL = po zasłonie screener BIEGA DO KOSZA! "
             "Jego obrońca musi zdecydować: pilnujesz BH czy gonisz screener'a? "
             "Jedna z nich jest wolna — BH czyta i podaje."),
            ("Kiedy BH powinien oddać podanie screener'owi?",
             ["Zawsze — screener jest zawsze lepiej ustawiony",
              "Gdy obrońca screener'a skupi się na BH zamiast na screener'ze",
              "Gdy BH jest zmęczony",
              "Tylko jeśli jest 3:0"],
             1,
             "Read P&R: obrońca screener'a 'idzie' za BH → screener jest wolny → podajesz! "
             "Obrońca screener'a pilnuje screener'a → BH atakuje sam. To ciągłe czytanie decyzji obrony."),
            ("'Pop' w P&R to...",
             ["Screener biega do kosza",
              "Screener cofa się za linię 3PT na rzut",
              "Screener wraca pod swój kosz",
              "Screener wychodzi na aut"],
             1,
             "Pop = screener cofa się za linię 3PT! Gdy screener jest dobrym strzelcem 3PT, "
             "obrońca musi za nim wyjść → zostawia przestrzeń dla BH. Albo screener strzela!"),
            ("Żalgiris zaczyna P&R w...",
             ["U9–U10", "U11–U13", "U14–U15", "U16–U18"],
             2,
             "Pick & Roll to cezura w curriculum Żalgirisu — wchodzi DOPIERO w U14–U15! "
             "Wcześniej wszystko opiera się na grze 1x1, 2x2 i ruchach bez zasłon. "
             "P&R wymaga dojrzałości taktycznej."),
        ]
    ),

    (
        "Obrona P&R — Flat, Side, Under, Trap",
        "taktyka", 8, "U15", 402,
        """Skoro w U14–U15 uczycie się P&R w ataku, musicie też wiedzieć jak go bronić!

5 sposobów obrony P&R (Żalgiris używa tych samych nazw od U14 do pierwszej drużyny):

FLAT:
• Obrońca BH prześlizguje się NIE za zasłoną ale BOK (flat) przy screener'ze
• Screener'owy obrońca pomaga, a potem wraca
• Dobry gdy BH nie jest strzelcem 3PT

SIDE (Hard Hedge):
• Obrońca screener'a WYCHODZI agresywnie na BH
• BH jest spowolniony, musi kozłować dalej
• Dobry przy Elite Ball Handlerach

STEP OUT:
• Obrońca screener'a wychodzi tylko KROK
• BH nie przechodzi swobodnie, ale screener'owy wraca szybko
• Kompromis między Flat a Side

TRAP (Blitz):
• Podwójne krycie BH — on i obrońca screener'a razem naciskają
• Pozostali 3 obrońcy kryją 4 zawodników
• Bardzo agresywne, ryzykowne

UNDER:
• Obrońca BH prześlizguje się POD zasłoną (między screener'em a kosem)
• Pozwala BH na szybki rzut 3PT!
• Używasz gdy BH NIE jest strzelcem""",
        [
            ("Coverage 'Under' w P&R oznacza...",
             ["Oba obrońcy idą pod koszem",
              "Obrońca BH przechodzi pod zasłoną — po stronie bliższej koszowi",
              "Screener'owy obrońca schodzi pod koszem",
              "Podwójne krycie od dołu"],
             1,
             "Under = obrońca BH 'prześlizguje się' pod zasłoną (po stronie kosza). "
             "To BEZPIECZNE przy BH który nie jest strzelcem 3PT. "
             "Ale jeśli BH jest strzelcem → ma otwarty rzut 3PT!"),
            ("Kiedy używasz coverdage'u 'Trap/Blitz'?",
             ["Zawsze — to najsilniejsza obrona",
              "Gdy chcesz wymusić błąd od BH i masz szybkich obrońców",
              "Gdy jesteś prowadzący o 20 pkt",
              "Gdy screener jest wysoki"],
             1,
             "Trap/Blitz = podwójne krycie przy zasłonie! Bardzo agresywne. "
             "Ryzykowne: 3 obrońców kryje 4 zawodników. "
             "Użyj gdy BH jest podatny na presję lub chcesz go zmusić do błędu."),
            ("Żalgiris używa tych samych nazw coverage'ów...",
             ["Tylko w U14–U15", "Od U14 do pierwszej drużyny",
              "Tylko w pierwszej drużynie", "Zależy od trenera"],
             1,
             "Jeden słownik od 14-latka do pierwszej drużyny! "
             "Gdy zawodnik z U14 awansuje do Żalgiris-3 czy Żalgiris-1, "
             "rozumie te same komendy. To klucz spójnego systemu."),
        ]
    ),

    (
        "Transition offense — 8 sekund do ataku",
        "taktyka", 6, "U15", 403,
        """Transition offense = atak w biegu, zaraz po zmianie posiadania.

Zasada Żalgirisu: '8 sekund na atak, 5 sekund do połowy'
Ale nie chodzi tylko o przepisy — chodzi o TEMPO!

Hierarchia decyzji w transition (Żalgiris):
1. Layup/łatwe 2 pkt → BIERZ
2. Otwarta trójka → BIERZ
3. Nic nie ma → zwalniasz, ustawiasz atak pozycyjny

'Pierwsze 3 kroki decydują' (Żalgiris):
• Jak szybko ruszysz po zmianie posiadania = jaka masz przewagę

Tory biegania:
• Ball handler środkiem — widzi całe boisko
• Dwa skrzydłowi SZEROKO przy liniach bocznych
• Czwarty zawodnik 'czwarty krok' — biegnie do elipsy
• Piąty zawodnik wraca lub biegnie niespieszenie

Kiedy zwalniasz transition?
• Nie ma przewagi liczebnej
• Zawodnicy nie są na właściwych torach
• Obrona wróciła""",
        [
            ("Ile sekund masz od początku posiadania do ataku wg przepisów FIBA?",
             ["8 sekund do ataku (po przekroczeniu połowy 5 sekund)",
              "24 sekundy na rzut, bez limitu do połowy",
              "10 sekund całości",
              "5 sekund do ataku"],
             0,
             "FIBA: 8 sekund na przejście przez połowę + 24 sekundy na rzut! "
             "Ale ważniejsze jest tempo psychologiczne — atakujesz gdy masz przewagę, "
             "nie czekasz na 24s. Każda sekunda to szansa dla obrony."),
            ("W transition, ball handler biegnie...",
             ["Przy linii bocznej",
              "Środkiem boiska — widzi całe boisko",
              "Przy obrońcy który cofnął się",
              "Za skrzydłowymi"],
             1,
             "Ball handler ŚRODKIEM! Stamtąd widzi obydwa skrzydła "
             "i może zdecydować gdzie podać. Skrzydłowi SZEROKO przy liniach — "
             "to rozkłada obronę."),
            ("Kiedy decydujesz się zwolnić transition i przejść do ataku pozycyjnego?",
             ["Gdy masz 20 pkt przewagi",
              "Zawsze po 8 sekundach",
              "Gdy nie ma przewagi liczebnej i obrona wróciła",
              "Gdy trener krzyczy z ławki"],
             2,
             "Transition ma sens gdy jest PRZEWAGA! Brak przewagi = zwolnij i zagraj pozycyjnie. "
             "Forsowanie złego fast breaku daje punkty przeciwnikowi. "
             "Czytaj sytuację: masz przewagę → jedź, nie ma → zatrzymaj się."),
        ]
    ),

    (
        "Co mówi eFG% i dlaczego jest ważniejszy niż FG%?",
        "statystyki", 5, "U15", 404,
        """FG% (skuteczność z gry) traktuje równo rzuty za 2 i za 3. To błąd!

eFG% = Effective Field Goal Percentage
eFG% = (FGM + 0.5 × 3PM) / FGA

Przykład:
• Zawodnik A: 5/10 z gry, 0 trójek → FG%=50%, eFG%=50%
• Zawodnik B: 4/10 z gry, 4 trójki z 4 prób → FG%=40%, eFG%=60%
• KTO jest skuteczniejszy? ZAWODNIK B — jego eFG% jest wyższy!

Dlaczego? Bo trójka daje 3 punkty, nie 2 — i to powinno być uwzględnione.

TS% (True Shooting Percentage):
TS% = PTS / (2 × (FGA + 0.44 × FTA))
Uwzględnia RÓWNIEŻ rzuty wolne (FT)!

Median eFG% w 2LM 24/25 (BasketKołcz Knowledge Base):
• Median liga: ~40.9% FG% przy median eFG% wyższym gdy uwzględnisz trójki
• P25 FG%: 33,3% — jeśli jesteś poniżej, to problem""",
        [
            ("eFG% różni się od FG% tym, że...",
             ["Liczy tylko rzuty za 3",
              "Uwzględnia wyższą wartość trójek (×1.5)",
              "Nie liczy rzutów wolnych",
              "Dzieli się przez liczbę meczów"],
             1,
             "eFG% dodaje 0.5 × 3PM do licznika — bo trójka jest warta 50% więcej niż dwójka. "
             "FG% traktuje je równo, co zawyża wartość zawodników rzucających z blizy i nieefektywnych."),
            ("Który eFG% jest lepszy?",
             ["Zawodnik A: 6/10 FGA, 0 trójek (eFG%=60%)",
              "Zawodnik B: 5/10 FGA, 3 trójki z 3 prób (eFG%=65%)",
              "Zawodnik C: 7/10 FGA, 0 trójek (eFG%=70%)",
              "Zawodnik D: 4/10 FGA, 4 trójki z 4 prób (eFG%=60%)"],
             2,
             "Zawodnik C: 7/10 = eFG% 70%! "
             "Zawodnik B: (5 + 0.5×3)/10 = 6.5/10 = 65%. "
             "eFG% 70% najwyższe. Ale pamiętaj — to zależy też od ILU rzutów próbuje! "
             "7/10 to dobry wynik — 4/6 byłby podobny eFG% przy mniejszym wolumenie."),
            ("True Shooting (TS%) uwzględnia dodatkowo...",
             ["Asysty", "Rzuty wolne (FT)", "Zbiorki", "Czas gry"],
             1,
             "TS% to najpełniejsza miara skuteczności! "
             "Bierze pod uwagę FGA + FTA (z wagą 0.44). "
             "Zawodnik który zdobywa punkty głównie przez rzuty wolne → niski FG%, ale wysoki TS%."),
        ]
    ),

    (
        "Rola vs statystyki — co naprawdę wygrywia",
        "mentalnosc", 5, "U15", 405,
        """Jeden z najtrudniejszych tematów mentalności: więcej punktów ≠ lepsza drużyna.

Pytanie: Co jest ważniejsze — twoje statystyki czy zwycięstwo drużyny?

Case study z bazy wiedzy:
• Zawodnik X: 20 pkt/mecz, drużyna przegrywa 60% meczów
• Zawodnik Y: 8 pkt, 4 asysty, 2 zbiórki, drużyna wygrywa 70% meczów
• Który jest bardziej wartościowy? ZAWODNIK Y.

Statystyki 'empty calories' (puste kalorie z bazy wiedzy BasketKołcz KB):
• Punkty zdobyte gdy wynik +20 lub -20 (nie wpłynęły na mecz)
• Wysokie użycie piłki (USG%) przy niskiej skuteczności (eFG%)
• Asysty przy złych decyzjach ze stratami

'Winning Impact' — co naprawdę wygrywia wg Żalgirisu:
• Plus/minus (czy drużyna wygrywa gdy grasz)
• Decyzje pod presją
• Gra bez piłki
• Komunikacja i obrona

Twoja wartość dla drużyny = Twoja rola w systemie + Twoja skuteczność w tej roli""",
        [
            ("Zawodnik z 18 pkt/mecz jest na pewno wartościowszy niż zawodnik z 9 pkt/mecz?",
             ["Tak — więcej punktów = większa wartość",
              "Nie — zależy od skuteczności, roli i wpływu na wynik drużyny",
              "Tak — ale tylko w play-offach",
              "Nie — zawodnik z 9 pkt nigdy nie jest potrzebny"],
             1,
             "Statystyki bez kontekstu nic nie mówią! "
             "18 pkt przy złej skuteczności i przegranym meczu to 'empty calories'. "
             "9 pkt w kluczowych momentach z dobrym spacingiem może być wartościowsze."),
            ("Co to są 'empty stats' (puste statystyki)?",
             ["Statystyki wypełnione zerami",
              "Punkty zdobyte gdy wynik jest już rozstrzygnięty — nie wpłynęły na mecz",
              "Statystyki zawodnika który nie gra",
              "Statystyki z przegranych meczów"],
             1,
             "Przykład: 10 pkt w 4. kwarcie przy wyniku -25. Wynik był already done. "
             "Te punkty podbijają statystyki, ale nie pomogły wygrać. "
             "Scout/trener widzi to — sprawdza KIEDY i PRZY JAKIM WYNIKU zdobywałeś punkty."),
            ("Żalgiris definiuje wartość zawodnika jako...",
             ["Sumę jego statystyk za sezon",
              "Jego rolę w systemie + skuteczność w tej roli",
              "Liczbę turniejów wygranych",
              "Jego wzrost i wagę"],
             1,
             "Wartość = rola × skuteczność! Doskonały 'role player' jest cenniejszy niż "
             "zły 'star player'. Żalgiris pyta: 'Czy mogę z nim wygrać?', "
             "nie 'Czy zdobywa dużo punktów?'"),
        ]
    ),

    # ══════════════════════════════════════════════════════════
    # U16–U18 | Strefa, pressing, decision making zaawansowane
    # Żalgiris: zone DEF, press DEF, rotacje, decision making,
    #           rzut 3PT po koźle i zasłonach, kontrola prędkości.
    # ══════════════════════════════════════════════════════════

    (
        "Strefa 2-3 — jak ją atakować",
        "taktyka", 7, "U18", 501,
        """Strefa 2-3 to najczęstsza strefa w koszykówce. Żalgiris uczy ataku strefy DOPIERO w U16–U18!

Dlaczego strefę uczy się późno?
Żalgiris: Strefa i presing pojawiają się DOPIERO w U16–U18.
Wcześniej wyłącznie krycie indywidualne. To filozofia — najpierw naucz się wygrywać 1x1.

Słabości strefy 2-3:
• Linia pasy między 2 górnych i 3 dolnych (tzw. 'elbow')
• Narożniki (corners) — 2 górnych zawodników musi bronić bardzo szeroko
• Wysoki post i 'short corner'

Jak atakować:
1. MIEJ CIERPLIWOŚĆ — strefa zmusza do szybkich decyzji
2. Ruch piłki > ruch zawodników — podawaj szybko wokół strefy
3. Atak 'elbow' — gra łokcia zmusza do kolapsowania strefy
4. Narożnik + wysoki post = 3 zawodników na ruch 2 obrońców
5. Dunk spot (dunker's spot) — zawodnik w rogu pod koszem

Zasada: 'Graj prostą koszykówkę — nie próbuj ośmieszyć strefy, tylko ją eksploruj!'""",
        [
            ("Strefa 2-3 ma największą słabość w...",
             ["Środku farby", "Na skrzydle bliżej 3PT",
              "W narożnikach i przy 'elbow' (łokciach)",
              "W centrum pola trójsekudowego"],
             2,
             "Narożniki = 'stretching zone'! Górny zawodnik strefy musi bronić bardzo szeroko. "
             "Elbow (łokieć) = miejsce między 2 górnych i 3 dolnych — tam jest dziura. "
             "Zaatakuj te miejsca ruchem piłki."),
            ("Dlaczego Żalgiris uczy ataku strefy dopiero w U16–U18?",
             ["Bo strefa jest rzadko używana",
              "Bo wcześniej zawodnicy nie są wystarczająco sprawni fizycznie",
              "Filozofia: najpierw naucz się wygrywać 1x1, a strefę atakuje się znajomością ind. obrony",
              "Bo przepisy zabraniają strefy poniżej U16"],
             2,
             "Żalgiris: 'Strefa i presing DOPIERO w U16–U18 — najpóźniej ze wszystkich elementów.' "
             "Najpierw perfekcyjna gra indywidualna. Kto umie grać 1x1 → umie atakować strefę. "
             "Odwrotność jest rzadsza."),
            ("Jak powinien poruszać się 'dunker' (w dunker's spot)?",
             ["Stać w miejscu przy koszu po silniejszej stronie",
              "Stale biegać pod koszem",
              "Ustawiać się w rogu farby po słabszej stronie (blind spot obrońcy)",
              "Wychodzić na wysoki post"],
             2,
             "Dunker's spot = kąt farby po słabszej stronie. "
             "Dolny obrońca strefy ma 'blind spot' — nie widzi go bezpośrednio. "
             "Cierpliwy zawodnik tam = stały zagrożenie pod koszem. "
             "Żalgiris: 'no jumping/lateral — only running' do tego miejsca."),
        ]
    ),

    (
        "Decision making — kiedy kozłować, podać, strzelić",
        "taktyka", 8, "U18", 502,
        """Decision making = podejmowanie decyzji. To najważniejsza umiejętność w U16–U18!

Zasada 'Ball Movement' Żalgirisu:
'One dribble to attack, two to score, THREE IS TROUBLE'
• 1 kozioł: atakujesz pozycję
• 2 kozłowania: finiszujesz pod koszem
• 3+: zużywasz kozłowania bez celu → obrona się resetuje

Shot Selection (wybór rzutu) wg Żalgirisu:
• Layup/łatwa dwójka ✓
• Rzut wolny (wymuś faul) ✓
• Otwarta trójka ✓
• Trudna dwójka (środek, obrońca blisko) ✗ → to 'anty-shot'
• 'Best possible shot, not the first available shot'

Czytanie gry (Decision Making):
• Zanim dostaniesz piłkę → już wiesz co zrobisz (pre-read)
• Po otrzymaniu: 0.5 sekundy na decyzję
• Patrzysz na HELP OBRONY — czy są wolni?

Żalgiris w U16–U18 uczy CZYTANIA GRY jako głównej umiejętności taktycznej.""",
        [
            ("Co oznacza zasada 'Three is trouble' Żalgirisu?",
             ["Trzecia kwarta jest najtrudniejsza",
              "Trzeci kozioł bez celu = obrona się resetuje i masz gorszą pozycję",
              "Trzech zawodników nie może stać razem",
              "Przy trzecim faulu musisz uważać"],
             1,
             "'One dribble to attack, two to score, THREE IS TROUBLE!' "
             "Każdy kozioł bez celu = czas dla obrony na reset. "
             "Decyzja PRZED kozłowaniem = mniej kozłowań."),
            ("'Best possible shot, not the first available shot' oznacza...",
             ["Czekaj na perfekcyjną szansę",
              "Odmawiaj każdy rzut poniżej 60% skuteczności",
              "Nie bierz pierwszej szansy jeśli możesz znaleźć lepszą w 1-2 sekundy",
              "Strzelaj tylko z narożnika"],
             2,
             "Pierwszy rzut może być 'ok', ale poczekaj chwilę — może kolega jest w lepszej pozycji. "
             "Nie 'hold the ball' na wieczność, ale 1 ruch piłki często tworzy otwarty rzut. "
             "Cierpliwość + ruch piłki = lepsza decyzja."),
            ("Pre-read w koszykówce to...",
             ["Czytanie reguł przed meczem",
              "Analiza przeciwnika przed sezonem",
              "Rozczytanie sytuacji zanim otrzymasz piłkę",
              "Czytanie sygnałów trenera"],
             2,
             "PRE-READ = decydujesz CO ZROBISZ zanim dostaniesz piłkę! "
             "Patrzysz: gdzie jest mój obrońca? Gdzie są moi koledzy? "
             "Gdy piłka do mnie leci → już wiem: podanie w lewo, albo atak, albo rzut. "
             "To eliminuje 'holding the ball' i błędy pod presją."),
            ("Decyzja 'atakuj sam vs podaj' zależy przede wszystkim od...",
             ["Własnej energii i formy dnia",
              "Tego jak obrona reaguje na twoją akcję",
              "Wskazówek trenera z ławki",
              "Wyniku na tablicy"],
             1,
             "Czytasz OBRONĘ! Obrońca zamknął Ci drogę? → podaj. "
             "Droga jest otwarta? → atakuj. Twój obrońca wyszedł za Twoim podaniem? → "
             "wbiegasz za jego plecami. Decyzja = odpowiedź na ruch obrony, nie plan wcześniej ustalony."),
        ]
    ),

    (
        "Rotacje obronne — co robisz gdy kolega jest miniony",
        "taktyka", 7, "U18", 503,
        """Rotacje obronne to działanie zespołowe gdy jeden obrońca jest przebity.

Sekwencja rotacji (Żalgiris):
1. Twój kolega jest miniony przez atakującego
2. TY (closest help) wychodzisz na atakującego — STOP!
3. Twój kolega wraca do krycia — może zająć twoje miejsce
4. Reszta zespołu przesuwa się by zakryć wolnych

Zasada: 'Jeden atakuje — jeden blokuje — reszta zamknięta'

Closest man to basket principle:
• Kto jest najbliżej kosza (ze strony obrony) → on jest pierwszym pomostem
• Niekoniecznie Twój zawodnik — chronisz kosz!

Komunikacja w rotacjach:
• 'Help!' — kolega mówi że potrzebuje pomocy
• 'Got ball' — Ty mówisz że kryjesz atakującego
• 'I'm back' — kolega wraca do krycia
• Żalgiris: 'Głośno, wcześnie, 3 razy PRZED zasłoną!'

Błąd: Czekasz aż kolega biega za minioną, a atakujący ma łatwy layup.""",
        [
            ("Gdy kolega jest miniony, Twoja pierwsza odpowiedzialność to...",
             ["Krzyczeć na niego żeby biegł szybciej",
              "Wyjść na atakującego by zatrzymać ruch do kosza",
              "Zostać przy swoim zawodniku",
              "Biec pod kosz po zbiórce"],
             1,
             "CLOSEST HELP! Gdy kolega jest miniony → TY wychodzisz by zatrzymać atakującego. "
             "Twój zawodnik jest teraz wolny → ktoś musi go zakryć = rotacja reszty!"),
            ("Co to jest 'help & recover'?",
             ["Pomoc medyczna po kontuzji",
              "Wychodzisz do pomocy przy atakującym, a potem wracasz do swojego zawodnika",
              "Pomoc trenerska przy taktyce",
              "Zawodnik pomaga w obronie, potem odpoczywa"],
             1,
             "Help & recover: WYCHODZISZ (help) by zatrzymać atak → "
             "Twój kolega wraca i kryje atakującego (recover) → "
             "TY wracasz do swojego zawodnika. Szybkie, zsynchronizowane."),
            ("Żalgiris mówi że komunikację w obronie zasłon powinieneś zacząć...",
             ["W momencie gdy zasłona jest postawiona",
              "Po wykonaniu zasłony gdy widzisz jej skutek",
              "Głośno, wcześnie — 3 razy PRZED zasłoną",
              "Sygnałami dłońmi tylko"],
             2,
             "'Głośno, wcześnie, 3 razy PRZED zasłoną!' "
             "Gdy widzisz że screener się ustawia → KRZYCZ zanim zasłona stoi! "
             "Kolega ma czas na wybór jak bronić. Po zasłonie jest za późno."),
        ]
    ),

    (
        "Rzut za 3 po zasłonie — technika i czytanie",
        "taktyka", 6, "U18", 504,
        """W U16–U18 dodajemy rzut za 3 po koźle i po zasłonach. To umiejętność pro!

Typy rzutów za 3:
• SPOT UP — stoisz i strzelasz z otwartego miejsca (base)
• OFF THE DRIBBLE (catch & shoot) — złapiesz i natychmiast strzelasz
• PULL UP — wychodzisz z kozłowania i strzelasz
• OFF SCREEN — wybiegasz z zasłony i strzelasz

Technika catch & shoot:
1. Widzisz podanie zanim do ciebie doleci → stopy już w pozycji rzutowej
2. Piłka leci → READY POSITION (stopy gotowe, kolana lekko ugięte)
3. Złap → rzut w jednym ruchu — bez 'dipping' (opuszczania piłki)
4. Follow-through: ręka wyciągnięta do kosza, palce wskazują dół

Zasada wyboru rzutu:
• Masz OTWARTĄ trójkę i jesteś w swoim zasięgu → bierz ZAWSZE
• Zamknięty obrońca blisko → nie bierz
• Żalgiris: 'strzelaj layupy, wolne i otwarte trójki — reszta to anty-shot'""",
        [
            ("Co to jest 'catch & shoot' (catch and shoot)?",
             ["Złap piłkę i kozłuj do strzału",
              "Złap piłkę i od razu strzelaj bez kozłowania",
              "Złap piłkę jedną ręką i rzuć",
              "Złap zasłonę i strzelaj"],
             1,
             "Catch & shoot = złap i strzelaj w JEDNYM RUCHU! "
             "Bez 'dipping' (opuszczania piłki na dół), bez kozłowania. "
             "Szybkość decyzji jest kluczowa — obrońca nie ma czasu na zbliżenie."),
            ("Przed rzutem za 3 po zasłonie (off screen) ważne jest...",
             ["Zwolnienie przy zasłonie",
              "Wyjście z zasłony z prędkością i gotowością do rzutu (żeby złapać i rzucać)",
              "Szukanie obrońcy",
              "Zatrzymanie się przy zasłonie"],
             1,
             "Kluczowe: TEMPO! Wychodzisz z zasłony szybko → obrońca walczy żeby wyjść → "
             "Ty już jesteś gotowy do rzutu. Spowolnienie przy zasłonie daje obrońcy czas na wyjście."),
            ("Wg Żalgirisu, jakie rzuty powinieneś BRAĆ?",
             ["Zawsze rzut gdy masz piłkę",
              "Tylko rzuty z bardzo bliskiej odległości",
              "Layupy, rzuty wolne i OTWARTE trójki — reszta to anty-shot",
              "Rzuty z środka boiska zawsze"],
             2,
             "'Best possible shot, not the first available!' "
             "Żalgiris shot selection: layup (2 pts easy) + wolny (wymuś faul) + otwarta 3PT. "
             "Trudna dwójka ze środka = anty-shot. Dużo prób → mała skuteczność."),
        ]
    ),

    (
        "Lokalizacja kontroli — kowal swego losu",
        "mentalnosc", 6, "U18", 505,
        """Najważniejsza cecha mentalna przyszłego profesjonalisty — locus of control.

Co to locus of control (LOC)?
Badania prof. Rottera: przekonanie o tym, kto/co kontroluje Twoje życie.

WEWNĘTRZNE LOC (internal):
• 'Ja sam decyduję o tym co osiągam'
• Po porażce: 'Co mogę zrobić inaczej?'
• Szukasz ALTERNATYWNYCH wariantów
• Aktywny w trudnych sytuacjach

ZEWNĘTRZNE LOC (external):
• 'Sędzia, trener, szczęście decydują'
• Po porażce: 'Mieliśmy pecha, sędzia był zły'
• Pasywny — 'co ma być to będzie'
• Wolisz być prowadzony

Co mówią badania sportowe?
• Zawodnicy wyższej klasy mają bardziej wewnętrzne LOC
• JUŻ po 2 MIESIĄCACH regularnego treningu LOC przesuwa się ku wewnętrznemu!
• LOC jest TRENOWALNY — możesz nad nim pracować

Dla zawodnika U16–U18 (z BasketKołcz KB):
'Kowal swego losu' = wewnętrzne LOC = klucz do kariery zawodowej.""",
        [
            ("Zawodnik z WEWNĘTRZNYM locus of control po porażce myśli...",
             ["'Mieliśmy pecha — sędzia nas skrzywdził'",
              "'Co mogę zrobić inaczej żeby osiągnąć lepszy wynik?'",
              "'Trener powinien inaczej ustawić składy'",
              "'Wszyscy grali źle, nie tylko ja'"],
             1,
             "Wewnętrzne LOC = szukam CO MOGĘ ZMIENIĆ. "
             "Nie szukam winnych na zewnątrz — szukam rozwiązań wewnątrz. "
             "To jest mentalna podstawa każdego profesjonalnego zawodnika."),
            ("Co mówią badania sportowe o LOC i poziomie osiągnięć?",
             ["Zewnętrzne LOC jest częstsze u lepszych zawodników",
              "LOC nie ma wpływu na osiągnięcia sportowe",
              "Zawodnicy wyższej klasy mają bardziej wewnętrzne LOC",
              "LOC jest wrodzone i nie można go zmienić"],
             2,
             "Badania Pawłowskiej (polska literatura): wewnętrzne LOC różnicuje poziom osiągnięć. "
             "Lepsi zawodnicy = bardziej wewnętrzne LOC. "
             "Co ważne: JUŻ po 2 miesiącach regularnego treningu LOC przesuwa się ku wewnętrznemu!"),
            ("LOC (locus of control) to...",
             ["Lokalizacja treningu (indoor vs outdoor)",
              "Przekonanie o tym, kto kontroluje twoje wyniki i życie",
              "Rodzaj motivacji do ćwiczeń",
              "Technika koncentracji przed meczem"],
             1,
             "LOC = przekonanie o źródle kontroli nad Twoim życiem. "
             "Wewnętrzne: 'Ja' kontroluję. Zewnętrzne: 'Środowisko' kontroluje. "
             "To psychologiczny fundament odpowiedzialności i akceptacji roli."),
            ("Dlaczego LOC jest ważny w przejściu do seniora?",
             ["Bo seniorzy mają inne przepisy",
              "Bo decyzje trenera i organizacji są poza Twoją kontrolą — wewnętrzne LOC pomaga skupić się na tym, co możesz kontrolować",
              "Bo fizyczność zastępuje mentalność w seniorze",
              "Bo nie ma już U18 ligi"],
             1,
             "W seniorze: trener cię nie wybiera, dostaniesz mniej minut, rola jest mniejsza. "
             "Zawodnik z zewnętrznym LOC → frustracja, obwinia innych. "
             "Zawodnik z wewnętrznym LOC → 'co mogę zrobić żeby dostać więcej minut?' "
             "LOC decyduje jak radzisz sobie z przejściem junior → senior."),
        ]
    ),

    (
        "Myślenie długoterminowe — kariera to maraton",
        "mentalnosc", 6, "U18", 506,
        """U18 to decydujący czas: decydujesz o KARIERZE, nie tylko o sezonie.

10 prawd o karierze zawodnika (z BasketKołcz Knowledge Base):

1. TALENT + CHARAKTER — sam talent nie wystarczy. Charakter decyduje 'jak daleko dojdziesz'
2. PRZEWIDYWALNOŚĆ — elitarni zawodnicy są KONSEKWENTNI, nie 'good days / bad days'
3. GRA BEZ PIŁKI — 80% czasu spędzasz bez piłki. Ta 80% decyduje czy grasz
4. REAKCJA NA BŁĄD — jak reagujesz po błędzie jest ważniejsze niż błąd sam
5. REPUTACJA — rynki koszykarskie są małe. Wszyscy wiedzą o TOBIE
6. ROLA vs STATYSTYKI — wygrywa ten kto akceptuje swoją rolę i jest w niej doskonały
7. OBRONA I DECYZJE — trenerzy seniorzy zatrudniają OBROŃCÓW z decyzjami
8. PROFESJONALIZM — codzienne nawyki: dieta, sen, praca indywidualna
9. COACHABILITY — czy słuchasz i uczysz się, czy bronisz się przed feedbackiem?
10. MYŚLENIE DŁUGOTERMINOWE — 1 sezon to nie kariera. Buduj 5-10-letni plan

Piramida kariery wg basektKołcz KB:
U14-16 fundamenty → U17-19 specjalizacja → Senior: translacja""",
        [
            ("Jak Żalgiris definiuje 'coachability'?",
             ["Mierzenie IQ zawodnika",
              "Czy słuchasz feedbacku i uczysz się, zamiast bronić swojego ego",
              "Czy dobrze grasz pod konkretnego trenera",
              "Zdolność do gry w różnych systemach"],
             1,
             "'Coachability' to jedna z 5 wartości Żalgirisu! "
             "Zawodnik nie-coachable: słyszy feedback → broni się → nie rośnie. "
             "Zawodnik coachable: słyszy feedback → analizuje → implementuje → rośnie szybciej."),
            ("Dlaczego 'gra bez piłki' jest tak ważna dla trenera seniora?",
             ["Bo w seniorze jest mniej posiadań na zawodnika",
              "Bo 80% czasu spędzasz bez piłki — ta 80% decyduje czy dostajesz minuty",
              "Bo w seniorze nie można kozłować tak dużo",
              "Bo trener chce zobaczyć stamina"],
             1,
             "80% czasu w meczu = bez piłki! Trener seniora patrzy: czy biegniesz? "
             "Czy tworzysz spacing? Czy robisz backdoor? Czy zastawiasz? "
             "To decyduje o minutach — nie tylko te 20% z piłką."),
            ("'Piramida kariery' Żalgirisu/KB mówi, że U14-U16 to czas na...",
             ["Specjalizację w jednej pozycji",
              "Translację umiejętności na grę seniorską",
              "Budowanie fundamentów technicznych i mentalnych",
              "Myślenie o transferze do NBA"],
             2,
             "Piramida: U14-16 = FUNDAMENTY (technika, taktyka, charakter) → "
             "U17-19 = SPECJALIZACJA (twoja rola, twój archetyp) → "
             "Senior = TRANSLACJA (czy to co umiesz 'przenosi się' na wyższy poziom). "
             "Fundament decyduje o szczycie piramidy!"),
        ]
    ),

]

# ──────────────────────────────────────────────────────────────────────────────
# HELPER
# ──────────────────────────────────────────────────────────────────────────────

AGE_LABELS = {
    "U8":  "U6–U8",
    "U10": "U9–U10",
    "U13": "U11–U13",
    "U15": "U14–U15",
    "U18": "U16–U18",
    "ALL": "Ogólne",
}


def connect():
    return psycopg2.connect(DB_URL, cursor_factory=psycopg2.extras.RealDictCursor)


def insert_curriculum(dry_run=False):
    conn = connect()
    cur = conn.cursor()

    inserted_lessons = 0
    inserted_questions = 0
    skipped = 0

    for (title, category, duration_min, pos_focus, order_num, content, quizzes) in CURRICULUM:
        cur.execute("SELECT id FROM player_lessons WHERE title=%s", (title,))
        existing = cur.fetchone()
        if existing:
            print(f"  [SKIP] '{title}' już istnieje (id={existing['id']})")
            skipped += 1
            continue

        age_label = AGE_LABELS.get(pos_focus, pos_focus)
        if not dry_run:
            cur.execute("""
                INSERT INTO player_lessons
                    (title, category, content, video_url, position_focus, stat_trigger,
                     order_num, active, duration_min)
                VALUES (%s, %s, %s, '', %s, '', %s, true, %s)
                RETURNING id
            """, (title, category, content, pos_focus, order_num, duration_min))
            lid = cur.fetchone()["id"]
        else:
            lid = f"DRY-{order_num}"

        print(f"  [OK] ({age_label}) '{title}' → id={lid}")
        inserted_lessons += 1

        for q_idx, (question, options, correct_idx, explanation) in enumerate(quizzes):
            if not dry_run:
                cur.execute("""
                    INSERT INTO player_quiz_questions
                        (lesson_id, question, options_json, correct_idx, explanation, order_num, active)
                    VALUES (%s, %s, %s, %s, %s, %s, false)
                """, (lid, question, json.dumps(options, ensure_ascii=False),
                      correct_idx, explanation, q_idx + 1))
            inserted_questions += 1

    if not dry_run:
        conn.commit()

    cur.close()
    conn.close()

    print(f"\n{'[DRY RUN] ' if dry_run else ''}Gotowe!")
    print(f"  Lekcji wstawiono: {inserted_lessons}")
    print(f"  Pytań wstawiono:  {inserted_questions}")
    print(f"  Pominięto (duplikaty): {skipped}")


def print_summary():
    """Wypisuje podsumowanie curriculum bez łączenia z bazą."""
    from collections import defaultdict
    groups = defaultdict(list)
    for item in CURRICULUM:
        title, cat, dur, pos, order, *_ = item
        groups[pos].append((order, title, cat, dur))

    print("\n=== PODSUMOWANIE CURRICULUM ===\n")
    total_lessons = 0
    total_q = 0
    for pos in ["U8", "U10", "U13", "U15", "U18"]:
        label = AGE_LABELS[pos]
        items = sorted(groups.get(pos, []))
        n_q = sum(len(CURRICULUM[i][6]) for i in range(len(CURRICULUM))
                  if CURRICULUM[i][3] == pos)
        print(f"  {label} ({len(items)} lekcji, {n_q} pytań):")
        for order, title, cat, dur in items:
            print(f"    [{order}] [{cat}] {title} ({dur} min)")
        total_lessons += len(items)
        total_q += n_q
        print()
    print(f"  ŁĄCZNIE: {total_lessons} nowych lekcji, {total_q} pytań quizów")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Seed curriculum Żalgiris do bazy")
    parser.add_argument("--dry-run", action="store_true",
                        help="Pokazuje co zostanie wstawione bez modyfikacji bazy")
    parser.add_argument("--summary", action="store_true",
                        help="Wypisuje podsumowanie curriculum (bez DB)")
    args = parser.parse_args()

    if args.summary:
        print_summary()
    else:
        print(f"Łączę z: {DB_URL}")
        print(f"Tryb: {'DRY RUN' if args.dry_run else 'WRITE'}\n")
        insert_curriculum(dry_run=args.dry_run)
