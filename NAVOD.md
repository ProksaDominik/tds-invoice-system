# Správa faktúr – TDS s.r.o.

Lokálna aplikácia na automatické spracovanie faktúr: prečíta PDF faktúru,
rozozná, či je **prijatá** (my platíme) alebo **vystavená** (nám platia),
zapíše údaje do Excel tabuliek a vytvorí udalosti v Google Kalendári:
udalosť v deň splatnosti každej faktúry + spoločné pripomienky
**20 a 10 dní pred 30.6.** so zoznamom všetkých faktúr splatných
do tohto termínu.

Všetko beží **lokálne na vašom počítači** – žiadne dáta sa neposielajú na
cudzie servery (okrem vytvorenia udalosti vo vašom Google Kalendári).

---

## 1. Ako to funguje (workflow)

1. **Nahratie** – v aplikácii stlačíte „Nahrať faktúru (PDF)“ a vyberiete
   jeden alebo viac PDF súborov.
2. **Extrakcia** – aplikácia z PDF prečíta: číslo faktúry, dodávateľa,
   odberateľa, IČO, dátum vystavenia, dátum splatnosti, sumy, DPH,
   variabilný symbol. Pri **prijatých** faktúrach (každý dodávateľ má inú
   šablónu) sa najprv skúsia regulárne výrazy; polia, ktoré sa nimi nenájdu,
   sa doplnia cez lokálny AI model (Ollama – pozri bod 2.6) – bežiaci
   priamo na vašom počítači, takže žiadny text z faktúry nikam neodchádza.
3. **Rozlíšenie typu** – podľa toho, kde sa nachádza IČO / názov
   TDS s.r.o.:
   - TDS je **dodávateľ** → faktúra **Vystavená** (iná firma platí nám),
   - TDS je **odberateľ** → faktúra **Prijatá** (platíme my),
   - ak sa to nedá určiť → typ „Neurčené“ a vyberiete ho ručne
     cez „Upraviť“.
4. **Záložná tabuľka** – údaje sa NAJPRV zapíšu do `data/zaloha.xlsx`
   so stavom „Čaká na schválenie“. PDF sa zároveň skopíruje do
   `data/archiv_pdf`.
5. **Kontrola** – v aplikácii záznam skontrolujete:
   - **Upraviť** – opravíte chybne prečítané údaje,
   - **Schváliť** – záznam sa zapíše do hlavnej tabuľky,
   - **Zamietnuť** – záznam sa do hlavnej tabuľky nedostane.
6. **Hlavné tabuľky** – po schválení sa riadok pridá do:
   - `data/faktury_prijate.xlsx` – prijaté faktúry,
   - `data/faktury_vystavene.xlsx` – vystavené faktúry.
   V zálohe riadok ostáva (so stavom „Schválené“) – slúži ako trvalá záloha.
7. **Google Kalendár** – pri schválení vystavenej faktúry sa:
   - vytvorí celodenná udalosť v deň **splatnosti** faktúry,
   - aktualizujú dve **spoločné pripomienky 20 a 10 dní pred 30.6.**
     (t. j. 10.6. a 20.6.) – každá obsahuje zoznam VŠETKÝCH
     vystavených faktúr splatných do najbližšieho 30.6.
     Pri schválení ďalšej faktúry sa zoznam v udalostiach automaticky
     doplní. Termín 30.6. sa opakuje každý rok.
   Upozornenia vám prídu z Google Kalendára (mobil aj počítač).

---

## 2. Inštalácia (raz)

### 2.1 Python

1. Stiahnite Python 3.10+ z <https://www.python.org/downloads/>.
2. Pri inštalácii **zaškrtnite „Add Python to PATH“**.

### 2.2 Knižnice

Otvorte príkazový riadok (cmd) v priečinku aplikácie a spustite:

```
pip install -r requirements.txt
```

### 2.3 Nastavenie firmy – DÔLEŽITÉ

Otvorte `config.json` a doplňte **skutočné IČO firmy TDS s.r.o.**
(teraz je tam vzorové `12345678`):

```json
"ico_firmy": "VAŠE_IČO",
```

Podľa IČO aplikácia rozlišuje prijaté a vystavené faktúry, preto musí
byť správne. Ďalšie voľby v `config.json`:

| Kľúč | Význam |
|---|---|
| `hromadny_termin` | každoročný termín splatnosti pre spoločné pripomienky (predvolené `30.06.`) |
| `hromadne_pripomienky_dni` | koľko dní pred termínom prídu spoločné pripomienky (predvolené 20 a 10) |
| `kalendar_len_vystavene` | `true` = udalosti len pre vystavené faktúry; `false` = aj pre prijaté |
| `kalendar_zapnuty` | `false` = kalendár úplne vypnúť |

### 2.4 Google Calendar API (raz, ~10 minút)

1. Otvorte <https://console.cloud.google.com/> a prihláste sa Google
   účtom, v ktorého kalendári chcete pripomienky.
2. Hore vľavo **vytvorte nový projekt** (napr. „TDS Faktury“).
3. V menu **APIs & Services → Library** vyhľadajte
   **Google Calendar API** a stlačte **Enable**.
4. **APIs & Services → OAuth consent screen**:
   - typ **External**, vyplňte názov aplikácie a váš e-mail,
   - v časti **Test users** pridajte svoj Google e-mail.
5. **APIs & Services → Credentials → Create Credentials →
   OAuth client ID**:
   - typ aplikácie: **Desktop app**,
   - stiahnite JSON a uložte ho ako **`credentials.json`**
     do priečinka aplikácie (vedľa `app.py`).
6. Pri prvom schválení faktúry sa otvorí prehliadač – prihláste sa
   a povoľte prístup. Vytvorí sa `token.json` a odvtedy všetko beží
   automaticky, bez ďalšieho prihlasovania.

### 2.5 AI doplnenie chýbajúcich údajov (Ollama, voliteľné)

Prijaté faktúry od rôznych dodávateľov nemajú spoločnú šablónu, preto
regulárne výrazy niekedy nenájdu všetky údaje. Aplikácia si vtedy vie
pomôcť lokálnym AI modelom cez **Ollama** – bezplatný nástroj, ktorý beží
priamo na vašom počítači (žiadne dáta sa neposielajú von, funguje aj
offline).

1. Stiahnite a nainštalujte Ollama z <https://ollama.com/download>
   (na tomto počítači je už nainštalovaná).
2. Stiahnite model (raz, cez príkazový riadok):
   ```
   ollama pull llama3.1:8b
   ```
3. Ollama beží ako služba na pozadí automaticky po inštalácii/reštarte.
   Aplikácia sa k nej pripája sama – nič ďalšie netreba nastavovať.

V `config.json` je možné doplnenie ovládať:

| Kľúč | Význam |
|---|---|
| `ollama_zapnute` | `false` = AI doplnenie úplne vypnúť (iba regulárne výrazy) |
| `ollama_model` | názov modelu, ktorý sa má použiť (predvolené `llama3.1:8b`) |
| `ollama_pokusy` | koľkokrát sa má AI skúsiť opýtať na stále chýbajúce polia (predvolené 3  malý model niekedy nájde pole až na 2. či 3. pokus) |

Ak Ollama nebeží alebo model nie je stiahnutý, aplikácia to potichu
preskočí a použije len to, čo našli regulárne výrazy – nič sa nepokazí.

Aplikácia navyše AI vždy spätne **prekontroluje** už nájdené sumy a
zákazku (nezávisle od regexu) – ak sa AI a regex nezhodnú, do poľa
„Poznámka“ sa pridá upozornenie, aby ste si to pri kontrole vedeli
overiť namiesto toho, aby tichá chyba prešla bez povšimnutia.

### 2.6 Vizuálna AI záchrana (voliteľné, pomalšie)

Text vytiahnutý z PDF je pri niektorých faktúrach potichu poškodený
(napr. chýbajúca číslica v sume alebo rozbité kódovanie písma priamo
v PDF) – aj keď faktúra pri zobrazení vyzerá úplne normálne. Ani
najlepší textový AI model takú chybu neodhalí, keďže dostane na vstup
už poškodený text. Preto sa pre tú menšinu prijatých faktúr, kde
niektorá suma stále chýba (alebo kde vyššie uvedená AI kontrola nahlási
nezhodu), skúsi ešte jedna posledná záchrana: **vizuálny** AI model
(`qwen2.5vl`), ktorý si sumy prečíta priamo z OBRÁZKA stránky, presne
tak, ako ich vidí človek.

Toto je výrazne pomalšie ako bežné textové doplnenie (rádovo 2–4 minúty
na jednu faktúru, podľa výkonu grafickej karty), preto sa nevolá na
každú faktúru – iba na tú malú časť, kde všetko ostatné zlyhalo.

1. Stiahnite model (raz, cez príkazový riadok):
   ```
   ollama pull qwen2.5vl:7b
   ```
2. V `config.json`:

| Kľúč | Význam |
|---|---|
| `vision_zapnute` | `false` = vizuálnu AI záchranu úplne vypnúť |
| `vision_model` | názov modelu (predvolené `qwen2.5vl:7b`) |

---

## 2.6 Presun na iný počítač

1. Skopírujte **celý priečinok aplikácie** (vrátane `data/`, `tessdata/`,
   `config.json`, `icon.png`/`icon.ico` a `credentials.json`/`token.json`
   ak už existujú) na nový počítač  napr. cez USB kľúč, OneDrive alebo zdieľaný disk.
   Priečinok `__pycache__` netreba kopírovať (vytvorí sa sám).
2. Na novom počítači spustite **`instalacia.ps1`** (klik pravým tlačidlom →
   „Spustiť s PowerShell"). Skript automaticky:
   - doinštaluje Python knižnice z `requirements.txt`,
   - nainštaluje Tesseract OCR (ak chýba)  jazykové dáta (slovenčina)
     už sú súčasťou skopírovaného priečinka `tessdata/`, netreba ich sťahovať znova,
   - vytvorí odkaz **„Sprava faktur TDS"** na ploche s ikonou aplikácie.
3. Ak Python ešte nie je nainštalovaný, skript vás na to upozorní 
   nainštalujte ho podľa bodu 2.1 a skript spustite znova.
4. Ak používate Google Kalendár a skopírovali ste aj `credentials.json`,
   pri prvom schválení faktúry na novom počítači sa otvorí prehliadač na
   prihlásenie (nové `token.json` sa vytvorí automaticky).

---

## 3. Denné používanie

1. Kliknite na odkaz **„Sprava faktur TDS“** na ploche (vytvorený cez
   `instalacia.ps1`), alebo dvojklik na `start.bat` (alebo `python app.py`).
2. **Nahrať faktúru (PDF)** → vyberte faktúry.
3. Skontrolujte riadky v tabuľke, prípadne **Upraviť**.
4. **Schváliť** → zápis do hlavnej tabuľky + udalosť v kalendári.

---

## 4. Súbory

| Súbor / priečinok | Účel |
|---|---|
| `app.py` | hlavná aplikácia (okno) |
| `extractor.py` | čítanie údajov z PDF |
| `excel_manager.py` | zápis do Excel tabuliek |
| `calendar_manager.py` | Google Kalendár |
| `config.json` | nastavenia (IČO, pripomienky…) |
| `data/zaloha.xlsx` | záložná tabuľka (kontrola pred schválením) |
| `data/faktury_prijate.xlsx` | hlavná tabuľka – prijaté faktúry |
| `data/faktury_vystavene.xlsx` | hlavná tabuľka – vystavené faktúry |
| `data/archiv_pdf/` | kópie všetkých nahratých PDF |
| `credentials.json` | prístup ku Google API (vytvoríte podľa 2.4) |
| `token.json` | uložené prihlásenie (vytvorí sa samo) |
| `tessdata/` | jazykové dáta pre OCR (slovenčina + angličtina) |
| `icon.png` / `icon.ico` | ikona aplikácie (okno, ikona na ploche) |
| `instalacia.ps1` | automatická inštalácia na novom počítači |
| `diagram_workflow.svg` | grafický diagram celého procesu |

---

## 5. Riešenie problémov

**„Neurčené“ pri type faktúry** – v PDF sa nenašlo IČO ani názov TDS.
Skontrolujte `ico_firmy` v `config.json`; typ vyberte ručne cez „Upraviť“.

**Prázdne polia po nahratí** – faktúra má netypické rozloženie. Skenované/
vytlačené a znova naskenované faktúry sa čítajú cez OCR (Tesseract,
automaticky), no pri nezvyčajnom rozložení môže byť rozpoznávanie nepresné –
údaje doplňte cez „Upraviť“.

**Kalendár zlyhal** – skontrolujte, že `credentials.json` je v priečinku
aplikácie a že ste svoj e-mail pridali medzi Test users (bod 2.4).
Ak sa zmenil Google účet, vymažte `token.json` a schváľte faktúru znova.

**Excel súbor sa nedá uložiť** – zavrite ho v Exceli; aplikácia doň
nemôže zapisovať, kým je otvorený.
