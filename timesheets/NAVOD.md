# Dochádzka – TDS s.r.o.

Aplikácia, ktorá naskenované papierové **PRACOVNÉ VÝKAZY** (mesačná
dochádzka pracovníka podľa zákaziek) premení na Excel tabuľku: spočíta
odpracované hodiny pri každej zákazke, vynásobí hodinovou sadzbou a
vypočíta sumu v €. Celá tabuľka aj sadzba ostávajú v Exceli **plne
upraviteľné** – pre prípad, že program niečo prečíta zle.

Rozpoznávanie rukopisu beží cez **Gemini API** (Google, bezplatný model
`gemini-flash-latest`) – potrebuje internet a bezplatný API kľúč. Ak
Gemini zlyhá (napr. vyčerpaná denná kvóta), aplikácia sama skúsi
**Azure AI Document Intelligence** (voliteľné, tiež má bezplatnú úroveň
– pozri bod 4b) a až úplne nakoniec lokálny Tesseract (funguje offline,
ale je najmenej presný na rukopise).

---

## 1. Prečo treba všetko kontrolovať

Výkazy sú **ručne písané**, nie strojovo vyplnené. Automatické
rozpoznávanie rukopisu (OCR/AI) nikdy nebude 100 % presné – aplikácia
preto:

1. Prečíta všetko naraz cez Gemini (meno, mesiac, rok, všetky riadky aj
   hodiny) – v testoch na reálnom výkaze rozpoznalo správne všetkých 9
   čísel zákaziek aj takmer všetky hodiny, čo lokálny Tesseract
   (bez Gemini) vôbec nezvládal,
2. **VŽDY** zobrazí celý výkaz na obrazovke ako upraviteľnú tabuľku
   pred zápisom do Excelu,
3. tlačidlom **„🔍 Zobraziť naskenovanú stranu"** viete kedykoľvek
   otvoriť originálny sken vedľa seba a hodnoty si overiť/opraviť.

Aj tak vždy skontrolujte najmä čísla zákaziek a hodiny pri hustejšie
popísaných riadkoch.

---

## 2. Ako to funguje (workflow)

1. **„📄 Načítať výkaz (PDF)"** – vyberte jeden alebo viac naskenovaných
   výkazov. Aplikácia z každej strany rozpozná otočenie, meno/mesiac/rok
   a hodiny pri jednotlivých zákazkách (stĺpce deň 1-31).
2. Výkaz sa objaví v zozname vľavo ako **„Čaká na kontrolu"**. Kliknutím
   naň sa vpravo zobrazí celá tabuľka.
3. **Skontrolujte a opravte** hodnoty (porovnajte so skenom cez 🔍).
   Ak program vynechal celý riadok zákazky, pridajte ho tlačidlom
   **„➕ Pridať riadok zákazky"**.
4. **„✅ Schváliť"** – zapíše výkaz do `data/dochadzka.xlsx`
   (hárok **Záznamy**) a prepočíta súhrn za každú zákazku. Pôvodné PDF
   sa presunie do `vykazy_spracovane/`.
5. **„📊 Otvoriť súhrn (Excel)"** – otvorí `data/dochadzka.xlsx` s
   hárkom **Súhrn podľa zákazok** (spolu hodín × hodinová sadzba = €).

Jeden PDF súbor = jeden pracovník za jeden mesiac (aj keby mal viac
strán, lebo mal viac ako 9 zákaziek). Ak naskenujete viacerých
pracovníkov naraz, uložte ich ako samostatné PDF súbory.

Spracovanie jednej strany cez Gemini trvá cca 30-70 sekúnd (posiela sa
cez internet) – pri viacstranovom výkaze teda pokojne aj pár minút.
Aplikácia počas toho zostáva použiteľná (beží na pozadí, s ukazovateľom
priebehu).

---

## 3. Zošit `data/dochadzka.xlsx`

| Hárok | Obsah |
|---|---|
| **Nastavenia** | Predvolená hodinová sadzba (€) – zmeňte tu, prejaví sa všade, kde nebola sadzba ručne prepísaná pre konkrétnu zákazku. |
| **Záznamy** | Surové dáta – jeden riadok = pracovník × mesiac × rok × zákazka, so stĺpcami Deň 1-31. **Toto je zdroj pravdy** – oprava tu (alebo pred schválením v appke) sa premietne do súhrnu. |
| **Súhrn podľa zákazok** | Jeden riadok na zákazku: spolu hodín (automaticky, zo Záznamov), hodinová sadzba (predvolene odkaz na Nastavenia, dá sa pre ktorúkoľvek zákazku ručne prepísať na inú sumu), spolu v €. |

Stĺpce **„Spolu hodín"** a **„Spolu v EUR"** v hárku Súhrn sa pri každom
ďalšom schválení výkazu prepočítajú nanovo – needitujte ich ručne
(zmena by sa stratila). Stĺpec **„Hodinová sadzba"** je naopak presne
na to, aby ste si ho upravovali, ako potrebujete.

Ak chcete výkaz schváliť znova (napr. po väčšej oprave) a jeho zdrojový
súbor už v Záznamoch je, aplikácia sa opýta, či staré záznamy nahradiť,
alebo pridať ako ďalšie (pozor na duplicity pri „pridať").

---

## 4. Gemini API kľúč (raz, ~2 minúty)

1. Otvorte <https://aistudio.google.com/apikey> a prihláste sa Google
   účtom (platobná karta sa nevyžaduje – ide o bezplatnú úroveň).
2. Kliknite **„Create API key"** a skopírujte vygenerovaný kľúč.
3. Vložte ho ako jediný riadok do súboru **`gemini_api_key.txt`**
   v priečinku aplikácie (vedľa `app.py`).

Bezplatná úroveň má obmedzený počet požiadaviek za deň/minútu, no pri
bežnom objeme (rádovo 15-20 výkazov mesačne) je táto hranica prakticky
nedosiahnuteľná. Tento kľúč je súkromný prístupový údaj – nezdieľajte
súbor `gemini_api_key.txt` s nikým a nevkladajte ho do verejných miest.

**Pozor:** na bezplatnej úrovni niektoré konkrétne (dátumovo/verzovo
pomenované) modely majú nulovú kvótu a vrátia chybu 429 hneď pri prvom
volaní – aplikácia je preto predvolene nastavená na model
`gemini-flash-latest` (alias na aktuálne dostupný bezplatný model),
ktorý funguje spoľahlivo. Ak by ste chceli iný model, zmeňte
`gemini_model` v `config.json`, ale najprv si overte, že váš účet naň
má nenulovú kvótu (chybová hláška z Google to vždy uvedie).

Bez tohto kľúča (alebo bez pripojenia na internet) aplikácia automaticky
skúsi Azure Document Intelligence (bod 4b), ak je nastavený, inak rovno
lokálny Tesseract – funguje, ale menej presne (viac oranžových/
nerozpoznaných políčok na ručnú kontrolu).

---

## 4b. Azure Document Intelligence kľúč (voliteľné, záložná cloudová cesta)

Táto cesta sa použije len vtedy, keď Gemini pre danú stranu zlyhá
(napr. vyčerpaná kvóta) – je teda voliteľná, appka funguje aj bez nej
(rovno preskočí na Tesseract).

1. Otvorte <https://portal.azure.com>, prihláste sa (platobná karta sa
   pri bezplatnej úrovni F0 nevyžaduje) a vytvorte zdroj **„Document
   Intelligence"** (Create a resource → hľadajte „Document
   Intelligence" → cenová úroveň **Free F0**, ~500 strán/mesiac).
2. Po vytvorení zdroja choďte do **„Keys and Endpoint"** – skopírujte
   **Endpoint** (napr. `https://vas-zdroj.cognitiveservices.azure.com/`)
   a **KEY 1**.
3. V `config.json` vyplňte `azure_di_endpoint` skopírovaným endpointom.
4. Vložte kľúč ako jediný riadok do súboru **`azure_di_key.txt`**
   v priečinku aplikácie (vedľa `app.py`) – presný názov súboru určuje
   `azure_di_kluc_subor` v `config.json`.

Tento kľúč je súkromný prístupový údaj – nezdieľajte súbor
`azure_di_key.txt` s nikým a nevkladajte ho do verejných miest. Ak
zostane `azure_di_endpoint` prázdny alebo súbor s kľúčom chýba, appka
Azure jednoducho preskočí bez chyby (ako pri chýbajúcom Gemini kľúči).

---

## 5. Inštalácia (raz)

Python aj všetky potrebné knižnice (openpyxl, ttkbootstrap, pdfplumber,
pytesseract, Pillow, google-genai, azure-ai-documentintelligence) sú na
tomto počítači už nainštalované. Ak spúšťate na inom počítači:

1. Nainštalujte Python 3.10+ (zaškrtnite „Add Python to PATH").
2. Nainštalujte [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
   (jazykové dáta slovenčiny sú súčasťou priečinka `tessdata/` – slúži
   len ako posledná záložná offline cesta, keď Gemini aj Azure zlyhajú).
3. V priečinku aplikácie spustite: `pip install -r requirements.txt`
4. Vytvorte `gemini_api_key.txt` podľa bodu 4 vyššie.
5. Voliteľne vytvorte `azure_di_key.txt` podľa bodu 4b vyššie.

## 6. Denné používanie

Dvojklik na **`start.bat`** (alebo `python app.py`).

---

## 7. Súbory

| Súbor / priečinok | Účel |
|---|---|
| `app.py` | hlavná aplikácia (okno) |
| `extractor.py` | rozpoznanie PDF skenu (otočenie, Gemini/Azure/Tesseract OCR) |
| `template.py` | kalibrované rozloženie tlačového formulára (Azure aj Tesseract záložné cesty ho používajú) |
| `config.json` | nastavenia (sadzba, priečinky, Gemini/Azure prístup) |
| `gemini_api_key.txt` | váš súkromný Gemini API kľúč (vytvoríte podľa bodu 4) |
| `azure_di_key.txt` | voliteľný súkromný Azure Document Intelligence kľúč (vytvoríte podľa bodu 4b) |
| `vykazy/` | sem môžete uložiť naskenované PDF pred načítaním |
| `vykazy_spracovane/` | schválené PDF sa sem presunú (archív) |
| `tessdata/` | jazykové dáta pre záložný lokálny OCR (slovenčina + angličtina) |

---

## 8. Riešenie problémov

**Meno/mesiac/rok sa neprečítali** – ak Gemini beží, toto je zriedkavé;
ak sa to stane (alebo beží záložný lokálny Tesseract), doplňte ich ručne
v hornej časti tabuľky pred schválením.

**Veľa oranžových (nerozpoznaných) políčok** – znamená, že sa použila
záložná lokálna cesta (skontrolujte `gemini_api_key.txt` a pripojenie na
internet). Hodnoty dopíšte ručne podľa 🔍 náhľadu skenu.

**Chyba 429 / „quota exceeded"** – buď je vyčerpaný bezplatný denný
limit (skúste znova nasledujúci deň), alebo `gemini_model`
v `config.json` ukazuje na model bez bezplatnej kvóty na vašom účte
(prepnite späť na `gemini-flash-latest`). Aplikácia sa v takom prípade
automaticky prepne na Azure (ak je nastavený, pozri bod 4b) a až potom
na lokálny Tesseract; presný dôvod uvidíte v žltom/modrom upozornení
nad tabuľkou pri danom výkaze. Pri vyčerpanej kvóte appka na 2 minúty
danú službu úplne preskočí, aby zvyšok dávky zbytočne nečakal na
márne opakované pokusy.

**„⚠️ Over rozloženie"** v zozname – aplikácia si nebola istá, že ide o
štandardné tlačivo (napr. veľmi naklonený sken). Dôkladne skontrolujte
úplne všetky hodnoty.

**Excel súbor sa nedá uložiť** – zavrite `dochadzka.xlsx` v Exceli;
aplikácia doň nemôže zapisovať, kým je otvorený (ponúkne to skúsiť znova).
