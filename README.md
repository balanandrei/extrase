# Extragere extrase de cont Trezorerie (PDF → Excel + DBF + XML)

Scriptul `extract_extrase.py` ia PDF-urile de extras de cont (cele semnate,
`TREZ321_ExtrasEP_PDFCLI_<CUI>_XML_SIGNED_*.pdf`), extrage XML-ul din interiorul
fiecărui PDF, le unește și produce pentru fiecare folder de intrare:

- un fișier **Excel** (`.xlsx`) — sumele sunt numere reale, afișate cu 2 zecimale
  (ex. `17226.00`, `530641.50`), deci se pot aduna;
- un fișier **DBF** (`.dbf`);
- un **XML concatenat** (toate extrasele într-un singur fișier);
- folderul `xml_extrase/` cu XML-ul extras din fiecare PDF.

Nu are nevoie de biblioteci externe — doar Python 3.

---

## Configurare (`config.toml`)

Deschide `config.toml` cu Notepad și pune folderele tale:

```toml
input_folders = [
    "C:\\Users\\tata\\Documents\\EXTRASE MAI 2026",
    "C:\\Users\\tata\\Documents\\EXTRASE IUNIE 2026",
]
output_folder = "C:\\Users\\tata\\Documents\\Extrase_rezultate"
```

- `input_folders` — unul sau mai multe foldere cu PDF-uri. **Fiecare folder
  produce propriul set de fișiere** (Excel + DBF + XML), într-un subfolder cu
  numele lui, în `output_folder`.
- Pe Windows folosește `\\` (dublu backslash) în căi, cum e mai sus.
- CUI-ul primăriei se detectează automat din numele fișierului, deci merge pentru
  orice primărie fără alte modificări.

---

## Cum se rulează pe Windows (pas cu pas pentru tata)

### 1. Instalează Python (o singură dată)

1. Intră pe <https://www.python.org/downloads/> și apasă butonul mare
   **„Download Python”**.
2. Rulează fișierul descărcat. **FOARTE IMPORTANT:** bifează căsuța
   **„Add Python to PATH”** din partea de jos a primei ferestre, apoi apasă
   **„Install Now”**.
3. Așteaptă să termine și apasă **Close**.

### 2. Pune fișierele la un loc

Copiază în același folder (ex. pe Desktop, un folder numit `Extrase`):
- `extract_extrase.py`
- `config.toml`
- `ruleaza_windows.bat`

### 3. Editează `config.toml`

Click-dreapta pe `config.toml` → **Deschide cu → Notepad**. Schimbă
`input_folders` și `output_folder` cu folderele tale (vezi mai sus). Salvează
(Ctrl+S) și închide.

### 4. Rulează

**Dublu-click pe `ruleaza_windows.bat`.**

Se deschide o fereastră neagră, apar mesaje despre câte PDF-uri au fost procesate,
iar la final scrie **„Gata!”**. Fișierele Excel și DBF sunt în folderul de output
pe care l-ai pus în config. Apasă o tastă ca să închizi fereastra.

> Dacă apare „python nu este recunoscut...”, înseamnă că la pasul 1 nu a fost
> bifat **„Add Python to PATH”**. Dezinstalează Python și reinstalează bifând
> căsuța.

---

## Rulare din linia de comandă (opțional, pentru avansați)

```
python extract_extrase.py                 # foloseste config.toml de langa script
python extract_extrase.py alt_config.toml # foloseste alt fisier de config
```

Pe Mac/Linux: `python3` în loc de `python`.

---

## Ce se întâmplă dacă un PDF e stricat?

Scriptul îl raportează ca `ESUAT`, **continuă** cu restul fișierelor și afișează
la final lista celor eșuate. Un singur PDF problematic nu oprește tot procesul.
