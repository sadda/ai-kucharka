# Python Crawler

## 1. Co je Python Crawler

Python Crawler je nástroj pro automatické zpracování dokumentů. Prochází připravené složky, vyhledává podporované dokumenty a získává z nich předem definované informace.

Crawler lze přizpůsobit různým typům dokumentace a způsobům zpracování. Společná část systému zůstává zachována a konkrétní pravidla lze definovat podle daného použití.

Výsledkem jsou strukturovaná data, která lze dále kontrolovat, zpracovávat nebo analyzovat.

---

## 2. Jak crawler funguje

Crawler dostane jako vstup složku s dokumenty. Složka může obsahovat další podsložky a dokumenty mohou být uložené v různých úrovních této struktury.

Crawler složky postupně prochází, vyhledává podporované dokumenty a získává z nich požadované informace.

Výsledek zpracování je uložen do souborů `.pkl` (Pickle). Ty slouží pro uložení zpracovaných dat a nejsou určené k běžnému ručnímu čtení.

Pro kontrolu a další práci s výsledky slouží `analysis.ipynb`, pomocí kterého lze obsah `.pkl` souborů načíst a zobrazit v přehledné podobě.

```text
Složka s dokumenty
        ↓
Procházení složek a podsložek
        ↓
Načtení dokumentů
        ↓
Získání požadovaných dat
        ↓
.pkl
        ↓
analysis.ipynb
        ↓
Zobrazení výsledků
```

<!-- OBRÁZEK: Jednoduché schéma průchodu dat crawlerem -->

---

## 3. Architektura: Core a specializované implementace

Crawler je rozdělen na společnou část `Core` a specializované implementace pro konkrétní typ dokumentace.

### Core

`Core` obsahuje společnou logiku crawleru. Díky tomu není potřeba vytvářet celý crawler znovu pro každé nové použití.

### Specializované implementace

Na `Core` navazují jednotlivé implementace, například:

- `COMTES`
- `CPF`

Ty určují, **co se má v dokumentech hledat a jak se mají nalezené informace zpracovat**.

Zjednodušeně:

```text
              Core
                │
        ┌───────┴───────┐
        ▼               ▼
     COMTES             CPF
        │               │
        ▼               ▼
 vlastní pravidla   vlastní pravidla
```

<!-- OBRÁZEK: Architektura Core → COMTES / CPF -->

---

## 4. Co potřebuji před spuštěním

Pro práci s crawlerem je potřeba:

- **Python**
- **Git**
- **Visual Studio Code** nebo jiný editor
- přístup k repozitáři [AI Kuchařka](https://github.com/sadda/ai-kucharka)
- závislosti potřebné pro běh projektu

Potřebné Python knihovny jsou definované přímo v projektu pomocí souboru `pyproject.toml`.

---

## 5. Instalace krok za krokem

### Stažení projektu

Nejprve naklonujte repozitář:

```bash
git clone <adresa-repozitare>
```

Poté otevřete staženou složku projektu například ve Visual Studio Code.

<!-- OBRÁZEK: Otevřený projekt AI Kuchařka ve Visual Studio Code -->

### Vytvoření virtuálního prostředí

Crawler vyžaduje **Python 3.11 nebo novější**.

Doporučuje se vytvořit samostatné virtuální prostředí:

```bash
python -m venv .venv
```

### Aktivace virtuálního prostředí

Na Windows:

```powershell
.\.venv\Scripts\Activate.ps1
```

Na Linuxu:

```bash
source .venv/bin/activate
```

### Instalace knihoven

Potřebné knihovny jsou definované v souboru `pyproject.toml`.

Po aktivaci virtuálního prostředí je lze nainstalovat společně s projektem:

```bash
pip install -e .
```

Mezi používané knihovny patří například `pandas`, `numpy`, `openpyxl`, `python-docx`, `tqdm` a `xlrd`.

Tím je základní prostředí crawleru připravené.

---

## 6. Konfigurace a příprava vstupních dat

Nastavení crawleru je uloženo v konfiguračním souboru:

```text
configs/company123.toml
```

Pro základní použití jsou důležité především tři hodnoty:

```toml
data_class = "..."
root_input = "C:/Crawler/Vstupni_data"
root_output = "results/Company123"
```

- `data_class` – určuje, **která implementace crawleru se použije** a podle jakých pravidel budou dokumenty zpracovány.
- `root_input` – cesta k hlavní složce obsahující vstupní Word a Excel dokumenty.
- `root_output` – umístění, kam crawler uloží výsledné `.pkl` soubory.

<!-- OBRÁZEK: Ukázka konfiguračního souboru company123.toml -->

### Vstupní data

Do složky nastavené pomocí `root_input` se uloží dokumenty určené ke zpracování. Crawler prochází také vnořené podsložky.

Například:

```text
Vstupni_data/
├── Zakazka_001/
│   ├── Nabidka.docx
│   └── Parametry.xlsx
│
├── Zakazka_002/
│   └── Dokumentace/
│       ├── Specifikace.docx
│       └── Vysledky.xlsx
│
└── Zakazka_003/
    └── Nabidka.docx
```

Konkrétní informace, které crawler z těchto dokumentů získává, závisejí na implementaci nastavené pomocí `data_class`.

<!-- OBRÁZEK: Příklad adresářové struktury vstupních dat -->

---

## 7. První spuštění crawleru

Crawler se spouští pomocí souboru `convert.py`. Spuštění je možné přímo z editoru nebo pomocí terminálu.

### Spuštění z editoru

Ve Visual Studio Code otevřete soubor `convert.py` a spusťte jej pomocí tlačítka **Run / Play**.

Při spuštění bez dalších parametrů se použije výchozí konfigurace:

```text
configs/company123.toml
```

### Spuštění z terminálu

Crawler lze spustit také z terminálu v kořenové složce projektu:

```bash
python convert.py
```

Pro použití jiného konfiguračního souboru:

```bash
python convert.py --config configs/moje.local.toml
```

### Spuštění vybraných složek

Pokud chcete zpracovat pouze vybrané složky ze vstupního adresáře, uveďte jejich názvy za příkazem:

```bash
python convert.py Zakazka_001 Zakazka_002
```

Pokud žádné názvy neuvedete, crawler zpracuje všechny položky ve složce nastavené pomocí `root_input`.

### Průběh zpracování

Během spuštění se v terminálu zobrazují názvy právě zpracovávaných složek a celkový průběh.

<!-- OBRÁZEK: Screenshot úspěšného spuštění crawleru ve VS Code -->

*Příklad úspěšného spuštění crawleru. Hodnota `100 %` znamená, že byly zpracovány všechny nalezené položky.*

Po dokončení jsou výsledné `.pkl` soubory uloženy do složky nastavené pomocí `root_output`.

### Kontrola po spuštění

Po dokončení zkontrolujte:

- zda průběh dosáhl hodnoty `100 %`,
- zda se v terminálu neobjevila chyba,
- zda byly ve výstupní složce vytvořeny očekávané `.pkl` soubory.

---

## 8. Výstupy a jejich použití

Po dokončení zpracování jsou výsledky uloženy do složky nastavené pomocí `root_output`.

Pro každý zpracovaný celek vzniká samostatný soubor `.pkl`.

Například:

```text
results/
└── Company123/
    ├── Zakazka001.pkl
    ├── Zakazka002.pkl
    └── Zakazka003.pkl
```

<!-- OBRÁZEK: Složka results s vytvořenými .pkl soubory -->

### Co je `.pkl`

`.pkl` (Pickle) je formát používaný Pythonem pro ukládání dat. Výsledné soubory nejsou určené k běžnému ručnímu čtení.

Pro další práci s výsledky slouží notebook:

```text
notebooks/analysis.ipynb
```

`analysis.ipynb` načítá vytvořené `.pkl` soubory a umožňuje získaná data dále zpracovat, kontrolovat, analyzovat a zobrazit v přehlednější podobě.

Výsledná data lze podle potřeby připravit také pro další použití, například export do Excelu.

Celý proces lze zjednodušeně znázornit:

```text
Word / Excel
      ↓
   Crawler
      ↓
    .pkl
      ↓
analysis.ipynb
      ↓
Zpracování a analýza výsledků
```

### Kontrola výstupů

Po dokončení crawleru stačí zkontrolovat výstupní složku a ověřit, že pro očekávané položky vznikly odpovídající `.pkl` soubory.

### Analýza výsledků

Pro podrobnější kontrolu a další zpracování se otevře `analysis.ipynb`.

Notebook pracuje s vytvořenými `.pkl` soubory a převádí jejich obsah do podoby vhodné pro další práci s daty.

<!-- OBRÁZEK: Screenshot otevřeného analysis.ipynb a zobrazených výsledků -->

---

## 9. Nejčastější problémy a jejich řešení

### Crawler se nespustí

Ověřte, že je aktivní správné Python prostředí a jsou nainstalované potřebné knihovny.

```bash
pip install -e .
```

### Crawler nenajde vstupní data

Zkontrolujte cestu ke vstupním datům `root_input` v konfiguračním souboru.

```toml
root_input = "C:/Crawler/Vstupni_data"
```

Cesta musí směřovat na existující složku obsahující dokumenty určené ke zpracování.

### Nesprávná struktura vstupních dat

Jednotlivé zpracovávané celky musí být umístěny v samostatných složkách ve vstupním adresáři.

Například:

```text
Vstupni_data/
├── Nabidka_001/
│   ├── Nabidka.docx
│   └── Parametry.xlsx
├── Nabidka_002/
│   └── Nabidka.docx
└── Nabidka_003/
    └── Dokumentace/
        └── Parametry.xlsx
```

Dokumenty mohou být umístěny také ve vnořených podsložkách.

### Nevznikl očekávaný výstup

Zkontrolujte průběh crawleru v terminálu a cestu k výstupním datům `root_output`.

```toml
root_output = "results/Company123"
```

Ve výstupní složce by po úspěšném zpracování měly vzniknout odpovídající `.pkl` soubory.

### Data se nezpracovala znovu

Pokud již pro danou položku existuje `.pkl` soubor, crawler standardně použije existující výsledek.

Pro nové zpracování lze v konfiguraci nastavit:

```toml
overwrite = true
```

### Použita nesprávná implementace

Zkontrolujte hodnotu `data_class` v konfiguračním souboru.

`data_class` určuje implementaci a tím pravidla, podle kterých crawler dokumenty zpracovává.
