## Co je Python Crawler

Python Crawler je aplikace určená k automatickému zpracování dokumentů uložených v připravených složkách.

Crawler prochází jednotlivé složky, vyhledává podporované dokumenty a získává z nich předem definované informace. Výsledná data lze následně využít pro další zpracování, kontrolu nebo analýzu.

Crawler je možné přizpůsobit různým společnostem a způsobům uložení dokumentace. Základní část systému zůstává společná a pro konkrétní společnost lze definovat vlastní způsob zpracování dokumentů a získávání požadovaných údajů.

Crawler aktuálně pracuje s dokumenty Microsoft Word a Microsoft Excel.

Podporované formáty jsou:

- Word: `.doc`, `.docx`, `.docm`
- Excel: `.xls`, `.xlsx`, `.xlsm`

Starší formáty `.doc` a `.xls` jsou před dalším zpracováním automaticky převedeny do novějších formátů.

Výsledkem zpracování jsou strukturovaná data uložená samostatně pro jednotlivé zpracovávané celky. Tyto výsledky lze následně načíst a dále analyzovat.

## Jak funguje

Crawler postupně prochází jednotlivé složky ve vstupním adresáři, vyhledává podporované dokumenty a z jejich obsahu získává požadované informace.

Zpracování lze zjednodušeně popsat následovně:

```text
Vstupní data
    ↓
Vyhledání dokumentů
    ↓
Načtení dokumentů
    ↓
Extrakce informací
    ↓
Uložení výsledků
```

### Core

Crawler je rozdělen na společnou část `Core` a implementaci pro konkrétní společnost.

`Core` obsahuje obecné funkce potřebné pro práci s dokumenty a daty. Zajišťuje například vyhledávání podporovaných souborů, jejich načtení a základní zpracování.

Díky tomu není nutné vytvářet celý crawler znovu pro každou společnost.

### Implementace společnosti

Pro konkrétní společnost lze vytvořit vlastní implementaci navázanou na `Core`.

Tato implementace určuje především:

- jaké dokumenty se mají zpracovávat,
- jaké informace se mají získávat,
- jakým způsobem se mají jednotlivé údaje vyhledávat,
- jaká pravidla se mají při zpracování použít.

Jednodušší údaje lze vyhledávat pomocí předem definovaných klíčových slov. U složitějších údajů lze použít vlastní pravidla podle struktury dokumentu nebo tabulky.

### Načtení a extrakce informací

Po nalezení dokumentů crawler načte jejich obsah a připraví jej pro další zpracování.

U běžných údajů crawler vyhledává předem definovaná označení a následně získává odpovídající hodnoty.

Například pokud dokument obsahuje tabulku:

| Materiál | 1.2343 |
|-----------|---------|

crawler může vyhledat označení `Materiál` a získat hodnotu `1.2343`.

Jeden údaj může mít v různých dokumentech více označení. Například materiál může být uveden jako `Materiál` nebo `Material`. Pro jeden údaj proto může být definováno více možných klíčových slov.

U složitějších údajů nemusí stačit pouze vyhledání názvu a hodnoty. V takovém případě lze použít specifická pravidla, která pracují například se strukturou tabulky.

Po dokončení zpracování jsou získané informace uloženy jako výsledek pro daný zpracovávaný celek.

Podrobnější informace o vstupních datech, výstupech a způsobu přizpůsobení crawleru pro další společnosti jsou uvedeny v dalších částech dokumentace.

## Vstupní data

Před spuštěním crawleru je potřeba připravit dokumenty, které mají být zpracovány, a určit jejich umístění.

Vstupní cesta se nastavuje v souboru `convert.py` pomocí proměnné `root_input`.

```python
root_input = r"C:\Python crawler - test"
```

nebo

```python
root_input = r"D:\Firemni dokumentace"
```

### Testovací data

Před prvním použitím je doporučeno ověřit funkčnost crawleru na menším vzorku dat.

```text
Python crawler - test/
├── Nabidka_001/
├── Nabidka_002/
├── Zkouska_001/
└── Projekt_001/
```

Pro testování by měly být použity pouze fiktivní dokumenty a údaje.

### Struktura dokumentace

Každá hlavní složka představuje jeden zpracovávaný celek, například nabídku, objednávku, projekt nebo zkoušku.

```text
Firemni dokumentace/
├── Nabidka_001/
├── Nabidka_002/
├── Nabidka_003/
└── Zkouska_001/
```

Dokumenty mohou být umístěny přímo ve složce nebo v jejích podsložkách.

```text
Nabidka_001/
├── Nabidka.docx
├── Technicke_udaje.xlsx
└── Prilohy/
    ├── Specifikace.docx
    └── Vysledky.xlsx
```

Crawler při vyhledávání prochází celou strukturu složky včetně podsložek.

### Podporované formáty

Word:

- `.doc`
- `.docx`
- `.docm`

Excel:

- `.xls`
- `.xlsx`
- `.xlsm`

Starší formáty `.doc` a `.xls` jsou před zpracováním automaticky převedeny do novějších formátů.

### Doporučení

Crawler zpracovává všechny podporované dokumenty nalezené ve složce a jejích podsložkách.

```text
Nabidka_001/
├── Nabidka.docx
├── Technicke_udaje.xlsx
└── Archiv/
    └── Stara_nabidka.docx
```


## Spuštění

Crawler je určen pro běh v prostředí Windows a využívá Python. Pro správné spuštění je potřeba připravit Python prostředí, nainstalovat závislosti projektu a následně spustit hlavní skript.

### Požadavky

Pro práci s crawlerem je potřeba:

- Windows
- Python
- Git
- Microsoft Word
- Microsoft Excel
- Visual Studio Code nebo jiný editor
- přístup k repozitáři crawleru

Python je využíván jako hlavní runtime prostředí crawleru. Pro ověření dostupnosti Pythonu lze použít:

```powershell
python --version
```

Pokud příkaz vrátí verzi Pythonu, je Python správně dostupný v systému.

### Stažení projektu

Projekt je uložen v Git repozitáři. Po naklonování repozitáře vznikne lokální kopie projektu, se kterou lze dále pracovat.

Repozitář lze stáhnout například pomocí:

```bash
git clone <adresa-repozitare>
```

Po stažení projektu otevřete jeho kořenovou složku.

### Virtuální prostředí

Pro crawler se používá samostatné Python virtuální prostředí `.venv`.

Virtuální prostředí umožňuje oddělit Python závislosti crawleru od ostatních projektů a aplikací nainstalovaných v počítači.

Vytvoření virtuálního prostředí:

```powershell
python -m venv .venv
```

Po vytvoření vznikne ve složce projektu adresář `.venv`.

Virtuální prostředí je potřeba před spuštěním crawleru aktivovat:

```powershell
.\.venv\Scripts\Activate.ps1
```

Po úspěšné aktivaci se v terminálu obvykle zobrazí název prostředí:

```text
(.venv) PS C:\...\comtes>
```

Od této chvíle se Python a instalované balíčky používají z tohoto virtuálního prostředí.

### Instalace závislostí

Projekt obsahuje soubor `pyproject.toml`, který slouží ke konfiguraci Python projektu a definici jeho závislostí.

Po vytvoření a aktivaci virtuálního prostředí je potřeba nainstalovat závislosti požadované crawlerem.

Konkrétní příkaz pro instalaci závislostí odpovídá způsobu, jakým je projekt nakonfigurován v `pyproject.toml`.

### Nastavení vstupních dat

Před spuštěním crawleru je potřeba připravit vstupní dokumenty.

Umístění vstupních dat se nastavuje v souboru `convert.py` pomocí proměnné `root_input`.

Podrobný popis struktury vstupních dat je uveden v kapitole **Vstupní data**.

### Spuštění crawleru

Po aktivaci virtuálního prostředí a nastavení vstupních dat lze crawler spustit pomocí hlavního Python skriptu:

```bash
python convert.py
```

Crawler následně načte připravená vstupní data a zahájí jejich zpracování.

### Spuštění pomocí PowerShellu

Crawler lze spustit také pomocí připraveného PowerShell skriptu:

```powershell
.\convert.ps1
```

PowerShell skript lze využít pro zjednodušení spuštění a automatizaci jednotlivých kroků potřebných pro běh crawleru.

### Doporučený postup

Při běžném spuštění crawleru postupujte následovně:

1. Otevřete kořenovou složku projektu.
2. Ověřte dostupnost Pythonu.
3. Aktivujte virtuální prostředí `.venv`.
4. Zkontrolujte nainstalované závislosti.
5. Zkontrolujte nastavení `root_input`.
6. Ověřte připravená vstupní data.
7. Spusťte `convert.py` nebo `convert.ps1`.
8. Po dokončení zkontrolujte vytvořené výstupy.

### Řešení běžných problémů

Pokud nelze spustit příkaz `python`, ověřte, zda je Python správně nainstalován a dostupný v systémové proměnné `PATH`.

Pokud nelze aktivovat virtuální prostředí, ověřte, zda byla složka `.venv` vytvořena a zda používáte PowerShell.

Pokud crawler nenajde žádné dokumenty, zkontrolujte hodnotu `root_input` a strukturu vstupních složek podle kapitoly **Vstupní data**.

Pokud při spuštění chybí některý Python modul, je pravděpodobné, že nejsou správně nainstalovány závislosti projektu nebo nebylo aktivováno virtuální prostředí.

Pokud archivní dokumentace nemá být součástí zpracování, doporučuje se ji uložit mimo vstupní adresář crawleru.

## Výstupy

Po spuštění crawleru jsou zpracovaná data uložena do výstupních souborů. Výstupy slouží jako mezikrok mezi samotným zpracováním dokumentů a jejich následným načtením nebo analýzou.

Hlavním výstupem crawleru jsou soubory ve formátu `.pkl`.

### Soubory `.pkl`

Soubor `.pkl` je binární soubor používaný Pythonem pro uložení zpracovaných dat.

Crawler do těchto souborů ukládá informace získané ze vstupních dokumentů. Data jsou uložena ve strukturované podobě, aby s nimi bylo možné dále pracovat v Pythonu nebo v dalších částech projektu.

Výsledný soubor může například obsahovat informace o:

- zpracovávaném dokumentu,
- zákazníkovi,
- zakázce,
- materiálu,
- chemickém složení,
- zkouškách,
- dalších údajích získaných ze vstupní dokumentace.

Obsah výsledných dat závisí na konkrétní implementaci crawleru a na údajích, které jsou pro danou společnost definovány ke zpracování.

### Struktura výstupů

Výstupní soubory jsou vytvářeny samostatně pro jednotlivé zpracovávané celky.

Výsledná struktura může například vypadat následovně:

```text
results/
├── Nabidka_001.pkl
├── Nabidka_002.pkl
├── Nabidka_003.pkl
└── Zkouska_001.pkl
```

Každý `.pkl` soubor obsahuje data získaná ze vstupních dokumentů příslušného zpracovávaného celku.

### Tok dat

```text
Word / Excel dokumenty
          │
          ▼
       Crawler
          │
          ▼
      .pkl soubor
          │
          ▼
   analysis.ipynb
```

### K čemu se `.pkl` soubory používají

`.pkl` soubory slouží především k uložení výsledků crawleru pro další zpracování.

Díky tomu není nutné při každé další analýze znovu procházet původní Word a Excel dokumenty. Výsledná data lze načíst přímo z již vytvořených `.pkl` souborů.

Soubory lze následně využít například:

- pro kontrolu výsledků crawleru,
- pro další zpracování v Pythonu,
- pro analýzu dat,
- jako vstup pro notebooky uložené ve složce `notebooks`,
- pro další části projektu.

### Složka `results`

Výsledné soubory jsou ukládány do složky `results`.

```text
results/
├── Nabidka_001.pkl
├── Nabidka_002.pkl
└── Nabidka_003.pkl
```

Složka `results` tak obsahuje výstupy vytvořené crawlerem a může být použita jako zdroj dat pro následnou analýzu.

### Kontrola výsledků

Po dokončení spuštění je doporučeno zkontrolovat složku `results` a ověřit, zda byly pro zpracovávané celky vytvořeny odpovídající výstupní soubory.

Pokud některý očekávaný výstup chybí, je vhodné zkontrolovat průběh zpracování a případné chybové hlášky.


### Přidání nové společnosti
•	vytvoření implementace 
•	vlastní dokumenty 
•	vlastní atributy 
•	klíčová slova / pravidla 
### Troubleshooting
•	chybějící knihovny 
•	špatná cesta 
•	problémy s dokumenty 
•	běžné chyby při spuštění 
