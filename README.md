# AI Kuchařka

[![CI](https://github.com/sadda/ai-kucharka/actions/workflows/ci.yml/badge.svg)](https://github.com/sadda/ai-kucharka/actions/workflows/ci.yml)

Python Crawler pro automatické zpracování dokumentů Word a Excel uložených ve složkách zakázek. Prochází jednotlivé zakázky, vyhledává podporované dokumenty a získává z nich předem definované informace. Zpracování lze přizpůsobit pro různé společnosti.

## 📖 [Dokumentace](https://sadda.github.io/ai-kucharka/)

## Spuštění

```bash
python convert.py                                      # výchozí konfigurace configs/company123.toml
python convert.py --config configs/moje.local.toml     # vlastní konfigurace
python convert.py ZAK001 ZAK002                        # pouze vybrané zakázky
```

Na Windows lze stejné argumenty předat skriptu `.\convert.ps1`. Vlastní konfiguraci vytvořte zkopírováním `configs/company123.toml` do `configs/<název>.local.toml`; tyto soubory git ignoruje.
