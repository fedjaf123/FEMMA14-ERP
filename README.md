# FEMMA14-ERP

Aplikacija za obradu kalkulacija, prometa i zaliha sa fokusom na precizne nabavne i prodajne cijene.

## Ključne funkcije
- Učitavanje i provjera kalkulacija (xlsx) po sheetovima.
- Inkrementalni upis kalkulacija sa provjerom uspješnog upisa SKU‑ova.
- Proračun prometa i zaliha sa izlaznim izvještajem.
- Historija artikla po SKU (kalkulacije) u UI‑u.

## Kako pokrenuti
1. Instaliraj Python 3.13.
2. Instaliraj zavisnosti (ako nisu instalirane):

```powershell
pip install -r requirements.txt
```

3. Pokreni aplikaciju:

```powershell
python FEMMA14.0.py
```

## Folderi i fajlovi
- `Kalkulacije/` – Excel kalkulacije.
- `nabavnecijene_kalkulacije.json` – agregirane cijene iz kalkulacija.
- `kalk_file_cache.json` – cache po fajlu za inkrementalni upis.
- `db_backup/` – backup baze.

## Napomena
- Ako kalkulacija bude preskočena, aplikacija će prikazati razlog i fajl će biti ponuđen ponovo.

