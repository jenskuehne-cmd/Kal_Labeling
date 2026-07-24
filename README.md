# Kal Labeling

Dieses Projekt erzeugt die Master-Target-Reiter und optional die Drucker-XLSX-Dateien aus dem Google-Sheet `KAU_PE MUs Summary`.

Enthalten sind:

- `Labeldruck_Master_Target_Colab_BEREINIGT.ipynb`: Colab-Version
- `kal_labeling.py`: terminalfähige Python-Version
- `Vorlage Importdatei.xlsx`: lokale XLSX-Vorlage
- `KAL Print_File_Master - Muster.csv`: Beispiel für die gewünschte Spaltenstruktur

## Zuordnung

Für jede gültige Source-Zeile aus Spalte AA werden drei Labelzeilen erzeugt:

| Labelzeile | Target A | Target C |
| --- | --- | --- |
| Equipment | EQM Nummer (Z) | Messstellen-Beschreibung (G) |
| Functional Location | Übersetzung auf neue FLO (X) | Messstellen-Beschreibung (G) |
| Legacy-Bezeichnung | Asset ID / alt (E) | leer |

Übernommen werden nur AA-Werte nach dem Muster `Zahl - Name`. Leere AA-Werte und Texte ohne führende Zahl werden ausgeschlossen.

## Colab

1. `Labeldruck_Master_Target_Colab_BEREINIGT.ipynb` in Google Colab hochladen.
2. **Runtime → Run all** ausführen.
3. Bei der Abfrage die `Vorlage Importdatei.xlsx` hochladen.
4. Google-Zugriff erlauben.
5. Nach der Master-Erstellung den Reiter `00_Datenvalidierung` prüfen.
6. XLSX-Dateien erst nach der Ja/Nein-Abfrage exportieren.

Der Reiter `00_Datenvalidierung` bricht den Lauf nicht ab. Er listet fehlende oder offensichtliche Fehlerwerte aus den relevanten Source-Spalten E, G, X und Z auf und verlinkt die konkrete Source-Zeile.

## Terminal-Version

### Installation

```bash
cd /Users/kuehnej/dev/Kal_Labeling
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Die Google APIs müssen für das verwendete Konto verfügbar sein. Für lokale User-Credentials kann Application Default Credentials verwendet werden. Alternativ kann ein Service-Account-Schlüssel über `--credentials-file` angegeben werden; dieser Schlüssel darf nicht in Git eingecheckt werden.

### Ausführen

Master-Target erstellen und vor XLSX-Export fragen:

```bash
python kal_labeling.py --template "Vorlage Importdatei.xlsx"
```

XLSX-Export ohne Rückfrage ausführen:

```bash
python kal_labeling.py \
  --template "Vorlage Importdatei.xlsx" \
  --export
```

Mit Service-Account:

```bash
python kal_labeling.py \
  --template "Vorlage Importdatei.xlsx" \
  --credentials-file /sicherer/pfad/service-account.json
```

## Text- und Barcode-Sicherheit

Die Werte werden mit `FORMATTED_VALUE` aus Google Sheets gelesen, damit Formelergebnisse und nicht Formeln übertragen werden. Führende Apostrophe werden entfernt. In den XLSX-Dateien wird jede Zielzelle als Text (`@`) formatiert und `quotePrefix=False` gesetzt.

## GitHub

Repository: https://github.com/jenskuehne-cmd/Kal_Labeling
