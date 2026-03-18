# Employee Timesheet Processor / Munkaidő Nyilvántartás Feldolgozó

[English](#english) | [Magyar](#magyar)

---

## English

### Overview

This tool automatically processes employee timesheet files and generates formatted output files for payroll calculations. It filters shifts by date and type, applies special rules for specific employees, and outputs a clean Excel file with employee names and their shift times.

### Features

✓ Automatically selects the chronologically earlier date from the input file  
✓ Filters shifts by allowed shift types  
✓ Applies special employee rules (e.g., Panácz exclusion, Horváth Bence special case)  
✓ Outputs clean Excel file with NAME, IN, OUT columns  
✓ Cross-platform support (Windows and macOS)  
✓ Automatic output file naming based on filter date  

### Quick Start

#### Windows
1. **Install dependencies**: Double-click `install_dependencies_v2.bat`
2. **Process file**: Drag your Excel file onto `process_timesheet_drag_drop.bat`
3. **Done!** Output file will be created as `rekordok_YYYY-MM-DD.xlsx`

#### Mac
1. **Install dependencies**: Double-click `install_dependencies_mac.command`
2. **Process file**: Drag your Excel file onto `process_timesheet_mac.command`
3. **Done!** Output file will be created as `rekordok_YYYY-MM-DD.xlsx`

### Requirements

- Python 3.7 or higher
- pandas
- openpyxl

### Input File Format

The script expects an Excel file with the following structure:
- Each employee name appears once (First name, Last name columns)
- Following rows contain that employee's shifts (Shift Number, Type, Start Date, In, Out columns)
- Multiple dates can be present in the file (maximum 2 unique dates)

**Example structure:**
```
First name | Last name | Shift Number | Type    | Start Date | In    | Out   |
-----------|-----------|--------------|---------|------------|-------|-------|
Anna       | Kiss      |              |         |            |       |       |
           |           | 1            | Hosszú  | 2026-03-13 | 09:00 | 17:00 |
Bence      | Nagy      |              |         |            |       |       |
           |           | 1            | Leo     | 2026-03-13 | 14:00 | 22:00 |
```

### Output File Format

Simple 3-column Excel file:

| Name       | In    | Out   |
|------------|-------|-------|
| KISS ANNA  | 09:00 | 17:00 |
| NAGY BENCE | 14:00 | 22:00 |

### Processing Logic

1. **Date Selection**: Finds all unique dates in the file and selects the chronologically earlier one
2. **Employee Processing**: For each employee, finds their shift(s) on the selected date
3. **Shift Selection**: If multiple shifts exist for the same employee on the same date, selects the chronologically earlier one
4. **Type Filtering**: Only includes shifts with allowed types
5. **Name Filtering**: Applies special rules for specific employees
6. **Output Generation**: Creates Excel file with employee names and shift times

### Allowed Shift Types

- Hosszú
- Leo
- Winston
- Mogumba
- Konyha
- Nappalos
- Poharas
- Kávézó pult
- Rács
- Első kert

### Special Rules

**Excluded:**
- All employees with last name "Panácz"

**Name Changes:**
- Bence Horváth with shift type "Konyha" → Output as "HORVÁTH BENCE POHARAS"

**Special Type Rules:**
- Shift type "Kávézó" is ONLY allowed for István Prihoda

### Error Handling

**"Hibás bemeneti fájl, több mint 1 napot tartalmaz."**
- The file contains more than 2 unique dates
- Solution: Use a file with only 1-2 dates

**"No valid dates found"**
- The Start Date column is empty or invalid
- Solution: Ensure the file has valid dates in the Start Date column

**"No valid records found"**
- No shifts match the filtering criteria
- Solution: Check that the file contains shifts with allowed types on the selected date

### File Structure

```
📁 Timesheet_Processor/
  ├── process_timesheets.py                    ← Core Python script
  │
  ├── Windows Files:
  │   ├── process_timesheet_drag_drop.bat      ← Drag & drop processor
  │   └── install_dependencies_v2.bat          ← Dependency installer
  │
  └── Mac Files:
      ├── process_timesheet_mac.command        ← Drag & drop processor
      └── install_dependencies_mac.command     ← Dependency installer
```

### Command Line Usage

```bash
python process_timesheets.py <input_file.xlsx> [output_file.xlsx]
```

**Examples:**
```bash
# Basic usage (auto-generates output filename)
python process_timesheets.py timesheet.xlsx

# Specify output filename
python process_timesheets.py timesheet.xlsx output.xlsx
```

### Troubleshooting

**Windows:**
- If `pip is not recognized`: Use `python -m pip install pandas openpyxl`
- If batch file doesn't work: Run `install_dependencies_v2.bat` first
- If Python not found: Install from https://www.python.org/downloads/

**Mac:**
- If permission denied: Run `chmod +x process_timesheet_mac.command`
- If Python not found: Install from https://www.python.org/downloads/
- If dependencies fail: Run `pip3 install pandas openpyxl`

---

## Magyar

### Áttekintés

Ez az eszköz automatikusan feldolgozza a munkaidő-nyilvántartási fájlokat és formázott kimeneti fájlokat generál a bérszámfejtéshez. Dátum és műszaktípus alapján szűri a műszakokat, speciális szabályokat alkalmaz bizonyos dolgozókra, és tiszta Excel fájlt állít elő a munkavállalók neveivel és műszakidőivel.

### Funkciók

✓ Automatikusan kiválasztja a kronológiailag korábbi dátumot a bemeneti fájlból  
✓ Szűri a műszakokat az engedélyezett műszaktípusok alapján  
✓ Speciális szabályokat alkalmaz bizonyos dolgozókra (pl. Panácz kizárás, Horváth Bence speciális eset)  
✓ Tiszta Excel fájlt állít elő NÉV, KEZDÉS, BEFEJEZÉS oszlopokkal  
✓ Többplatformos támogatás (Windows és macOS)  
✓ Automatikus kimeneti fájl elnevezés a szűrési dátum alapján  

### Gyors Kezdés

#### Windows
1. **Függőségek telepítése**: Dupla kattintás az `install_dependencies_v2.bat` fájlra
2. **Fájl feldolgozása**: Húzd rá az Excel fájlodat a `process_timesheet_drag_drop.bat` fájlra
3. **Kész!** A kimeneti fájl `rekordok_ÉÉÉÉ-HH-NN.xlsx` néven jön létre

#### Mac
1. **Függőségek telepítése**: Dupla kattintás az `install_dependencies_mac.command` fájlra
2. **Fájl feldolgozása**: Húzd rá az Excel fájlodat a `process_timesheet_mac.command` fájlra
3. **Kész!** A kimeneti fájl `rekordok_ÉÉÉÉ-HH-NN.xlsx` néven jön létre

### Követelmények

- Python 3.7 vagy újabb
- pandas
- openpyxl

### Bemeneti Fájl Formátum

A szkript egy olyan Excel fájlt vár, amelynek szerkezete a következő:
- Minden dolgozó neve egyszer szerepel (Keresztnév, Vezetéknév oszlopok)
- A következő sorok az adott dolgozó műszakjait tartalmazzák (Műszak szám, Típus, Kezdés dátuma, Kezdés, Befejezés oszlopok)
- Több dátum is lehet a fájlban (maximum 2 egyedi dátum)

**Példa szerkezet:**
```
Keresztnév | Vezetéknév | Műszak szám | Típus   | Kezdés dátuma | Kezdés | Befejezés |
-----------|------------|-------------|---------|---------------|--------|-----------|
Anna       | Kiss       |             |         |               |        |           |
           |            | 1           | Hosszú  | 2026-03-13    | 09:00  | 17:00     |
Bence      | Nagy       |             |         |               |        |           |
           |            | 1           | Leo     | 2026-03-13    | 14:00  | 22:00     |
```

### Kimeneti Fájl Formátum

Egyszerű 3 oszlopos Excel fájl:

| Név        | Kezdés | Befejezés |
|------------|--------|-----------|
| KISS ANNA  | 09:00  | 17:00     |
| NAGY BENCE | 14:00  | 22:00     |

### Feldolgozási Logika

1. **Dátum kiválasztás**: Megkeresi az összes egyedi dátumot a fájlban és kiválasztja a kronológiailag korábbit
2. **Dolgozó feldolgozás**: Minden dolgozó esetében megkeresi a kiválasztott dátumon lévő műszak(oka)t
3. **Műszak kiválasztás**: Ha több műszak létezik ugyanazon dolgozónak ugyanazon a napon, a kronológiailag korábbit választja
4. **Típus szűrés**: Csak az engedélyezett típusú műszakokat veszi figyelembe
5. **Név szűrés**: Speciális szabályokat alkalmaz bizonyos dolgozókra
6. **Kimenet generálás**: Excel fájlt hoz létre a dolgozók neveivel és műszakidőivel

### Engedélyezett Műszaktípusok

- Hosszú
- Leo
- Winston
- Mogumba
- Konyha
- Nappalos
- Poharas
- Kávézó pult
- Rács
- Első kert

### Speciális Szabályok

**Kizárva:**
- Minden "Panácz" vezetéknevű dolgozó

**Név változtatások:**
- Horváth Bence "Konyha" típusú műszakkal → Kimenet: "HORVÁTH BENCE POHARAS"

**Speciális típus szabályok:**
- A "Kávézó" műszaktípus CSAK Prihoda Istvánnak engedélyezett

### Hibakezelés

**"Hibás bemeneti fájl, több mint 1 napot tartalmaz."**
- A fájl több mint 2 egyedi dátumot tartalmaz
- Megoldás: Használj olyan fájlt, amely csak 1-2 dátumot tartalmaz

**"No valid dates found"**
- A Kezdés dátuma oszlop üres vagy érvénytelen
- Megoldás: Győződj meg róla, hogy a fájl érvényes dátumokat tartalmaz a Kezdés dátuma oszlopban

**"No valid records found"**
- Egyetlen műszak sem felel meg a szűrési feltételeknek
- Megoldás: Ellenőrizd, hogy a fájl tartalmaz-e engedélyezett típusú műszakokat a kiválasztott dátumon

### Fájl Szerkezet

```
📁 Munkaidő_Feldolgozó/
  ├── process_timesheets.py                    ← Python szkript
  │
  ├── Windows fájlok:
  │   ├── process_timesheet_drag_drop.bat      ← Drag & drop feldolgozó
  │   └── install_dependencies_v2.bat          ← Függőség telepítő
  │
  └── Mac fájlok:
      ├── process_timesheet_mac.command        ← Drag & drop feldolgozó
      └── install_dependencies_mac.command     ← Függőség telepítő
```

### Parancssori Használat

```bash
python process_timesheets.py <bemeneti_fájl.xlsx> [kimeneti_fájl.xlsx]
```

**Példák:**
```bash
# Alapvető használat (automatikusan generálja a kimeneti fájlnevet)
python process_timesheets.py munkaidő.xlsx

# Kimeneti fájlnév megadása
python process_timesheets.py munkaidő.xlsx kimenet.xlsx
```

### Hibaelhárítás

**Windows:**
- Ha `pip is not recognized`: Használd a `python -m pip install pandas openpyxl` parancsot
- Ha a batch fájl nem működik: Futtasd először az `install_dependencies_v2.bat` fájlt
- Ha a Python nem található: Telepítsd innen: https://www.python.org/downloads/

**Mac:**
- Ha permission denied: Futtasd a `chmod +x process_timesheet_mac.command` parancsot
- Ha a Python nem található: Telepítsd innen: https://www.python.org/downloads/
- Ha a függőségek telepítése sikertelen: Futtasd a `pip3 install pandas openpyxl` parancsot

---

## Version / Verzió

**Version 2.0** - Optimized for new shift report format  
**2.0 verzió** - Optimalizálva az új műszak riport formátumra

## License / Licenc

Proprietary / Zárt forráskódú

## Support / Támogatás

For issues or questions, please refer to the troubleshooting sections above.  
Problémák vagy kérdések esetén kérjük, nézd meg a fenti hibaelhárítási részeket.
