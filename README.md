# MewBox — Mewsie's Multi-XML Toolbox

Tired of doing monthly maintenance with the same files, different values, every single time?
	**MewBox** is a swiss-army-knife of a toolbox that generates, adjusts, and audits most of the XML files you deal with on a regular basis for rotational patchwork, so you can spend less time
		copy-pasting and more time actually building content.

Works best when you're already organized and documented! If not, prepare to start! (It's a good habit!)

> *Three months of throwing myself at a screen so you don't have to.*

This tool was made on behalf of the GMs of the private servers before me, whom of which, I watched struggled, time and time again, on behalf of the servers they built and loved and maintainted.

It was inspired by kind people who offered me a helping hand in that same struggle, attempting to do this sisyphusian task of maintenance, over and over again.

This is the least I can do to try to pay it forward. I hope it saves you as much time as it saved me.

May your gardens grow ever-fruitful, ever-flourishing, ever-faithful.

— *Mewsie*

---

## Getting Started

**Requirements:** Python 3.9 or newer · `openpyxl` (installed automatically by the launcher)

**To launch:** double-click `launch.bat`. It will check your Python installation, install any missing dependencies, create the default output folders (`libconfig/`, `reports/`, `MyShop/`), and start the toolbox.

# NOTE - IF YOU ARE USING TO-TOOLBOX/A REGULAR LIBCONFIG (NOT SPLIT XML TABLES):

**YOU MUST EXPORT EACH TABLE FIRST USING TO-TOOLBOX!**

1. https://github.com/TricksterOnline/TO-Toolbox
2. DAT Viewer > Open DAT > libconfig.dat > Export XML (libconfig.xml)
3. LibConfig Editor > Load LibConfig.XML > (libconfig.xml)
4. Use the search to define the table you want to edit (e.g. any ItemParam, PresentItemParam2, R_ShopItem [pick a specific one], etc.)
5. Click the table
6. Export > CSV

This is the foundation this tool works on. The formatted CSV might need altering in order to function with this tool for the most part,
but there are templates already to help you with that that you can download in the tool itself.

**The one exception for this is the "Toolbox Swapper" tool--that's designed specifically for you!**

This will enable you to make your CSV into an XML that can be combined later and back again.

**Use the Toolbox Swapper to change your CSV (YourTableName.CSV) to the XML.**

You'll essentially be using the same formatting that the split libconfig servers use in their XML files from this point on in the program.

**Then, use the XML Bulk Updater to combine the XML from the program to the XML from your CSV.**
**After, use the Toolbox Swapper to create a CSV again (YourTableName.CSV).**

**Go back into To-Toolbox.**
1. LibConfig Editor > Load LibConfig.XML > (libconfig.xml)
2. Import Table > (YourTableName.CSV)
3. You've updated the whole table!

So, yes, even the full-libconfig servers can utilize this tool :)

**Please be aware**: If you are making a new table from scratch in a full libconfig, this will not be helpful!
	I believe this is potentially possible for a S2 server to accomplish, however, it will have to be
	a seperate tool to properly import to your XML. I'll see what I can do for you guys later.

---

## What's Inside

### Generators
| Tool | What it makes |
|---|---|
| **ItemParam Generator** | Rows for `itemparam2.xml`, `ItemParamCM2.xml`, etc. with dropdown selectors, tooltips on every field, compound/exchange support, and a `PresentItemParam2` row builder. Remembers your last-used IDs between sessions. |
| **Box XML Generator** | Full `PresentItemParam2` rows plus optional `RecycleExceptItem`, `Compound_Potion`, `ExchangeShopContents`, and `R_ShopItem` rows. Supports MyShop (`libcmgds_e` + SQL) output. Wide-format and tall-format CSV both accepted. |
| **Set Item Generator** | `CMSetItemParam.xml` rows for fashion/item sets, up to 8 pieces per set. |
| **Compound / Exchange / Shop Generator** | `Compound_Potion.xml` + `Compounder_Spot.xml`, `ExchangeShopContents.xml` + `Exchange_Location.xml`, and `R_ShopItem.xml` — all from a single screen. |

All generators accept **CSV or Excel import** and include **downloadable template variants** so you can see exactly which columns are required vs optional. Templates range from minimal (just the essentials) to full (every supported column with an example row).

### Adjusters & Updaters
| Tool | What it does |
|---|---|
| **Box Rate / Count Adjuster** | Adjusts `DropRate_#` and `ItemCnt_#` slots in `PresentItemParam2.xml`. Can target leaf boxes inside nested parent boxes, rewrite `Type` and `DropCnt`, and warns you if an Egalitarian box has 7+ items (vanilla engine limit). |
| **NCash Updater** | Updates `<Ncash>` values across ItemParam XML files. Two modes: simple CSV of IDs + values, or parent-box mode that recursively walks `PresentItemParam2` to find and update the actual leaf items. |

### XML Utilities
| Tool | What it does |
|---|---|
| **Reorder XML** | Re-sorts rows in any XML file by a chosen field, numerically or alphabetically. |
| **Row Counter / Updater** | Counts rows in XML files and optionally rewrites the `RowCount` attribute in the table header. |
| **Fix ItemParam** | Cleans up common formatting issues in ItemParam XML files (CDATA wrapping, decimal precision, tag structure). |
| **Row Duplicator** | Duplicates rows across one or more XML files, applying math equations to any field (e.g. `y+1` to increment IDs). |
| **Mass Variable Manipulation** | Applies bulk edits — math expressions, text replace, regex replace, conditional filtering — to any field across all matched rows in an XML file. Genuinely powerful and worth reading the built-in warning before using. Also allows CSVs now! |

### Auditors & Reporters
| Tool | What it does |
|---|---|
| **Range Auditor** | Scans an XML file's numeric and text fields, reports used ranges, gaps, and unique values. |
| **XML Comparator** | Compares two XML files side by side to find rows present in one but not the other, with optional CSV lookup. |
| **Data Extract** | Exports selected fields from any XML file to CSV, with optional row filtering and a live row preview. |
| **ID Checker** | Scans one or more XML files for duplicate or conflicting `<ID>` values. |
| **XML Bulk Updater** | Merges one or more raw-ROW patch files (output from any toolbox generator) into a full structured XML table. Rows with matching IDs are replaced; new IDs are appended. RowCount is updated automatically. Output filename is taken directly from the TABLE name attribute. The source file's formatting is preserved exactly — no indentation added, no XML declaration inserted. |

### Other Tools
| Tool | What it does |
|---|---|
| **Toolbox Swapper** | Converts between XML tables and CSV — in both directions. XML → CSV exports all field values from a structured table. CSV → XML rebuilds the table, optionally using a Base XML to preserve the original FIELDINFO headers and structure. Raw-ROW XML files and plain CSVs (without TABLE headers) are supported with a warning prompt before proceeding. |
| **MyShop + DB Generator** | Make MyShop Listings for MyShop easily (Polar Files only) |

---

## CSV / Excel Import

Every generator that accepts a spreadsheet uses **fuzzy column matching** — column names are normalised (case-insensitive, spaces and underscores ignored) before matching, so `ChrTypeFlags`, `chr_type_flags`, and `chrtypeflags` all work. Columns you don't include are simply left at their defaults.

Download the template variants from each generator's start screen to see exactly how flexible each format is.

---

## Settings

Open the **Settings** menu from the bottom of the sidebar to configure:

- **Output directories** — libconfig, reports, and MyShop folders
- **File Naming** — optional timestamp suffix on every output file; option to skip all file picker dialogs and write directly to the configured folders
- **XML Filename Overrides** — rename any output file (sortable by tool or A–Z)
- **CSV Exports Directory** — where Toolbox Swapper writes its CSV output files
- **UI & Performance** — disable tooltips if the interface feels slow on your machine
- **Advanced Manual Renaming** — enables regex/pattern replacement in manual input fields (read the warning first)
- **Variable Editor** — inspect and edit any value persisted from a previous session

---

## Output

By default, every tool shows an output screen with tabs for each generated file and a **Save As…** button on each tab. 
The **Export All** button sends everything to your configured folders at once. 
If you enable *Skip file picker* in Settings, Export All writes directly without asking.

XML files go to the **libconfig** folder. 
Log and report files go to the **reports** folder. 
MyShop files go to the **MyShop** folder. 
CSV exports from Toolbox Swapper go to the **csv_exports** folder.

---

## Notes

- The fashion creation tool was removed from this version — it made the file too large to work with and refactoring it in time wasn't possible. It will be a seperate tool.
- Found a bug? Please report it! I will try to fix it, but I may rewrite this as a whole.
- Session data (last-used IDs, field values) is stored locally in your user profile and persists between launches. Use the Variable Editor in Settings to inspect or clear it.
