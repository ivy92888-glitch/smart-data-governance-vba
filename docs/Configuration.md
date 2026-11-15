# Configuration

Every adjustable constant is at the top of `src/SmartDataGovernance.bas`.
**Source code is the single source of truth** — numbers here reference
names only.

## Three Modules

| Module | Entry | What it does | Input sheets | Output |
|---|---|---|---|---|
| **Deduplicate + Merge** | `RunModule1()` | Merge rows by link+identifier, internal sep-aware dedup, delete redundant rows | LM_Data | LM_Data (rows deleted in-place) |
| **Calibrate By Identifier** | `RunModule2()` | Weighted freq+recency scoring picks best identifier, cleans link suffixes, auto-updates siblings | LM_Data + Reference | Reference (col D overwritten) |
| **Clean Text Fields** | `RunModule3()` | CN: prefer detailed; EN: keep most-caps then format; Material: mixed CN/EN + sep normalize | Reference | Reference (cols B/C/J modified) |
| **Run all** | `RunAllGovernance()` | Run Init → Module 1 → Module 2 → Module 3 | All three | All three |

---

## Sheet Name Constants

| Constant | Default | When to change |
|---|---|---|
| `SHEET_LM_DATA`  | `"LM_Data"`  | Your history table has a different name |
| `SHEET_REF_DATA`  | `"Reference"` | Your reference / target table has a different name |

## Column Letter Constants

### LM_Data sheet (Module 1 & 2)

| Constant | Default | What it is |
|---|---|---|
| `COL_LM_NAME_CN`   | `"A"` | CN product name |
| `COL_LM_NAME_EN`   | `"B"` | EN product name |
| `COL_LM_CODE`      | `"C"` | Identifier / code |
| `COL_LM_LINK`      | `"H"` | Entity link |
| `COL_LM_MATERIAL`  | `"I"` | Material string |
| `COL_LM_BATCH_ID`  | `"J"` | Batch / container identifier |

### Reference sheet (Module 2 & 3)

| Constant | Default | What it is |
|---|---|---|
| `COL_REF_BATCH`    | `"A"` | Batch / container identifier (with optional "-backup" suffix) |
| `COL_REF_NAME_CN`  | `"B"` | CN product name |
| `COL_REF_NAME_EN`  | `"C"` | EN product name |
| `COL_REF_CODE`     | `"D"` | Identifier / code (Module 2 writes here) |
| `COL_REF_LINK`     | `"I"` | Entity link (Module 2 reads here, strips suffixes) |

> Module 3 reads Material from col J of Reference — hardcoded in `RunModule3_CleanFields`.
> Adjust that one if your material column is different.

## Weighted Scoring Constants

| Constant | Default | What it is | When to change |
|---|---|---|---|
| `FREQ_WEIGHT`    | 0.6 | Frequency score weight | Want recency to matter more? Shift to 0.4/0.6 |
| `RECENCY_WEIGHT` | 0.4 | Recency score weight (row numbers act as recency proxy) | Must sum to ~1.0 with FREQ_WEIGHT |

## Runtime Populated Arrays

These are populated in `InitGovernance()` before any module runs. Change
them there or in the `RunAllGovernance()` entry if you run modules
independently.

### `linkSuffixes` — Suffixes stripped from links before matching

Default: `Array("-stocklow", "-outofstock", "-default", "-backup", "-new", "-slow", "-clearance", "-recommended")`

### `specialWords` — EN words that stay UPPERCASE after formatting

Default: `Array("LED", "CD", "DVD", "USB", "HDMI", "CPU", "GPU", "RAM", "ROM", "SSD", "HDD", "IP", "TV", "UK", "US", "EU", "DE", "PL", "FBA")`

### `specialMaterials` — EN material abbreviations that survive mixed CN/EN cleaning

Default: `Array("PVC", "PP", "PE", "ABS", "PMMA", "PETG", "TPU", "TPE", "LED", "USB", "HDMI", "FRP", "GFRP")`

---

## How Each Algorithm Works

### Dedup + Merge (Module 1)

1. Row-internal: CN names dedup on "/" separator; EN names dedup+format; materials dedup on "+"
2. Build groups keyed by `link & "|" & identifier`
3. Within each group: find max row number (keep row), merge A/B/I/J columns from all group rows into keep row (dedup-aware separator-aware)
4. Delete redundant rows **bottom-up** (so row numbers stay valid while deleting)

### Calibrate By Identifier (Module 2)

1. Read LM_Data → strip suffixes from each link → build score table
2. For each `link|identifier` pair: frequency count + row number sum
3. Normalize both dimensions → weighted score per identifier
4. Pick highest-scoring identifier per stripped-back link
5. Walk Reference sheet → write best code where stripped-back link matches
6. **Auto-update siblings**: rows whose batch ID = current batch + "backup" suffix **and** whose CN name includes current CN name → also update their code

### Clean Text Fields (Module 3)

| Field | Rule | Priority |
|---|---|---|
| CN name | Split on "/" → skip base names that have a more detailed (contains "-") variant → dedup | Detailed (has "-") > base |
| EN name | Split on "/" → keep variant with most uppercase letters → FormatENName() | More caps > fewer caps |
| Material | Split on "/" → each part: split on "+" → keep CN tokens + EN tokens that match specialMaterials → drop generic EN ("frame", "material") → normalize "加" and space to "+" → strip lone leading/trailing "+" | CN > special EN > generic EN |

## Sibling Matching Rule (Module 2)

```
Anchor row batch ID:    OOCU8799790
Sibling row batch ID:   OOCU8799790backup       ← matches (append / strip "backup")
Anchor CN name:         Heavy-duty nylon net
Sibling CN name:        Heavy-duty nylon net / plastic net   ← "Heavy-duty nylon net" is INCLUDED
→ YES, sibling gets auto-calibrated to new identifier code
```

The suffix string is hardcoded as `"备用"` in `UpdateSiblings()`. Change it
if your system uses a different sibling marker.

## Second Domain Example — Address Pool & Postal-Code Governance

To show generality, here's how you'd point these three modules at a
completely different use case:

| Original scenario | Second scenario |
|---|---|
| LM_Data = history table (entity links + identifiers) | **Shipping address history** (postal codes + city names + door numbers) |
| Reference = product link table (needs calibration) | **Current address pool** (needs dedup + postal-code calibration + format cleanup) |
| Dedup + Merge = same link → merge product records | **Same door number + postal code → merge address records** |
| Calibrate = weighted best identifier per link | **Weighted best postal code per door number** |
| CleanFields = CN/EN name + material cleaning | **Street name format cleanup** (abbreviate "Street" → "St.", drop duplicate apartment numbers) |
| linkSuffixes = `-stocklow`, `-backup` | `-temporary`, `-newbuilding`, `-temp` |
| specialWords = LED/CD/USB... | Common street abbreviations: ST, AVE, RD, BLVD, CT, DR, PL |
| specialMaterials = PVC/PP/ABS... | Special address qualifiers: STE, APT, UNIT, PO BOX |

Same code, same algorithms, different domain.

---

## Files

```
src/SmartDataGovernance.bas   ← single source of truth
examples/sample-data.xlsx       ← fictional demo workbook
README.md                        ← Chinese-primary project doc
docs/Configuration.md            ← this file
LICENSE
.gitignore
.gitattributes
```
