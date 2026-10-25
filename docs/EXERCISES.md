# Extending the Workbook in Excel

The workbook is complete as committed. These optional exercises extend it with features
that have to be created inside Excel (Microsoft 365 or Excel 2021 on Windows):
Power Query connections and a PivotTable. Allow about 60 minutes in total.

Before you start, save a copy of the workbook, and commit your work to Git after each exercise.

---

## Exercise 1 – Load the Power Query queries (about 25 min)

**Goal:** repeat the cleaning done by the workbook formulas, this time with Power Query.

1. Open `workbook/Precast_Cost_Analysis.xlsx`.
2. **Data > Get Data > Launch Power Query Editor**.
3. Create the folder parameter:
   **Home > Manage Parameters > New Parameter** → Name `ProjectFolder`, Type `Text`,
   Current Value = the folder where you cloned this repository (no trailing backslash).
4. For each file, in this order: `01_Components_Clean.pq`, `02_BOM_Clean.pq`, `03_Sales_Clean.pq`
   - **Home > New Source > Other Sources > Blank Query**
   - **Home > Advanced Editor**, delete the text, paste the file contents, click **Done**
   - Rename the query to match the file name (for example `Components_Clean`)

   The order matters: `BOM_Clean` uses `Components_Clean`.
5. If a message about **privacy levels** appears, click **Continue**, set the sources to
   **Organizational**, and click **Save**.
6. **Home > Close & Load To… > Table > New worksheet** for each query.
7. **Check the result:**
   - `Components_Clean` has **18 rows** (RB-15M appears only once).
   - `BOM_Clean` shows **status = CHECK** on the same lines as the `BOM` sheet.
   - The sum of `extended_cost_cad` for one item matches `Cost_Rollup` column D.

**Privacy note:** the `ProjectFolder` parameter stores your folder path, which may include your
Windows user name. Before committing, you can change it to a placeholder such as
`C:\path\to\precast-cost-sourcing-analysis`.

---

## Exercise 2 – PivotTable and PivotChart (about 20 min)

**Goal:** monthly revenue by product category.

1. Click any cell in the `Sales_Clean` table from Exercise 1 (or use the `Sales` sheet).
2. **Insert > PivotTable > New Worksheet**. Rename the sheet `Pivot_Sales`.
3. Set the fields:
   - **Rows:** `invoice_month` (if Excel adds Years or Quarters, keep **Months** only)
   - **Columns:** `category`
   - **Values:** `revenue_cad` (Sum), formatted as currency with no decimals
4. **PivotTable Analyze > Insert Slicer** → `ship_to_country`.
5. **PivotTable Analyze > PivotChart** → Stacked Column.
6. Compare the grand total with `Summary!B4` (about $5.85 M). Any difference comes from
   invoice lines with missing prices or unknown items.

---

## Exercise 3 – Fix an issue and watch the workbook update (about 10 min)

1. On `Components`, enter **2.10** as the unit cost of `SEAL-BUT`.
2. Observe:
   - On `Data_Quality`, checks DQ-02 and DQ-09 go down.
   - On `Cost_Rollup`, which items change from **Incomplete** to **Complete**? Why are
     `PC-RCP-1200` and `PC-MH-1200B` still Incomplete?
3. On `Open_Questions`, set Q02 to **Closed** and write the answer.

---

## Screenshots

If you complete Exercises 1 and 2, add these images to `docs/screenshots/` and the README:

| File | What to capture |
|---|---|
| `power_query.png` | The Power Query Editor showing the **Applied Steps** pane for `BOM_Clean` |
| `pivot_sales.png` | The PivotTable, slicer and chart |

Markdown for an image: `![PivotTable](docs/screenshots/pivot_sales.png)`

---

## For older Excel versions

The workbook uses `XLOOKUP`. To build a version that uses `INDEX/MATCH` instead:

```bash
python scripts/build_workbook.py --compat
```

Example of the same lookup written both ways (BOM sheet, cell G2):

```
XLOOKUP:      =XLOOKUP(F2,Components!$I$2:$I$500,Components!$B$2:$B$500,"NOT FOUND")
INDEX/MATCH:  =IFERROR(INDEX(Components!$B$2:$B$500,MATCH(F2,Components!$I$2:$I$500,0)),"NOT FOUND")
```
