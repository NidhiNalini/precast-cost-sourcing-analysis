# Hands-on Exercises

These steps are done in Microsoft Excel (Microsoft 365 or Excel 2021). They add the parts that
cannot be created by a script: Power Query connections, PivotTables, and XLOOKUP formulas.
Allow about 60 to 90 minutes in total. Commit your work to Git after each exercise.

---

## Exercise 1 – Power Query: load and clean the raw files (about 25 min)

**Goal:** repeat the cleaning done by the workbook formulas, this time with Power Query.

1. Open `workbook/Precast_Cost_Analysis.xlsx`.
2. **Data > Get Data > Launch Power Query Editor**.
3. Create the folder parameter:
   **Home > Manage Parameters > New Parameter** → Name `ProjectFolder`, Type `Text`,
   Current Value = the folder where you cloned this repository (no trailing backslash).
4. For each file `01_Components_Clean.pq`, `02_BOM_Clean.pq`, `03_Sales_Clean.pq`:
   - **Home > New Source > Other Sources > Blank Query**
   - **Home > Advanced Editor**, delete the text, paste the file contents, click **Done**
   - Rename the query to match the file (for example `Components_Clean`)
5. Read each step in the **Applied Steps** pane and click through them. Make sure you can
   explain what each step does.
6. **Home > Close & Load To… > Table > New worksheet** for each query.
7. **Check your result:**
   - `Components_Clean` should have **18 rows** (RB-15M appears only once).
   - `BOM_Clean` should show **status = CHECK** on the same lines as the `BOM` sheet.
   - The sum of `extended_cost_cad` for one item should match `Cost_Rollup` column D.

**Commit:** `Add Power Query cleaning queries`

---

## Exercise 2 – PivotTable and PivotChart (about 20 min)

**Goal:** monthly revenue by product category.

1. Click any cell in the `Sales_Clean` table (from Exercise 1).
   *(If you skipped Exercise 1, use the `Sales` sheet.)*
2. **Insert > PivotTable > New Worksheet**. Rename the sheet `Pivot_Sales`.
3. Set the fields:
   - **Rows:** `invoice_month` (if Excel groups it, keep **Months** only)
   - **Columns:** `category`
   - **Values:** `revenue_cad` (Sum). Format as currency with no decimals.
4. **Insert > Slicer** → `ship_to_country`.
5. **PivotTable Analyze > PivotChart** → Clustered Column or Line.
6. Compare the PivotTable grand total with `Summary!B4`. Note any difference and why
   (hint: unknown items and missing prices).
7. Take a screenshot and save it as `docs/screenshots/pivot_sales.png`.

**Commit:** `Add sales PivotTable and chart`

---

## Exercise 3 – Convert INDEX/MATCH to XLOOKUP (about 15 min)

The workbook uses INDEX/MATCH so it works in every Excel version.
XLOOKUP is shorter and easier to read.

On the `BOM` sheet, replace the formula in **G2**:

```
Before: =IFERROR(INDEX(Components!$B$2:$B$500,MATCH(F2,Components!$I$2:$I$500,0)),"NOT FOUND")
After:  =XLOOKUP(F2,Components!$I$2:$I$500,Components!$B$2:$B$500,"NOT FOUND")
```

The fourth argument of XLOOKUP (`if_not_found`) replaces IFERROR.
Copy the formula down, then do the same for **H** (comp_uom) and **K** (unit_cost_cad).

**Check:** `Data_Quality` counts and `Cost_Rollup` values must not change.

**Commit:** `Use XLOOKUP for BOM lookups`

---

## Exercise 4 – Fix an issue and watch the workbook update (about 10 min)

1. On `Components`, enter **2.10** as the unit cost of `SEAL-BUT`.
2. Observe:
   - `Data_Quality` DQ-02 and DQ-09 drop.
   - On `Cost_Rollup`, which items change from **Incomplete** to **Complete**? Why are
     `PC-RCP-1200` and `PC-MH-1200B` still Incomplete?
3. On `Open_Questions`, set Q02 to **Closed** and write the answer.
4. Add a new row to the `CHANGELOG.md` describing the change.

**Commit:** `Update butyl sealant cost and close Q02`

---

## Exercise 5 – Screenshots for the README (about 10 min)

Save these images in `docs/screenshots/` and add them to the README under **Screenshots**:

| File | What to capture |
|---|---|
| `cost_rollup.png` | The Cost_Rollup sheet |
| `data_quality.png` | The Data_Quality sheet |
| `sourcing.png` | The Sourcing sheet |
| `pivot_sales.png` | Your PivotTable and chart |
| `power_query.png` | The Power Query Editor showing Applied Steps |

Markdown to add an image: `![Cost roll-up](docs/screenshots/cost_rollup.png)`

**Commit:** `Add screenshots to README`
