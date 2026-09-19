# 📗 Financial Performance Dashboard — Excel Project

A standalone, pure-Excel data analysis project: no Python, no SQL, no external
tools required to use it — just Excel (or LibreOffice Calc). Built to show core
analyst Excel skills: formulas, lookups, data validation, conditional formatting,
and an interactive dropdown-driven dashboard.

Dataset: 25 companies across 5 sectors and 5 regions in the Arab world, quarterly
financials from 2020–2024 (500 rows). Data is synthetically generated for
portfolio/learning purposes.

---

## 📸 Screenshots

**Dashboard** — dropdown filters, live KPIs, and a filtered revenue chart:

![Dashboard](screenshots/dashboard.png)

**Data sheet** — Excel Table with conditional formatting:

![Data sheet](screenshots/data_sheet.png)

---

## 📂 What's in the workbook

`Financial_Dashboard_Excel_Project.xlsx` has 5 sheets:

| Sheet | What it shows |
|---|---|
| **Read Me** | Overview + how to convert the summary tables into native PivotTables |
| **Data** | The cleaned dataset as a native Excel Table, with conditional formatting (data bars on profit margin, color scale on net profit) |
| **Lookup** | Pick any company from a dropdown → `INDEX/MATCH` and `SUMIF`/`AVERAGEIF` pull its Sector, Region, total revenue, profit, and average margin live |
| **Pivot_Summary** | Sector × Year and Sector × Region revenue cross-tabs, built entirely with `SUMIFS` |
| **Dashboard** | **Interactive**: pick a Sector and Year from the dropdowns and every KPI card, the revenue chart, and the Top-5-companies table update instantly |

---

## ⚙️ Excel skills demonstrated

- **Excel Tables** (structured references, auto-expanding ranges)
- **Lookup formulas**: `INDEX/MATCH` (with a note on the 1-line `XLOOKUP`
  equivalent for Excel 365/2021)
- **Aggregation formulas**: `SUMIFS`, `AVERAGEIFS`, `SUMIF`, `AVERAGEIF`
- **Ranking**: `LARGE` + `INDEX/MATCH` to build a live Top-5 table
- **Data validation**: dropdown lists driving the Lookup and Dashboard sheets
- **Conditional formatting**: data bars and 3-color scales
- **Nested `IF` logic** to make `SUMIFS`/`AVERAGEIFS` respect an "All" option
  in a dropdown (a common real-world dashboard pattern)
- **Native charts** wired to formula-driven ranges, so they update with the filters
- Formula-driven throughout — every KPI is calculated from the Data table, not typed in

---

## ▶️ How to use it

1. Open `Financial_Dashboard_Excel_Project.xlsx` in Excel or LibreOffice Calc.
2. Go to the **Dashboard** sheet.
3. Use the **Sector** and **Year** dropdowns (yellow cells) to filter — everything
   below updates automatically.
4. Go to the **Lookup** sheet and pick a company from the dropdown to see its
   full profile.
5. (Optional) On the **Data** sheet, select the table and go to
   **Insert → PivotTable** to build your own native pivot view alongside the
   formula-based `Pivot_Summary` sheet.

---

## 💡 Key figures (unfiltered)

- **Total Revenue (2020–2024):** $2,905,209,411
- **Total Net Profit:** $843,902,957
- **Average Profit Margin:** 29.2%
- **Top company by net profit:** FinEdge FI (Finance sector) — ~$50.8M

---

## 🛠️ Tool

`Microsoft Excel` (built and validated with `openpyxl` + LibreOffice Calc for
formula-correctness checking — zero formula errors across 110 formulas)
