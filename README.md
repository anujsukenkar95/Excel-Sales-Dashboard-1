
# 📊 Sales Dashboard

## 📝 Overview

This project is an interactive, macro-enabled Excel Sales Dashboard (`Sales Dashboard 1st.xlsm`). It is designed to track, visualize, and analyze key sales metrics, providing actionable insights through dynamic filtering and automated features.

## ✨ Key Features

Based on the dashboard's architecture, it includes the following capabilities:

* **Interactive Data Aggregation:** Utilizes four distinct Pivot Tables to summarize complex data efficiently.

* **Dynamic Visualizations:** Features at least three distinct charts to visualize sales trends and metrics visually.

* **User-Friendly Filtering:** Incorporates Slicers, allowing users to quickly filter the dashboard data with a single click.

* **Automated Workflows (VBA):** Built as an `.xlsm` file containing a VBA project (`vbaProject.bin`), enabling custom automated actions and macros.

* **Organized Layout:** Structured across multiple worksheets to separate raw data from the main visual dashboard.


## 💻 Prerequisites

* **Microsoft Excel:** Version 2013 or newer is recommended.
* **Macros Enabled:** Because this is a `.xlsm` file, you must click **"Enable Content"** upon opening the file to ensure the VBA macros function correctly.

## 🚀 How to Use

1. **Open the File:** Launch `Sales Dashboard 1st.xlsm` in Microsoft Excel.
2. **Enable Macros:** If prompted by Excel's security warning, click "Enable Content" to activate the automated features.
3. **Interact with Slicers:** Use the clickable slicer buttons to filter the charts and PivotTables by specific categories (e.g., date ranges, regions, or product lines).
4. **Update Data:**
    * Navigate to the raw data sheet.
    * Paste your updated sales data.
    * Go to `Data` > `Refresh All` (or use the designated macro button) to update the Pivot Tables and Charts automatically.



## 🛠️ Customization

To modify the backend calculations, unhide the calculation/PivotTable sheets. You can adjust the PivotTable fields or edit the VBA code by pressing `ALT + F11` to open the Visual Basic Editor.

![First Dashboard Snapshot](dashboard-snap-1.png)

![Second Dashboard Snapshot](dashboard-snap-2.png)
