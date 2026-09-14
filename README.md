# Automated Workbook Data Transformation

A data preparation and transformation pipeline built with Power Query to automate inventory and quantity standardizations.

## Overview
This project automates the cleaning and transformation of raw workbook data. It utilizes custom Power Query logic to dynamically adjust unit quantities (e.g., standardizing varied pack sizes into consistent base weights) and clean categorical data, ensuring the dataset is fully sanitized and ready for downstream business intelligence modeling.

## Features
* **Automated Data Cleaning:** Removes unnecessary columns, filters nulls, and standardizes data types.
* **Custom Quantity Calculations:** Implements conditional formatting and calculated columns to convert specific item quantities based on item names.
* **Unit Standardization:** Replaces varied unit of measurement (UoM) tags with consistent formatting across the dataset.

## Repository Structure
* `/data`: Contains sample input datasets. *(Note: Ensure sensitive raw data is added to `.gitignore`)*
* `/workbooks`: The main `.xlsx` or `.pbix` files containing the automated Power Query connections.
* `/assets`: Documentation and screenshots of the transformation logic.

## Getting Started
1. Clone this repository to your local machine.
2. Open the primary workbook located in the `/workbooks` directory.
3. Navigate to **Data > Queries & Connections** (Excel) or **Transform Data** (Power BI) to view the applied steps in the Power Query Editor.
4. To apply the automation to new data, update the source file path in the first Applied Step and click **Refresh All**.
