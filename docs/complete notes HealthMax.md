# HealthMax — Net Revenue Management in Excel

**Objective:** Analyze the shampoo market and generate actionable insights to grow the business.  

**Context:** As the Category Manager at HealthMax, the market leader in the shampoo industry, you are responsible for identifying opportunities to increase revenue, optimize pricing and promotions, and strengthen market share. The analysis focuses on leveraging Excel to structure, clean, and interpret market data in order to provide actionable recommendations for strategic decision-making.  

**Dataset Source:** Fictitious dataset provided by DataCamp for educational purposes. Both the company and the dataset are fictional.  

## Data Preparation and Exploration

- Changed the column **`Values Month`** data type to *Currency* and removed decimals, since the values represent whole numbers.  

### Pivot Table — Brands per Supplier

- Created a new worksheet named **`Brands per Supplier`**.  
- Built a Pivot Table with:  
  - **Rows:** Suppliers and Brands (nested)  
- Purpose: To group brands under their respective suppliers and clearly identify which brands belong to each supplier.  


### Pivot Table — HealthMax Growth

- Created a new worksheet named **`HealthMax Growth`**.  
- Built a Pivot Table with:  
  - **Rows:** Brands  
  - **Values:** Sum of *Values Month*  
  - **Columns:** Year  
  - **Filter:** Supplier (set to *HealthMax*)  
- Adjusted the value display to show **Year-over-Year Growth %** compared to the previous year.  
- Since only three months of 2023 data are available, the year 2023 was hidden from the Pivot Table to avoid misleading comparisons.  

### Calculated Columns — Year-to-Date (YTD)

- Created a new column **`Units YTD`** using the following formula:  
  ```excel
  =SUMIFS(H:H, D:D, [@Brand], E:E, [@Region], F:F, [@Year], G:G, "<=" & G1)
  ```
Purpose: Calculate the Year-to-Date cumulative sum of units based on 'Units Month', grouped by Brand, Region, and Year.

Created a new column Values YTD using a similar formula:
````excel
=SUMIFS(I:I, D:D, [@Brand], E:E, [@Region], F:F, [@Year], G:G, "<=" & [@Month])
````
Difference: The sum range is 'Valiues Month' instead of the Units column, in order to compute YTD for values.

Goal: Build a running total by Brand, Region, and Year, accumulating monthly results.


### Calculated Column — Units MAT (Moving Annual Total)

- Created a new column **`Units MAT`** (Moving Annual Total).  
- The formula reuses the logic from **`Units YTD`** and adds the remaining portion from the previous year:  
  ```excel
  =SUMIFS(H:H, D:D, [@Brand], E:E, [@Region], F:F, [@Year], G:G, "<=" & [@Month]) 
   + SUMIFS(H:H, D:D, [@Brand], E:E, [@Region], F:F, [@Year]-1, G:G, ">" & [@Month])
  ```

Purpose: Calculate a rolling 12-month total of units by summing:

YTD units for the current year up to the selected month.

Units from the previous year after the selected month.

### Calculated Column — Values MAT (Moving Annual Total)

- Created a new column **`Values MAT`** following the same logic as **`Units MAT`**.  
- Formula:  
  ```excel
  =SUMIFS(I:I, D:D, [@Brand], E:E, [@Region], F:F, [@Year], G:G, "<=" & [@Month]) 
   + SUMIFS(I:I, D:D, [@Brand], E:E, [@Region], F:F, [@Year]-1, G:G, ">" & [@Month])
   ```
Purpose: Calculate the rolling 12-month total of values (sales), combining:

YTD values from the current year up to the selected month.

Values from the previous year after the selected month.

### Pivot Table — MAT Total Value Category

- Created a new worksheet named **`MAT Total Value Category`**.  
- Built a Pivot Table with:  
  - **Filters:** Year and Month (set to March 2023).  
  - **Values:** *Values MAT*.  
- **Purpose:** Display the Moving Annual Total (MAT) for March 2023 (April 2022 - March 2023).

**Result:**  
- Total turnover of the last 12 months: **$140,958,153**  

### Pivot Table — Market Share

- Created a new worksheet named **`Market Share`**.  
- Built a Pivot Table with:  
  - **Rows:** Year and Month  
  - **Columns:** Brand  
  - **Values:** Sum of *Values Month*  
- **Purpose:** Compare monthly brand performance over time as a foundation for market share analysis.  

### Market Share Analysis

- Updated the Pivot Table values to display **percentage of market share per year**.  
- Sorted brands by size, i.e., by their market share contribution.  
- Created a Pivot Chart (Line Chart) to visualize market share trends across brands.  
- Added a **Slicer** to filter results by Region.  

**Insight:**  
- *Starbust*, a HealthMax brand, achieved the **highest market share in 2023 within the South region**, outperforming its share in other regions.  

## New Worksheet — Internal Sales Data

- Added a new worksheet named **`Internal Sales Data`**.  
- Loaded a CSV file with the same name into this worksheet.  
- **Columns included:**  
  - Brand  
  - Product  
  - Pack Size (ml)  
  - ProductID  
  - Retail Price  
  - Net Price  
  - COGS  
  - Volume 2022  

  - **Purpose:** Provide input data for **Net Revenue Management** analysis.  

  ### Data Cleaning — Volume 2022

- Issue: The column **`Volume 2022`** contained values with periods as thousand separators (e.g., `1.156.348`).  
- Although formatted as *Accounting*, Excel interpreted them as text, preventing calculations such as sum or average.  
- **Solution:** Applied *Find & Replace* to remove all periods (`.`), converting the values into proper numeric format.  
- Verified that Excel now recognizes the column as numeric (status bar shows *Sum* and *Average* instead of only *Count*).  

### Calculated Columns — Profitability Metrics

- Created a new column **`Gross Profit per Unit`**:  
  - Formula: `Net Sales – COGS`  

- Created a new column **`Gross Profit per Product`**:  
  - Formula: `Gross Profit per Unit * Volume 2022`  

- Created a new column **`Gross Margin`**:  
  - Formula: `Gross Profit per Product / Net Sales 2022`  

#### Aggregated Metric — Weighted Average Gross Margin

- Calculated the **Total Weighted Average Gross Margin** by dividing:  
  - **Total Gross Profit per Product** ÷ **Total Net Sales 2022**  

  ### Calculated Column — Net Sales Contribution

- Created a new column **`Net Sales Contribution`**:  
  - Formula: `Net Sales 2022 / Total Net Sales 2022`  
  - Displayed as a percentage format.  

## New Worksheet — Profitability Margin

- Added a worksheet named **`Profitability Margin`**.  
- Built a Pivot Table with:  
  - **Rows:** ProductID  
  - **Columns:** Gross Margin (formatted with no decimals)  
  - **Values:** Net Sales Contribution  

### Scatter Plot — Profitability vs. Contribution

- Copied the Pivot Table results and pasted them below the table, since Excel does not allow direct scatter plots from Pivot Tables.  
- Created a scatter plot using the copied dataset.  
- Placed the chart over the pasted data to avoid redundancy.  

**Result:**  
- A scatter plot visualizing products by **profitability (Gross Margin)** and their **contribution to Net Sales**.  
- Enables quick identification of the most profitable products and their relative sales weight.  

## New Worksheet — New Category Opportunity

- Added a worksheet named **`New Category Opportunity`**.  
- Built a Pivot Table with:  
  - **Rows:** Subcategories  
  - **Columns:** Years (excluded 2023 due to incomplete data)  
  - **Values:** Units Month (Sum)  
- Configured the values to display as **% Difference From**, using **2018** as the base year.  
- Purpose: Compare the growth of each subcategory against 2018, with a focus on analyzing 2022 vs. 2018.  

**Insight:**  
- The **Organic** subcategory demonstrated the highest growth between 2018 and 2022, highlighting a market opportunity to capture a greater share in this segment.  

## New Worksheet — Organic Shampoo Launch

- Imported data from a CSV file named **`Organic_shampoo_launch`**.  
- HealthMax plans to launch **two new products in 2024** within the **Organic** subcategory:  
  - *HerbEssentials*  
  - *Herbashine*  
- The Organic subcategory is expected to reach a **total volume of 9,819,037 units**.  
- According to HealthMax’s internal estimates, **Herbashine** is projected to generate the highest profit among the two products.  

## New Worksheet — 50ml Shampoo 2024

- Created a copy of the **`Internal Sales Data`** worksheet and named it **`50ml Shampoo 2024`**.  
- Added a new column **`Price per ml`**, calculated as:  
  - `Retail Price / Pack Size (ml)`  

### Simulation — Starbust Ultra Soft (50ml Format)

- Selected the best-selling product (*Starbust Ultra Soft 100ml*) and simulated a new 50ml format.  
- Filled in the following columns:  
  - **Brand, Product, Pack Size (ml), ProductID, Retail Price**  
- Calculated the **Retail Price** for the 50ml format by:  
  - `50 * (Price per ml of the 100ml format)`  
  - Applied a **50% mark-up** over the 100ml *Price per ml*.  
- **Net Price:** $2.30  
- **COGS:** $0.70  
- Assumed **Volume = 10%** of the volume of the 100ml format.  

### Results

- **Gross Profit:** $185,015.68  
- **Gross Margin:** 70%  

## Promotion Analysis — Shinez (2022)

### Pivot Table and Chart
- Created a new worksheet named **`Promotion Graph`**.  
- Built a Pivot Table based on **`external Data`** with:  
  - **Filters:** Brand = *Shinez*, Year = 2022  
  - **Values:** Monthly sales  
- Inserted a bar Pivot Chart to visualize sales performance throughout 2022.  

### Imported Data
- Imported a CSV file named **`promotion_analysis`** into a new worksheet with the same name.  

### Calculations
- **Value Sales:** Sales in promotion months, calculated using VLOOKUP:  
  ```
  =VLOOKUP([@Month], 'Promotion Graph'!$A$5:$B$16, 2, 0)
  ```

**Baseline Sales**: Average sales in non-promotion months, calculated with AVERAGEIFS:
=AVERAGEIFS('Promotion Graph'!$B$5:$B$16, 'Promotion Graph'!$A$5:$A$16, "<>4", 'Promotion Graph'!$A$5:$A$16, "<>8", 'Promotion Graph'!$A$5:$A$16, "<>11")

**Uplift**:
=[@Value Sales] - [@Baseline Sales]

**ROI**:
=([@Uplift] - [@Costs]) / [@Costs]

#### **Insight**
The “Buy 2, Get 1 Free” promotion delivered the highest ROI among the three tested campaigns.


## Forecasting — Market and Net Sales (2023–2024)

### Data Preparation
- Created a new worksheet named **`Sales Until 2022`**.  
- Inserted a Pivot Table based on **`External Data`** with:  
  - **Rows:** Years (excluding 2023 due to incomplete data)  
  - **Values:** Sum of monthly sales  
  - **Filter:** Supplier = *HealthMax*  

### Forecasting
- Used the **Forecast Sheet** option to generate a new worksheet named **`Forecast 2024`**, based on the *Sales Until 2022* data.  
- Added two new columns:  
  - **`Net Sales`** → Extracted from *Net Sales 2022* in **`Internal Sales Data`**, applied only for 2022.  
  - **`Ratio`** → Calculated as `Net Sales / Market Value` for 2022.  

### Projections
- Applied the 2022 ratio to forecast **Net Sales** for 2023 and 2024.  

### Insights
- **2023:** Net Sales projected to decline by **0.93%** compared to 2022.  
- **2024:** Net Sales projected to increase by **3.99%** vs. 2022 and by **4.975%** vs. 2023.  


## Net Revenue Management — Waterfall Analysis

### Worksheet Setup
- Created a new worksheet named **`Waterfall`** with the following columns:  
  - **Estimated Net Sales 2023** → Value taken from the **`Forecast 2024`** sheet (projected Net Sales for 2023).  
  - **Natural Growth** → Calculated as `Net Sales 2024 – Net Sales 2023` (baseline growth without initiatives).  
  - **Organic Shampoo** → Added expected Net Sales for 2024 from *Herbashine*, identified as the most profitable Organic product.  
  - **50ml Shampoo** → Added expected Net Sales contribution from the 50ml product launch.  
  - **Estimated Net Sales 2024** → Sum of all initiatives above, assuming no cannibalization across products.  

### Visualization
- Created a **Waterfall Chart** titled *Estimated Net Sales 2024* to illustrate the incremental contributions of each initiative.  

### Insight
- **Estimated Net Sales 2024:** **$24,805,694.99**  
- This result is achieved through **Net Revenue Management initiatives** (Herbashine + 50ml Shampoo),  
  surpassing the growth that would have been achieved with natural growth alone.  
- Without Net Revenue Management, growth would have been **4.98%**.  
- With NRM initiatives applied, growth reached **16.46%**, equivalent to **$3,506,242.32**.  