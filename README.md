# Excel-VBA-FMCG-Price-Promotion-Analysis
End-to-end Excel VBA data model for evaluating price elasticity, promotion lift, and sales performance. Automates data preparation and integration, calendar mapping, KPI computation, and scenario simulations for FMCG analysis.

Upload in progress

<img src="https://media1.giphy.com/media/v1.Y2lkPTc5MGI3NjExOWl4cjE4M3dicGx0Y2Q1Z2I1MHFscHpnaTN3bGphc3ZraXZzdjJmdSZlcD12MV9pbnRlcm5hbF9naWZfYnlfaWQmY3Q9Zw/emySgWo0iBKWqni1wR/giphy.gif" width="150" alt="Loading animation">

<div align="center">
<img width="700" height="607" alt="Workflow" src="https://github.com/user-attachments/assets/56e514fd-a49d-45e7-9fd9-79b4561f070d" />
</div>

## Business question definition
The key business questions are defined in order to clarify the commercial objectives of the analysis and ensure all subsequent work remains aligned with decision-making needs and value delivery. The FMCG company requires a deep understanding of how pricing, promotions, and competitive dynamics influence sales and margin performance, therefore pricing decisions, promotional planning, and resource allocation can be made more effectively.

### Key Business Questions:

**1. Price sensitivity:** How do changes in price affect sales volume, revenue, and margin across different products and store segments?<br>
**2. Promotion effectiveness:** To what extent do promotions generate incremental demand versus merely shifting existing sales?<br>
**3. Profitability of promotions:** Which promotional campaigns drive positive margin contribution, and which result in margin erosion?<br>
**4. Promotion ROI:** What is the return on investment (ROI) of promotional activities when comparing incremental sales to promotional costs?<br>
**5. Competitive price positioning:** How do our price levels compare to competitor offerings, and how does this positioning influence customer purchasing behavior?<br>
**6. Media and marketing influence:** What role does media or marketing support play in amplifying or moderating the effects of pricing and promotional activities?

## Analytical approach design

In this step, the analytical techniques and evaluation logic are defined to ensure that the analysis is consistent, repeatable, and aligned with commercial decision-making needs. The approach specifies how sales performance will be decomposed into baseline and incremental components, how price responsiveness will be measured, and how promotional and competitive effects will be quantified.

The key performance indicators are categorized to reflect different analytical perspectives:
- Sales metrics: sales volume, revenue, margin, and average selling price.
- Baseline and incremental Metrics: baseline demand, incremental uplift during promotions, incremental revenue, and incremental margin.
- Price sensitivity metrics: price elasticity and price–volume relationships.
- Promotion effectiveness metrics: promotional uplift, cost of promotion, and ROI.
- Competitive positioning metrics: price index and relative price gap to competitors.
- Media and marketing influence metrics: media support levels and their interaction with sales and promotion outcomes.

## Data structure and integration setup

### Dataset Description

**Note:** All data used in this portfolio project is entirely fictional and does not represent any real companies, products, brands, stores, or market activities. The dataset is a scaled-down demo version created for portfolio presentation purposes.

The analysis uses eight integrated tables simulating a fictional FMCG company's promotional and sales activities across five major Hungarian retailers, covering the full 2024 calendar year. 

The dataset contains the following tables:

`Sales`: YearWeek, WeekStart, StoreID, SKU, Units, NetPrice_LCU, NetRevenue_LCU, PromoFlag, FeatureDisplayFlag, OnInvoiceDiscount_Pct, OffInvoiceRebate_Pct, Returns_Units

`Products`: SKU, Brand, Category, Segment, PackSize_ml, UnitsPerCase, LaunchDate, Status, StdUnitCost_LCU

`Stores`: StoreID, Retailer, Channel, Region, Format

`Calendar`: YearWeek, Date, WeekStart, WeekEnd, Month, Quarter, FiscalPeriod, HolidayFlag, Season, ISOWeek, ISOYear

`Promos`: PromoID,	SKU,	StoreID,	WeekStart, WeekEnd,	Mechanic,	Depth_Pct,	FeatureDisplayFlag, CoopFunding_LCU, Comments,	OverlapCount

`Pricelist`: YearWeek, SKU, StoreID,	ListPrice_LCU,	AvgNetPrice_LCU,	AvgUnitCost_LCU

`Competitor`: YearWeek, CompetitorBrand, SKU_Comp, AvgPrice_LCU, PromoFlag

`Media`: YearWeek, Channel, Spend_LCU, Impressions, GRPs

This step establishes a consistent data foundation that enables reliable analysis and automated reporting. The objective is to organize the raw data into a structured model where each dataset (Sales, Products, Stores, Calendar, Promos, Pricelist, Competitor, Media) is connected through common keys and standardized formats.

### Data Integration Model (STAR-Schema data model)

The data is integrated using conformed dimensions such as Product, Store, and Date, ensuring that performance metrics can be analyzed consistently across categories, channels, and time periods. During this step, data types are aligned, lookup relationships are validated, and the datasets are prepared for efficient processing in VBA.

To support the analytical model, the workbook is organized as:
- Fact table: Sales transactions and promotional activity records.
- Dimension tables: Products (product attributes) Stores (store attributes), and Calendar (calendar variables).
- Supporting tables: Pricelist (list of prices), Competitor (competitor price data), and Media (media activity inputs).

<div align="center">
<img width="450" height="601" alt="starschema" src="https://github.com/user-attachments/assets/08d2b879-32bc-4c7c-ad74-344ac8ed1cc6" />
</div>

### DataValidation sheet

Manual checks (COUNTIF/COUNTIFS, SUMPRODUCT) verify keys and relationships before any modeling or automation.

### LookupMap

Join blueprint: It shows which tables join on which keys, and the exact lookup formulas (INDEX/MATCH, COUNTIFS keys, helper keys).

## Data preparation and cleaning

## Modeling and KPI computation

## Dashboard and reporting
