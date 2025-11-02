# Excel-VBA-FMCG-Price-Promotion-Analysis
End-to-end Excel VBA data model for evaluating price elasticity, promotion lift, and sales performance. Automates data preparation and integration, calendar mapping, KPI computation, and scenario simulations for FMCG/Retail analysis.

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
- Sales Metrics: sales volume, revenue, margin, and average selling price.
- Baseline and Incremental Metrics: baseline demand, incremental uplift during promotions, incremental revenue, and incremental margin.
- Price Sensitivity Metrics: price elasticity and price–volume relationships.
- Promotion Effectiveness Metrics: promotional uplift, cost of promotion, and ROI.
- Competitive Positioning Metrics: price index and relative price gap to competitors.
- Media and Marketing Influence Metrics: media support levels and their interaction with sales and promotion outcomes.

## Data structure and integration setup

**Note:** All data used in this portfolio project is fictional and does not represent any real companies, products, stores, brands, or market activities.

This step establishes a consistent data foundation that enables reliable analysis and automated reporting. The objective is to organize the raw data into a structured model where each dataset (Sales, Products, Stores, Calendar, Promos, Pricelist, Competitor, Media) is connected through common keys and standardized formats.

The data is integrated using conformed dimensions such as Product, Store, and Date, ensuring that performance metrics can be analyzed consistently across categories, channels, and time periods. During this step, data types are aligned, lookup relationships are validated, and the datasets are prepared for efficient processing in VBA.

To support the analytical model, the workbook is organized as:
- Fact table: Sales transactions and promotional activity records.
- Dimension tables: Products (product attributes) Stores (store attributes), and Calendar (calendar variables).
- Supporting tables: Pricelist (list of prices), Competitor (competitor price data), and Media (media activity inputs).

This structured data model enables the analysis pipeline to run consistently, supports repeatability, and provides a clean foundation for the KPI calculations and dashboard automation that follow.

### DataValidation sheet

Manual checks (COUNTIF/COUNTIFS, SUMPRODUCT) verify keys and relationships before any modeling or automation.

### LookupMap

Join blueprint: It shows which tables join on which keys, and the exact lookup formulas (INDEX/MATCH, COUNTIFS keys, helper keys).

## Data preparation and cleaning

## Modeling and KPI computation

## Dashboard and reporting
