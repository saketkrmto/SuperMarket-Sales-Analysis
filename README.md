# SuperMarket-Sales-Analysis
Supermarket Sales Analysis transforms 500 transaction records into actionable business insights using Python and an Excel VBA dashboard. It identifies top-selling products, high-performing branches, category demand, payment preferences, and customer trends to support smarter inventory planning, branch monitoring, and sales decisions.
# Supermarket Sales Analysis — Data Analytics Project

**Student:** Saket Kumar Mahato  
**Institute:** Bangalore Institute of Technology  
**Program:** IBM SkillsBuild Data Analytics with AI Academic Internship Program

## 1. Project Overview

This project analyzes **500 supermarket transactions** to identify patterns in products, branches, categories, customer types, payment methods, ratings, and sales over time.

## 2. Problem Statement

Raw supermarket transaction records do not immediately show which products, branches, categories, or payment methods have stronger observed performance. The objective is to clean, summarize, visualize, and interpret the data so business decisions can be supported by evidence.

## 3. Objectives

- Load and inspect the dataset.
- Check missing values, duplicate rows, and data consistency.
- Validate `Sales = Quantity × Unit Price`.
- Calculate key performance indicators.
- Analyze sales by product, branch, category, customer type, and payment method.
- Study ratings and daily sales.
- Use groupby and pivot-table analysis.
- Create clear visualizations.
- Convert observations into business insights without unsupported causal claims.

## 4. Dataset

The workbook contains 500 transaction records with:
`Invoice ID`, `Date`, `Branch`, `City`, `Customer Type`, `Gender`, `Product`, `Category`, `Quantity`, `Unit Price`, `Payment`, `Rating`, `Sales`.

**Dataset source:** Internship/project-provided dataset.  
**Dataset link:** [](https://docs.google.com/spreadsheets/d/1QIX__4VObHFMEXnRM2xJyXmB5JAB2peHrJcQ41_U9TE/edit?usp=sharing)

## 5. Key Results

| Metric | Result |
|---|---:|
| Total sales | ₹244,411.08 |
| Average transaction | ₹488.82 |
| Average rating | 3.99/5 |
| Missing values | 0 |
| Duplicate rows | 0 |

### Major observations

- **Highest-selling product:** Cheese — ₹27,906.30
- **Highest-performing branch:** Branch C (Mumbai) — ₹72,469.45
- **Highest-selling category:** Beverages — ₹56,108.24
- **Most-used payment method:** UPI — 127 transactions
- **Average transaction:** Member ₹483.14; Normal ₹497.07

## 6. Methodology

```text
Raw Dataset → Data Loading → Exploration → Quality Checks
→ Sales Validation → KPI Analysis → GroupBy/Pivot Analysis
→ Visualization → Evidence-based Observations → Business Insights
```

## 7. Masterclass 2 Learning Applied

Masterclass 2, **“AI Discovers the Story,”** is directly reflected through:

- Exploring rows, columns, data types, unique categorical values, and sample records.
- Checking descriptive statistics and data quality.
- Creating KPI summaries.
- Grouping data by business dimensions.
- Creating product, category, branch, payment, rating, and time visualizations.
- Using pivot-table analysis.
- Turning observed patterns into concise business insights.
- Avoiding invented causes when the data only shows patterns or associations.

AI-assisted analysis is treated as a reasoning and communication aid; numerical findings are validated against the actual dataset.

## 8. Masterclass 3 Learning — Documented Scope

Masterclass 3, **“AI Predicts What Happens Next,”** introduced:

- Customer-level aggregation
- RFM analysis
- Churn target definition
- Logistic regression
- Accuracy, precision, recall, and confusion matrix
- Churn probability and customer risk levels
- Precision–recall trade-offs
- Target leakage
- Association vs. causation

These are documented as **learning outcomes and future-scope concepts**, not implemented features of this supermarket project.

### Why churn prediction is not implemented

The current dataset does not contain a persistent `Customer_ID` or sufficient customer purchase history. A valid churn model therefore cannot be constructed without assumptions that could create misleading results.

A future e-commerce/customer dataset with historical purchases and a separated observation/outcome period would be suitable for the Masterclass 3 workflow.

## 9. Business Insights

1. Monitor high-sales products and categories for inventory planning.
2. Study the observed sales patterns of Branch C.
3. Monitor the Beverages category closely.
4. Maintain reliable UPI support because it is the most-used payment method in this sample.
5. Track Member vs Normal transaction values before designing membership campaigns.
6. Track ratings alongside sales to monitor customer experience.
7. Collect customer-level history in future projects for segmentation and predictive analytics.

## 10. Technologies Used

- Python
- Pandas
- NumPy
- Matplotlib
- Seaborn
- Jupyter Notebook
- Excel / tabular data

## 11. Project Files

```text
Saket_SupermarketSalesAnalysis/
├── Saket_SupermarketSalesAnalysis.ipynb
├── requirements.txt
├── Saket_SupermarketSalesAnalysis_ProjectReport.docx
├── README.md
└── dataset/
    └── supermarket_sales.xlsx
```

The development workbook used earlier is `DOC-20260924-WA0000.xlsx`. Place it beside the notebook, or update the notebook path if you use another filename.

## 12. Installation

```bash
pip install -r requirements.txt
```

## 13. How to Run

1. Put the dataset workbook beside the notebook.
2. Open `Saket_SupermarketSalesAnalysis.ipynb` in Jupyter Notebook/JupyterLab.
3. Run cells from top to bottom.
4. Review the tables, charts, observations, and business insights.

## 14. Future Scope

- Add `Customer_ID` and historical customer purchases.
- Perform customer segmentation and RFM analysis.
- Define churn using an observation period and later outcome period.
- Train and evaluate a logistic-regression churn model.
- Compare precision and recall according to business costs.
- Add predicted churn probabilities and customer risk levels.
- Explore sales forecasting if sufficient historical data is available.

## 15. Conclusion

The project demonstrates a complete introductory data-analytics workflow and directly applies Masterclass 2 EDA and AI-assisted analytical thinking. Masterclass 3 predictive-analytics topics are recorded as transferable learning and future scope where the current dataset is not sufficient for valid churn modeling.

**Author:** Saket Kumar Mahato
