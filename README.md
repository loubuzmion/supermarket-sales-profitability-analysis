# Supermarket Sales & Profitability Analysis (Excel Analytics Project)
 
**Tools Used:** Microsoft Excel (Pivot Tables, Data Normalization, Lookup & Text Functions, Conditional Filtering)  
**Dataset:** 1,000 Supermarket Transactions across 3 Branches in Myanmar (Jan 1, 2019 – Mar 30, 2019)

---

## 📌 Project Overview

This project analyzes Q1 2019 retail transaction data across three supermarket branches (**Branch A – Yangon**, **Branch B – Mandalay**, and **Branch C – Naypyitaw**) to evaluate revenue drivers, product-line profitability, customer loyalty segments, data quality gaps, and customer satisfaction metrics.

## 📊 Executive Summary & KPIs

| Metric | Value |
| :--- | :--- |
| **Total Transactions** | 1,000 |
| **Total Gross Revenue** *(incl. 5% Tax)* | $322,966.75 |
| **Net Sales** *(Tax Excluded)* | $307,587.38 |
| **Total Gross Profit** | $57,117.28 *(17.69% blended margin)* |
| **Total Units Sold** | 5,510 items *(5.51 avg items/basket)* |
| **Average Transaction Value** | $322.97 *(Median: $253.85)* |
| **Overall Customer Rating** | 6.97 / 10 *(Median: 7.0)* |

## 🔍 Key Findings & Business Insights

#### 1. Product Line Revenue vs. Gross Profit
While top-line revenue is relatively balanced across all six categories ($49.2k–$56.1k), varying gross margins (10% to 28%) create major differences in bottom-line profitability:

| Product Line | Transactions | Units Sold | Total Revenue | Gross Margin | Gross Profit | Profit Share |
| :--- | ---: | ---: | ---: | ---: | ---: | ---: |
| **Health and beauty** | 152 | 854 | $49,193.74 | 28% | $13,774.25 | 24.12% |
| **Food and beverages** | 174 | 952 | $56,144.84 | 23% | $12,913.31 | 22.61% |
| **Home and lifestyle** | 160 | 911 | $53,861.91 | 19% | $10,233.76 | 17.92% |
| **Electronic accessories** | 170 | 971 | $54,337.53 | 15% | $8,150.63 | 14.27% |
| **Sports and travel** | 166 | 920 | $55,122.83 | 12% | $6,614.74 | 11.58% |
| **Fashion accessories** | 178 | 902 | $54,305.90 | 10% | $5,430.59 | 9.51% |

  * Insight: Health and Beauty generates the lowest revenue ($49,193.74) but contributes the highest profit ($13,774.25)—more than 2.5× the profit of Fashion Accessories ($5,430.59) despite selling fewer units.

#### 2. Branch & Regional Performance

| Branch | City | Transactions | Total Revenue | Avg. Basket | Median Rating |
| :---: | :--- | ---: | ---: | ---: | ---: |
| **C** | Naypyitaw | 328 | $110,568.71 | $337.10 | 7.1 |
| **A** | Yangon | 340 | $106,200.37 | $312.35 | 7.1 |
| **B** | Mandalay | 332 | $106,197.67 | $319.87 | 6.7 |

  * Insight: Naypyitaw (Branch C) outperforms other locations in total revenue due to a ~7.9% higher average basket size than Yangon. Mandalay (Branch B) trails in customer satisfaction (6.7 median rating vs. 7.1 elsewhere).
#### 3. Customer Segmentation: Member vs. Normal
  * Overall Spend: Members (501 transactions) generated $164,223.44 ($327.79/txn) compared to $158,743.31 ($318.12/txn) for Normal customers (499 transactions).
  * Category Divergence: Members spend most on Food and Beverages ($31,357.62) and least on Electronic Accessories ($24,498.50). Conversely, Normal customers spend most on Electronic Accessories ($29,839.04).
#### 4. Data Hygiene & Review Engagement
  * Missing Payment Data: Audit of Column M (Payment) identified 31 missing entries (3.1% of records) requiring POS validation rules.
  * Review Engagement: 71.0% of shoppers left text feedback, averaging a 7.02 rating compared to 6.86 among non-reviewers.

## 🛠️ Excel Tasks & Methodology Summary

| Task | Analytical Objective | Excel Techniques Applied | Key Output |
| :---: | :--- | :--- | :--- |
| **Q1** | Morning peak bulk purchases at Branch A (`< 12:01 PM`) | Multi-criteria filtering (`Branch = "A"`, `Time < 12:01 PM`, `MAX Quantity = 10`) | 7 Invoice IDs identified (`232-11-3025`, `851-28-6367`, `447-15-7839`, `325-77-6186`, `443-59-0061`, `811-03-8790`, `243-55-8457`) |
| **Q2** | Exact record lookup preserving original sort order | `XLOOKUP` / `INDEX-MATCH` on Invoice `886-77-9084` | Rating = 5.5 *(Branch C, Electronic accessories)* |
| **Q3** | Regional customer satisfaction benchmark | `MEDIAN` + `IF` array filtering for Mandalay & Yangon | 13.8 *(Mandalay 6.7 + Yangon 7.1)* |
| **Q4** | Date standardization | Custom date formatting & `TEXT` normalization | All 1,000 dates standardized to `YYYY/MM/DD` |
| **Q5** | Customer review length feature engineering | `LEN()` with blank/null handling (`0` for empty reviews) | Character counts added across 1,000 records |
| **Q6** | Text engagement in lowest-revenue transactions | Sorting by Total ascending + `SUM` of review lengths for bottom 10 | 346 characters total across bottom 10 orders |
| **Q7** | Missing data audit in Payment column | `ISBLANK` / `TRIM` null detection & `COUNTIF` | 31 missing/blank entries flagged |
| **Q8** | Loyalty program behavioral comparison | Pivot Table *(Customer type × Product line)* | Members lead total sales ($164.2k vs. $158.7k) |
| **Q9** | Product line revenue ranking | Pivot Table aggregation *(Sum of Total)* | Food & Beverages ranks #1 ($56,144.84) |
| **Q10** | Supermarket gross profit modeling | Margin-weighted profit formulas (`Revenue × Margin %`) | $57,117.28 total gross profit |

## 💡 Strategic Recommendations
1. Promote High-Margin Categories: Bundle Health and Beauty (28% margin) and Food and Beverages (23% margin) in prominent end-cap displays and checkout promotions to expand overall store profitability.
2. Convert Electronic Accessories Buyers into Members: Normal shoppers outspend Members by +$5,340.54 in Electronic Accessories. Offering instant membership discounts on electronics can capture high-intent non-members into the loyalty program.
3. Investigate Customer Experience at Branch B (Mandalay): Conduct targeted review text analysis on Mandalay's lower-rated transactions (median 6.7) to resolve operational bottlenecks (e.g., checkout queues, signage, or store temperature noted in customer comments).
4. Enforce POS Payment Validation: Make the Payment field mandatory at checkout to eliminate the 3.1% missing-value rate.
