# Securitisation Model — Waterfall Cash Flow Simulator

> This model simulates the cash flow distribution of a mortgage-backed securitisation across senior, mezzanine, and equity tranches under user-defined default scenarios. Built on 34,143 real US mortgage loans from the FHFA public database, it incorporates soft and hard credit enhancements, a 90-day past due default definition, an equity distribution cap, and a reserve account mechanism. The model allows stress testing of tranche performance under default rates ranging from near-zero to approximately 35% of the pool.

---

## Default Model Parameters

> **Note:** All parameters below are configurable within the model. Values shown are default values.

| Parameter | Default Value |
|---|---|
| Pool size | $10,393,354,106 |
| Number of loans | 34,143 |
| Loan tenure | 360 months |
| Senior tranche | 50% at 3.38% |
| Mezzanine tranche | 45% at 7.79% |
| Equity tranche | 5% at 9.02% (lifetime cap applies) |
| Soft credit enhancement | 1% (WAC 6.65% → investor rate 5.65%) |
| Default definition | 90 DPD (Basel) |
| Data source | FHFA FHLBank Public Use Database |

---

## Overview

The data was collected from the U.S. Federal Housing FHFA: Public Use Database - Federal home Loan Bank system (FHLBank). The raw data contained 21 columns x 34,274 rows and it went through a data preprocessing stage to be used in the securitisation model, and the new dataset had the following dimensions, which are (not including calculated columns) :  13 columns x  34143 rows.

### Motive

Securitization has been at the centre when we encounter financial risk, more so after the Global financial crisis. Whilst learning about the framework of the model, there are not many models available where one can see on a real time basis how values are computed, what all caveats do we actually face, what assumptions do we have to make etc. and how changing one key assumption such as the soft or hard credit enhancement or the tranche value affects the distribution of cash flows. The examples we see are oversimplified such as by either fixing the loan value for each borrower or having a fixed set of loans defaulting in each year. 

Although, those examples are great to understand the underlying concept but, something was missing and the questions such as "How do credit modellers build these models? What all do they think about before entering the model preparation, what all caveats do they face while building the model etc?” always existed. 


---

## How to Run the Model

1. Open the file and navigate to the sheet titled **"Waterfall Structure"**
2. To run without credit enhancements, go to cell **U4** and select the number of defaults to simulate from the drop-down menu.
3. To default loans at a specific month, enter a month number in cell **U5**. To simulate at random months, leave **U5** blank.
4. Click **"Run Computation"**
5. Click **"Show Securitisation Calculation"** to view WAC computation, data source details, or to modify PD values per rating grade.
6. Click **"Clear/Reset"** to clear previous output — note that prior outputs are not saved automatically.

### Adjusting Tranche Sizes and Credit Enhancements

Tranche sizes and credit enhancement layers can be adjusted in **columns J to M**. All cells highlighted in light gold accept user input. Once the desired values are entered, click **"Optimise I/Y for Tranches"** to compute the resulting interest rates for each tranche.

---

## Key Assumptions

**1. Default Definition**

A borrower is considered in default when they have missed 3 consecutive monthly payments (90 days past due), at which point the bank proceeds to sell the collateral to recover the outstanding amount.

**2. Default Timing**

All defaults are assumed to occur at the end of the month in which the 90 DPD threshold is breached, consistent with the Basel 90 DPD framework.

**3. Actual vs Expected Defaults**

The model supports two simulation modes:

- **Actual Defaults:** Defaults have already occurred historically. The model reconstructs the waterfall retrospectively. All loan PDs are set to 100% in this mode.
- **Expected Defaults:** Defaults have not yet occurred. The model simulates future scenarios to assess tranche vulnerability under different credit enhancement and tranche size configurations. 

For example, a 60% senior tranche behaves notably differently from a 50% senior tranche under a 1% soft and 5% hard credit enhancement.

*The model does not compute the present value of tranches in either mode.*

**4. Exposure at Default (EAD)**

EAD is calculated as:

```
EAD = Outstanding Balance + Costs Incurred + Lost Benefits + Interest Payments Not Received (past 3 months)
```

For simplicity, costs incurred are assumed to be zero, and lost benefits are approximated as three months of interest payments not received.

**5. Loss Given Default (LGD)**

LGD comprises the costs and benefits of recovery. In practice, when a borrower defaults, three outcomes are possible: (a) Cure, (b) Close, or (c) Litigation. For simplicity, this model assumes that upon reaching the 90 DPD mark, the bank immediately proceeds to litigation, sells the collateral within one month, and closes the loan. A flat LGD of 45% is applied uniformly across all loans by default, it can be changed if needed and the models asks if you wish to change the LGD value.

** Note : ** A single LGD value would be taken for all loans. 

**6. PD Assignment**

Probability of Default values are sourced from S&P's external credit rating tables and mapped to the credit score grades present in the data. The mapping detail is available on the **"Data for Securitisation"** sheet.

PD values can be modified prior to running the simulation. The PDs used reflect a **Through-the-Cycle (TTC)** estimation approach — they do not respond to current or forward-looking macroeconomic changes and represent average risk across a full credit cycle.

**7. Default Rate Assumption**

The model does not assume a constant yearly default rate. A constant default rate would require tracking which loans defaulted in prior periods across every iteration, resulting in cascading loops that are computationally impractical in Excel/VBA and difficult to audit. Instead, the user specifies the number of defaults to simulate, which are then randomly assigned to loans and months.

**8. Recoveries**

Full recoveries are assumed to occur with a one-month lag following the default booking date. This simplification avoids the need for a stochastic multi-period recovery model. More realistic recovery tracking is noted as a future enhancement below.

**9. Waterfall Computation — Three Stages**

- **Stage 1:** The user selects the number of defaults. Loan-level details (loan ID, default month, term, outstanding balance, and cumulative missed cash flows) are computed and displayed in columns W–AD.
- **Stage 2:** Unique default months are identified and sorted. Cumulative outstanding balance, EAD, and missed payments are aggregated by default month. These are displayed in columns AF–AM and serve as waterfall inputs.
- **Stage 3:** For each month, gross cash flow is computed from the full pool amortisation schedule. Missed payments from loans defaulting at months *i*, *i+1*, and *i+2* are deducted (90 DPD logic), and recoveries from month *i−1* are added. The resulting net cash flow is distributed through the waterfall: senior first, then mezzanine, then equity (subject to the lifetime cap), with any residual flowing to the reserve account. In shortfall scenarios, write-downs proceed from reserves → equity → mezzanine → senior.

**10. Why Amortisation Logic and Not Bond-Style Payment?**

The underlying pool consists of amortising loans where each monthly payment includes both interest and principal. By end of term, the pool is fully amortised and the bank has no large principal balance available for a bullet repayment. Modelling this as a bond would require the bank to reinvest all principal receipts at a specified rate throughout the 360-month term — and to separately track how defaults affect that reinvestment balance — adding substantial complexity with limited incremental insight for the model's core purpose.

---

## Limitations and Future Enhancements

The current model covers the essential mechanics of a securitisation waterfall including soft and hard credit enhancements, tranche size flexibility, an interest rate optimiser, and an equity distribution cap. The following enhancements would further increase realism:

**Variable Interest Rates and Changing WAC**
The model assumes fixed interest rates for all loans throughout the term. In practice, retail loans increasingly carry variable rates reset quarterly or annually, which would alter the WAC at each reset and affect all downstream cash flow calculations.

**Prepayment Modelling**
Borrowers typically retain the right to prepay, particularly during low-rate environments. Prepayments accelerate principal return to the pool, introduce reinvestment risk, and shorten the effective life of tranches. Modelling this properly requires per-loan CPR and SMM simulation alongside an interest rate model, which exceeds VBA's practical capabilities at this pool size.

> *To view a prepayment model incorporating variable interest rates and mid-month prepayments, see the companion project* ***AmortizeXcelerator***.

**Multi-Period Recovery Tracking**
In practice, banks track recoveries over 2–3 years before booking a final LGD. Implementing this would require generating stochastic recovery values for each defaulted loan across *n* future periods. For 12,000 simulated defaults tracked monthly over 3 years, this translates to approximately 400,000 additional rows of data per simulation run — beyond what Excel handles efficiently.

**Borrower-Level or Grade-Level LGD**
LGD is fundamentally a function of collateral quality, not just counterparty rating. The current flat 45% LGD is a portfolio-level approximation. A more accurate implementation would assign LGD at the loan or grade level, calibrated against collateral type and historical recovery data.

**Macroeconomic Default Drivers**
Defaults are not random — they cluster around recessions and economic shocks. Incorporating a systematic factor (e.g., via a single-factor Gaussian copula) would allow the model to generate correlated default scenarios where more loans default in bad economic periods and fewer in good ones given the number of defaults selected. This would produce more realistic tranche stress outcomes, particularly for equity and mezzanine.

---

## Data Source

U.S. Federal Housing Finance Agency (FHFA) — [Public Use Database, Federal Home Loan Bank System](https://www.fhfa.gov/data/pudb)

---

*Built in Microsoft Excel with VBA. Tested on Excel 2019 and Microsoft 365.*
