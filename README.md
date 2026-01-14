# FutureSecure-Pricing-VBA-Automation

Phase A: Dynamic Spreadsheet Model (Excel VBA)

Overview:
This phase focuses on building a Dynamic Gross Premium Model for the "FutureSecure" product using Excel and VBA. The goal is to ensure the model remains flexible and less prone to manual errors when assumptions change.

Key Features:
1.Automated Pricing: The model calculates Gross Premiums for ages 30, 40, and 50 based on a 10-Year Renewable Term Assurance structure.
2.Switching Capability: Users can switch between Base and Shocked mortality rates on an "Input" tab.
3.Dynamic Assumptions: The premium updates automatically when the Interest Rate (set at 4% p.a.) or mortality basis is modified, without the need to manually change cell formulas.

VBA Implementation:
1.Macro Integration: The workbook uses VBA to handle the "switching" logic and ensure the calculation engine updates correctly across all age scenarios.
2.Data Validation: Ensures that the input parameters meet the Actuarial Standard of Practice (ASOP) before executing the pricing run.

How to Run:
1.Download and open GP ASSIGNMENT VBA LOCKED-1.xlsm.
2.Enable Macros when prompted by Excel.
3.Go to the "Input" tab.
4.Select the mortality basis (Base/Shocked) and input the interest rate.
5.Observe the automatic updates for Gross Premiums at ages 30, 40, and 50.
