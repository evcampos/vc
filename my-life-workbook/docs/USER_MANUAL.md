# [MY LIFE] Workbook - User Manual

## Introduction

Welcome to the **[MY LIFE]** financial management system. This comprehensive Excel workbook helps you manage personal and family office finances with automated imports, transaction classification, investment correlation, and executive dashboards.

---

## Core Concepts

### Worksheets Overview

| Worksheet | Purpose | User Interaction |
|-----------|---------|------------------|
| **FILES PATHS** | Configuration of data source locations | Configure once, update when paths change |
| **FILES STRUCTURE** | Column mapping definitions | Reference only |
| **BANKS** | All checking account transactions | Auto-populated, review and classify |
| **CARDS** | All credit card transactions | Auto-populated, review and classify |
| **INVESTMENTS** | Investment movements linked to banks | Auto-populated, check correlations |
| **OPUS** | External investments (outside banking system) | Manual entry, auto-updated costs |
| **DEBTS** | Personal loans and liabilities | Manual entry, auto-updated values |
| **INDEXES** | Financial indexes (CDI, SELIC, IPCA, etc.) | Import or manual entry |
| **CATEGORIES** | Transaction classification rules | Add/edit mappings as needed |
| **DASHBOARD** | Executive view with KPIs and charts | Read-only, refresh to update |

---

## Daily Operations

### Importing New Transactions

**Frequency**: Monthly or as needed

1. **Export data from your banks/cards**
   - Download transaction exports (CSV or Excel format)
   - Save to the locations configured in FILES PATHS

2. **Run Full Import**
   - Click the "Full Import" button (or run `RunFullImport` macro)
   - Wait for completion (progress shown in status bar)
   - Review completion message

3. **Review imported data**
   - Check BANKS worksheet for new transactions
   - Check CARDS worksheet for new charges
   - Verify dates and amounts look correct

### Classifying Transactions

**Purpose**: Categorize transactions for reporting and analysis

**Automatic Classification**:
- Runs automatically during Full Import
- Uses keyword matching from CATEGORIES sheet
- Two-pass approach: exact match first, then partial match

**Manual Classification**:
1. Find unclassified transactions (Category = "UNCLASSIFIED")
2. Determine appropriate category
3. Add keyword mapping to CATEGORIES sheet
4. Re-run classification: `ClassifyAllTransactions`

**Adding Category Mappings**:
1. Go to CATEGORIES worksheet
2. Add new row with:
   - **Category**: Main category (e.g., "Food", "Transportation")
   - **Subcategory**: Specific type (e.g., "Restaurants", "Groceries")
   - **Keywords**: Pipe-separated keywords (e.g., "RESTAURANT|IFOOD|RAPPI")
   - **Date Added**: Automatically filled

**Example**:
```
Category: Entertainment
Subcategory: Streaming Services
Keywords: NETFLIX|SPOTIFY|APPLE MUSIC|DISNEY PLUS
```

**Best Practices**:
- Use UPPERCASE for keywords
- Include common variations and abbreviations
- Be specific to avoid false matches
- Test with `ClassifyAllTransactions` after adding

### Viewing Unclassified Transactions

**Method 1 - Macro**:
- Run `ShowUnclassifiedTransactions` macro
- Review pop-up list of unclassified items

**Method 2 - Manual**:
- Filter BANKS sheet: Category = "UNCLASSIFIED"
- Filter CARDS sheet: Category = "UNCLASSIFIED"

---

## Investment Management

### Understanding Investment Correlation

**Concept**:
- When you invest money, it leaves your bank account and enters an investment
- When you redeem, it leaves the investment and enters your bank account
- The system automatically correlates these movements

**Correlation Rules**:
- **Application**: Negative value in BANK ↔ Positive value in INVESTMENT
- **Redemption**: Negative value in INVESTMENT ↔ Positive value in BANK
- Values must match (within 1 cent tolerance)
- Dates must be within ±3 days

**Correlation Status**:
- **MATCHED**: Successfully correlated with bank transaction
- **UNMATCHED**: No matching bank transaction found

**Reviewing Correlations**:
1. Go to INVESTMENTS worksheet
2. Check "Correlation Status" column
3. For MATCHED items, see Correlation ID
4. For UNMATCHED items, investigate why:
   - Missing bank transaction?
   - Date outside tolerance?
   - Amount mismatch?

**Manual Correlation** (if needed):
- Identify the bank and investment transactions
- Note the Correlation ID pattern
- Add matching ID to both records manually

### Managing OPUS Investments

**Purpose**: Track investments outside the banking system

**Fields**:
- **Company**: Investment name or company
- **Investment Cost**: Original amount invested
- **Capital Cost (%)**: Annual cost of capital (optional)
- **Updated Cost**: Auto-calculated based on indexes
- **Start Date**: Investment start date
- **Currency**: BRL or USD
- **Prior Management Value (USD)**: Previous value if applicable
- **Accumulated Value**: Total current value

**Updating OPUS Values**:
- Run `UpdateOPUSValues` macro
- Uses appropriate index based on currency:
  - BRL → CDI
  - USD → FED_FUNDS
- Calculates time-weighted returns

---

## Debt Management

### Tracking Debts

**Purpose**: Monitor personal loans and liabilities with capital cost adjustments

**Fields**:
- **Creditor**: Name of lender
- **Interest Rate (%)**: Annual interest rate
- **Amount Paid**: Original debt amount
- **Updated Amount**: Auto-calculated current value
- **Currency**: BRL or USD
- **Start Date**: Debt origination date

**Updating Debt Values**:
- Run `UpdateDebtValues` macro
- Calculates current value using:
  - BRL debts → CDI index
  - USD debts → FED_FUNDS index
- Accounts for time value of money

**Adding New Debt**:
1. Go to DEBTS worksheet
2. Add new row with creditor info
3. Fill in amount, currency, start date
4. Run `UpdateDebtValues` to calculate updated amount

---

## Financial Indexes

### Supported Indexes

| Index | Purpose | Currency | Update Frequency |
|-------|---------|----------|------------------|
| **CDI** | Brazilian interbank rate | BRL | Daily (business days) |
| **SELIC** | Brazilian base rate | BRL | Daily |
| **IPCA** | Brazilian inflation index | BRL | Monthly |
| **USD/BRL** | Dollar exchange rate | Both | Daily |
| **FED_FUNDS** | US Federal Funds rate | USD | Daily |

### Index Structure

Each index entry has:
- **Index Name**: CDI, SELIC, IPCA, USD/BRL, FED_FUNDS
- **Date**: Date of the value
- **Index Value (%)**: The rate as percentage
- **Cumulative Factor**: Auto-calculated compound factor from first date

### Importing Index Data

**Method 1 - CSV Import**:
```
CSV Format:
Date,Value
2024-01-01,11.65
2024-01-02,11.70
...
```

Run macro:
```vba
Call ImportIndexFromCSV("/path/to/cdi_data.csv", "CDI")
```

**Method 2 - Manual Entry**:
1. Go to INDEXES worksheet
2. Add new row:
   - Index Name: CDI
   - Date: 2024-01-15
   - Index Value (%): 11.65
   - Cumulative Factor: (leave blank)
3. Run `UpdateAllIndexes` to calculate factors

**Method 3 - Macro Entry**:
```vba
Call AddIndexEntry("CDI", #1/15/2024#, 11.65)
```

### Updating Indexes

**Recommended Frequency**: Monthly minimum, weekly preferred

1. Download latest index data from official sources:
   - **CDI/SELIC**: Banco Central do Brasil (BCB)
   - **IPCA**: IBGE
   - **USD/BRL**: BCB
   - **FED_FUNDS**: Federal Reserve

2. Import using CSV or manual entry

3. Run `UpdateAllIndexes` to recalculate cumulative factors

---

## Dashboard Usage

### Accessing the Dashboard

1. Click on **DASHBOARD** worksheet tab
2. View executive summary and KPIs
3. Use filters to drill down

### Dashboard Sections

#### 1. Filters (Top Section)

**Available Filters**:
- **Year**: Select specific year or current
- **Month**: Select month or "All"
- **Institution**: Filter by bank/card or "All"
- **Currency**: BRL, USD, or "All"

**Note**: Filters are for display only. To apply filters, modify values and refresh dashboard.

#### 2. Executive KPIs

**Total Income**:
- Sum of all positive values (inflows)
- From BANKS and INVESTMENTS
- Based on current filters

**Total Expenses**:
- Sum of all negative values (outflows)
- From BANKS, CARDS, INVESTMENTS
- Based on current filters

**Balance**:
- Total Income - Total Expenses
- Net cash flow for the period

#### 3. Consolidated Tables

**Consolidated Cash**:
- All bank and investment transactions
- Grouped by: Bank, Month, Year, Kind
- Shows: Inflows, Outflows, Net Value

**Consolidated Cards**:
- All credit card transactions
- Grouped by: Bank, Month, Year
- Shows: Total Value

**Consolidated Transactions**:
- All classified transactions
- Grouped by: Bank, Month, Year, Category
- Shows: Total Value per category

**Consolidated Net Debts**:
- All debts with current values
- Shows: Opening Balance, Payments, Updated Balance

### Refreshing the Dashboard

**When to refresh**:
- After importing new data
- After classifying transactions
- After updating indexes
- After manual data changes

**How to refresh**:
- Run `RefreshDashboard` macro
- Or run `RunQuickRefresh` for full recalculation

**Refresh time**: 5-30 seconds depending on data volume

---

## Health Check System

### Running Health Checks

**Purpose**: Validate data integrity and identify issues

**How to run**:
1. Click "Health Check" button
2. Or run `RunFullHealthCheck` macro
3. Review report

### Health Check Categories

#### 1. Workbook Structure
- Verifies all required worksheets exist
- **Status**: PASS/FAIL

#### 2. Imported Data
- Checks that data exists in BANKS, CARDS, INVESTMENTS
- **Status**: PASS (data exists) / WARNING (no data)

#### 3. Transaction Classification
- Counts classified vs unclassified transactions
- **Status**:
  - PASS: All classified
  - WARNING: <10% unclassified
  - FAIL: >10% unclassified

#### 4. Investment Correlation
- Checks correlation status
- Calculates unmatched balance
- **Status**:
  - PASS: All matched
  - WARNING: <20% unmatched
  - FAIL: >20% unmatched

#### 5. Index Data Availability
- Verifies index data exists and is current
- **Status**:
  - PASS: Updated within 7 days
  - WARNING: 7-30 days old
  - FAIL: >30 days old or missing

#### 6. Data Integrity
- Checks for invalid dates and values
- **Status**: PASS (all valid) / FAIL (issues found)

### Interpreting Results

**Report Format**:
```
[PASS] Workbook Structure
    All required worksheets exist

[WARNING] Transaction Classification
    Some transactions unclassified
    Details: 15 banks, 8 cards

[FAIL] Index Data Availability
    Index data is outdated
    Details: Last update: 2024-01-15 (45 days ago)
```

**Action Items**:
- **PASS**: No action needed
- **WARNING**: Should address but not critical
- **FAIL**: Requires immediate attention

---

## Common Workflows

### Monthly Close Process

1. **Export bank/card data** (monthly statements)
2. **Update FILES PATHS** if any paths changed
3. **Run Full Import** (`RunFullImport`)
4. **Review unclassified** transactions
5. **Add category mappings** as needed
6. **Re-classify** (`ClassifyAllTransactions`)
7. **Update indexes** with latest data
8. **Run Health Check** (`RunFullHealthCheck`)
9. **Review Dashboard**
10. **Export reports** or take screenshots for records

### Quarterly Investment Review

1. **Import latest investment data**
2. **Run correlation** (`CorrelateInvestmentsWithBanks`)
3. **Review INVESTMENTS sheet** for unmatched items
4. **Manually correlate** if needed
5. **Update OPUS values** (`UpdateOPUSValues`)
6. **Calculate returns** (compare updated vs original costs)
7. **Update Dashboard**

### Annual Budget Planning

1. **Set Dashboard filters** to previous year
2. **Review Consolidated Transactions** by category
3. **Calculate monthly averages** per category
4. **Export to separate budget worksheet**
5. **Set targets** for current year
6. **Create tracking vs budget** (custom formulas)

---

## Tips and Best Practices

### Data Quality

1. **Consistent file formats**: Use same export format every time
2. **Clean descriptions**: Remove special characters if causing issues
3. **Date formats**: Ensure dates parse correctly (DD/MM/YYYY or YYYY-MM-DD)
4. **Currency consistency**: Always specify BRL or USD

### Performance

1. **Clear old data periodically**: Archive data older than 2 years to separate workbook
2. **Run Quick Refresh** instead of Full Import when only updating calculations
3. **Disable screen updating**: Already done in macros
4. **Close other workbooks**: Improves Excel performance

### Classification Accuracy

1. **Start broad, refine later**: Begin with general categories
2. **Review monthly**: Check new merchants and add mappings
3. **Use specific keywords**: "RESTAURANT ABC" better than just "RESTAURANT"
4. **Avoid over-matching**: Don't make keywords too generic

### Backup and Security

1. **Regular backups**: Save monthly copies with date suffix
2. **Secure storage**: Keep in encrypted folder or secure cloud storage
3. **Password protect**: Add workbook password in Excel
4. **VBA protection**: Protect VBA project to prevent accidental changes

---

## Troubleshooting

### Issue: Import returns no data

**Possible causes**:
- Incorrect file path
- Empty source file
- Wrong file format
- Column structure mismatch

**Solutions**:
- Verify path in FILES PATHS sheet
- Open source file manually to check data
- Ensure CSV has headers in row 1
- Check column order matches expected format

### Issue: Classification not working

**Possible causes**:
- No keywords in CATEGORIES sheet
- Keywords don't match descriptions
- Case sensitivity (shouldn't be an issue)

**Solutions**:
- Add keyword mappings
- Check actual description text in BANKS/CARDS
- Use `ShowUnclassifiedTransactions` to see patterns
- Add more variations of keywords

### Issue: Correlation failing

**Possible causes**:
- Dates outside tolerance (±3 days)
- Values don't match
- Same sign (both positive or both negative)
- Missing transactions

**Solutions**:
- Check date difference
- Verify amounts match exactly
- Ensure opposite signs (out vs in)
- Import missing bank transactions

### Issue: Dashboard shows zeros

**Possible causes**:
- Dashboard not refreshed
- No data imported
- Filters excluding all data

**Solutions**:
- Run `RefreshDashboard`
- Check BANKS/CARDS have data
- Reset filters to "All"

---

## Advanced Features

### Custom Formulas

You can add custom formulas in any worksheet. For example:

**Monthly expenses by category**:
```excel
=SUMIFS(BANKS!D:D, BANKS!E:E, "Food", BANKS!B:B, ">="&DATE(2024,1,1))
```

**Investment return percentage**:
```excel
=(Updated Cost - Investment Cost) / Investment Cost
```

### Pivot Tables

Create pivot tables from:
- BANKS data: Analyze by category, bank, month
- CARDS data: Analyze by merchant, installments
- INVESTMENTS: Review by institution

### Charts

Add custom charts to DASHBOARD or separate "Reports" sheet:
- Line chart: Income vs Expenses over time
- Pie chart: Expenses by category
- Bar chart: Bank comparison
- Stacked area: Cash flow trends

---

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Option + F8` | Open Macros dialog |
| `Option + F11` | Open VBA Editor |
| `Cmd + Shift + F` | Toggle worksheet filters |
| `Cmd + T` | Create table |
| `Cmd + Home` | Go to cell A1 |

---

## Appendix: Macro Reference

### Main Macros

| Macro | Purpose | When to Use |
|-------|---------|-------------|
| `InitializeWorkbook` | Create all worksheets and headers | First time setup |
| `RunFullImport` | Import all data and process | Monthly, after exporting new data |
| `RunQuickRefresh` | Refresh calculations and dashboard | After manual changes |
| `RefreshDashboard` | Update dashboard only | After filtering or viewing changes |
| `ClassifyAllTransactions` | Re-classify all transactions | After adding categories |
| `ShowUnclassifiedTransactions` | List unclassified items | When reviewing classification |
| `RunFullHealthCheck` | Validate data integrity | Monthly or when troubleshooting |
| `UpdateAllIndexes` | Recalculate index factors | After importing index data |
| `UpdateDebtValues` | Recalculate debt values | Monthly or after index updates |
| `UpdateOPUSValues` | Recalculate investment values | Monthly or quarterly |

---

## Getting Help

### Built-in Help

1. **Health Check**: Run `RunFullHealthCheck` to identify issues
2. **VBA Comments**: Review VBA code comments for technical details
3. **This Manual**: Refer to relevant sections

### External Resources

- **Excel for Mac Help**: help.excel.microsoft.com
- **VBA Reference**: Microsoft Office VBA documentation
- **Financial Index Sources**:
  - BCB (Brazil): www.bcb.gov.br
  - Federal Reserve (USA): www.federalreserve.gov

---

**End of User Manual**

For technical details and VBA architecture, see **TECHNICAL_REFERENCE.md**
