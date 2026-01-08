# [MY LIFE] Excel Workbook - Complete Setup Guide

## Overview

This guide will help you set up and deploy the **[MY LIFE]** financial management workbook for Excel on MacOS.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Creating the Workbook](#creating-the-workbook)
3. [Importing VBA Modules](#importing-vba-modules)
4. [Initial Configuration](#initial-configuration)
5. [Data Import Setup](#data-import-setup)
6. [Running Your First Import](#running-your-first-import)
7. [Using the Dashboard](#using-the-dashboard)
8. [Troubleshooting](#troubleshooting)

---

## Prerequisites

### Required Software
- **Microsoft Excel for Mac** (Office 365 or standalone version)
- **VBA enabled** in Excel preferences

### Enable VBA in Excel for Mac

1. Open Excel
2. Go to **Excel > Preferences**
3. Click **Security & Privacy**
4. Under **Developer Macro Settings**, select:
   - ☑ Enable all macros
   - ☑ Trust access to the VBA project object model
5. Click **OK**

### Show Developer Tab

1. Go to **Excel > Preferences**
2. Click **Ribbon & Toolbar**
3. In the right column under **Customize the Ribbon**, check **Developer**
4. Click **Save**

---

## Creating the Workbook

### Step 1: Create New Macro-Enabled Workbook

1. Open Excel
2. Create a new blank workbook
3. **Save As**: Choose **Excel Macro-Enabled Workbook (.xlsm)**
4. Name it: `MY_LIFE.xlsm`
5. Save to a convenient location

### Step 2: Open VBA Editor

1. Press `Option + F11` (or go to **Developer > Visual Basic**)
2. You should see the **Microsoft Visual Basic for Applications** window

---

## Importing VBA Modules

You need to import all VBA modules into your workbook.

### Step 3: Import Each Module

For each `.bas` file in the `vba-modules` folder:

1. In the VBA Editor, go to **File > Import File...**
2. Navigate to the `vba-modules` folder
3. Select the first module file (e.g., `modConfig.bas`)
4. Click **Open**
5. Repeat for all modules in this order:

   **Required modules (in order):**
   1. `modConfig.bas` - Configuration management
   2. `modUtilities.bas` - Utility functions
   3. `modImportBanks.bas` - Bank import logic
   4. `modImportCards.bas` - Card import logic
   5. `modImportInvestments.bas` - Investment import and correlation
   6. `modClassification.bas` - Transaction classification
   7. `modCapitalCost.bas` - Capital cost calculations
   8. `modIndexes.bas` - Financial indexes management
   9. `modDashboard.bas` - Dashboard aggregation
   10. `modHealthCheck.bas` - Health check system
   11. `modMain.bas` - Main orchestration

### Step 4: Verify Module Import

In the VBA Editor's **Project Explorer** (left panel), you should see all modules listed under:
```
VBAProject (MY_LIFE.xlsm)
  └─ Modules
      ├─ modCapitalCost
      ├─ modClassification
      ├─ modConfig
      ├─ modDashboard
      ├─ modHealthCheck
      ├─ modImportBanks
      ├─ modImportCards
      ├─ modImportInvestments
      ├─ modIndexes
      ├─ modMain
      └─ modUtilities
```

### Step 5: Save the Workbook

1. Press `Cmd + S` to save
2. Close the VBA Editor (`Cmd + Q`)

---

## Initial Configuration

### Step 6: Initialize Workbook Structure

1. In Excel, go to **Developer > Macros** (or press `Option + F8`)
2. Select **InitializeWorkbook**
3. Click **Run**
4. Confirm the initialization when prompted
5. Wait for completion message

This will create all required worksheets with proper headers:
- FILES PATHS
- FILES STRUCTURE
- BANKS
- CARDS
- INVESTMENTS
- OPUS
- DEBTS
- INDEXES
- CATEGORIES
- DASHBOARD

### Step 7: Configure File Paths

1. Navigate to the **FILES PATHS** worksheet
2. For each data source you want to use, enter the full file path in column B

**Example:**
```
Source              | File Path
--------------------|---------------------------------------------
ITAU_BANK          | /Users/yourname/Documents/Data/itau_bank.csv
NUBANK_BANK        | /Users/yourname/Documents/Data/nubank.csv
C6_BANK            | /Users/yourname/Documents/Data/c6_bank.csv
ITAU_CARD          | /Users/yourname/Documents/Data/itau_card.csv
INVESTMENTS        | /Users/yourname/Documents/Data/investments.csv
```

**Important Notes:**
- Use **absolute paths** (full path from root)
- Ensure files are in CSV or Excel format
- Leave blank any sources you don't have data for

---

## Data Import Setup

### Step 8: Prepare Your Data Files

Your source files should follow these structures:

#### Bank Files (ITAU_BANK, NUBANK_BANK, C6_BANK, BB_BANK)
```csv
Date,Description,Value
2024-01-15,SALARY DEPOSIT,5000.00
2024-01-16,ATM WITHDRAWAL,-200.00
2024-01-17,TRANSFER TO SAVINGS,-1000.00
```

**Columns:**
- Column 1: Date (DD/MM/YYYY or YYYY-MM-DD)
- Column 2: Transaction Description
- Column 3: Value (positive for inflows, negative for outflows)

#### Card Files (ITAU_CARD, NUBANK_CARD, C6_CARD)
```csv
Card Number,Purchase Date,Category,Description,Installment,Value
****1234,2024-01-15,Shopping,AMAZON MARKETPLACE,1/1,150.00
****1234,2024-01-16,Food,RESTAURANT XYZ,2/3,100.00
```

**Columns:**
- Column 1: Card Number
- Column 2: Purchase Date
- Column 3: Category (raw from bank)
- Column 4: Description
- Column 5: Installment (e.g., "1/1" or "2/12")
- Column 6: Value

#### Investment Files
Same format as bank files (Date, Description, Value)

### Step 9: Set Up Categories

1. Go to the **CATEGORIES** worksheet
2. Add your category mappings

**Example:**
```
Category        | Subcategory    | Keywords
----------------|----------------|---------------------------
Food            | Restaurants    | RESTAURANT|IFOOD|RAPPI
Food            | Groceries      | SUPERMARKET|MERCADO
Transportation  | Uber/Taxi      | UBER|99|TAXI
Transportation  | Gas            | POSTO|GAS|COMBUSTIVEL
Housing         | Rent           | RENT|ALUGUEL
Utilities       | Electricity    | LIGHT|ENERGIA|ELECTRICITY
Utilities       | Water          | WATER|AGUA|SABESP
Entertainment   | Streaming      | NETFLIX|SPOTIFY|AMAZON PRIME
Shopping        | Online         | AMAZON|MERCADO LIVRE
```

**Tips:**
- Use pipe (|) to separate multiple keywords
- Keywords are case-insensitive
- More specific keywords first for better matching

### Step 10: Set Up Indexes (Optional but Recommended)

1. Go to the **INDEXES** worksheet
2. You can manually add index data or import from CSV

**Manual Entry:**
- Index Name: CDI, SELIC, IPCA, USD/BRL, FED_FUNDS
- Date: Date of index value
- Index Value (%): The rate percentage
- Cumulative Factor: Leave blank (auto-calculated)

**Or use macro to import:**
```vba
' In VBA Editor, Immediate Window (Cmd + G), type:
Call ImportIndexFromCSV("/path/to/cdi_data.csv", "CDI")
```

---

## Running Your First Import

### Step 11: Execute Full Import

1. Go to **Developer > Macros** (or `Option + F8`)
2. Select **RunFullImport**
3. Click **Run**
4. Wait for completion (may take several minutes depending on data volume)

**This will:**
- Import all bank transactions
- Import all card transactions
- Import investment transactions
- Correlate investments with bank movements
- Classify all transactions using your category mappings
- Calculate capital costs
- Update dashboard

### Step 12: Review Results

After import completes:

1. Check **BANKS** worksheet - should contain all bank transactions
2. Check **CARDS** worksheet - should contain all card transactions
3. Check **INVESTMENTS** worksheet - look for "MATCHED" in Correlation Status
4. Check **DASHBOARD** worksheet - should show your KPIs

---

## Using the Dashboard

### Step 13: Navigate to Dashboard

1. Click on the **DASHBOARD** worksheet tab
2. You'll see:
   - **Filters**: Year, Month, Institution, Currency
   - **Executive KPIs**: Total Income, Total Expenses, Balance
   - **Consolidated Tables**: All aggregated data

### Step 14: Refresh Dashboard

Whenever you import new data or make changes:

1. Run the macro: `RunQuickRefresh`
2. Or run: `RefreshDashboard` to update just the dashboard

### Step 15: Run Health Check

To validate data integrity:

1. Run macro: `RunFullHealthCheck`
2. Review the report for any issues
3. Address warnings or failures as needed

---

## Creating Buttons for Easy Access

### Step 16: Add Macro Buttons

You can add buttons to any worksheet for easy access:

1. Go to **Developer > Insert**
2. Click **Button** (Form Control)
3. Draw the button on your worksheet
4. In the **Assign Macro** dialog, select the macro (e.g., `RunFullImport`)
5. Click **OK**
6. Right-click the button and choose **Edit Text** to rename it

**Recommended buttons:**
- **Initialize Workbook** → `InitializeWorkbook`
- **Full Import** → `RunFullImport`
- **Quick Refresh** → `RunQuickRefresh`
- **Refresh Dashboard** → `RefreshDashboard`
- **Health Check** → `RunFullHealthCheck`
- **Classify Transactions** → `ClassifyAllTransactions`
- **Show Unclassified** → `ShowUnclassifiedTransactions`

---

## Troubleshooting

### Common Issues

#### 1. **Macros Won't Run**
- **Solution**: Check Excel preferences for macro security settings
- Enable all macros and trust VBA project access

#### 2. **File Not Found Errors**
- **Solution**: Verify file paths in FILES PATHS sheet
- Use absolute paths, not relative paths
- Check file permissions (read access required)

#### 3. **Import Returns No Data**
- **Solution**: Check source file format
- Ensure CSV has headers in row 1
- Verify column structure matches expected format

#### 4. **Transactions Not Classified**
- **Solution**: Add more keywords to CATEGORIES sheet
- Run `ShowUnclassifiedTransactions` to see what's missing
- Add mappings for common merchants

#### 5. **Investment Correlation Fails**
- **Solution**: Check that investment values are opposite sign from bank
- Verify dates are within tolerance (±3 days)
- Manually review unmatched transactions

#### 6. **Dashboard Shows Zeros**
- **Solution**: Run `RefreshDashboard` macro
- Ensure data was imported successfully
- Check that transactions have dates and values

#### 7. **Capital Cost Calculations Wrong**
- **Solution**: Verify INDEX data is populated and up to date
- Run `UpdateAllIndexes` to recalculate cumulative factors
- Check currency fields are correct (BRL/USD)

---

## Regular Maintenance

### Monthly Tasks

1. **Import new data**:
   - Export new transactions from your banks
   - Save to configured file paths
   - Run `RunFullImport`

2. **Update indexes**:
   - Import latest CDI, SELIC, IPCA, USD/BRL, FED_FUNDS data
   - Run `UpdateAllIndexes`

3. **Review unclassified transactions**:
   - Run `ShowUnclassifiedTransactions`
   - Add new category mappings as needed
   - Re-run `ClassifyAllTransactions`

4. **Run health check**:
   - Execute `RunFullHealthCheck`
   - Address any warnings or errors

### Quarterly Tasks

1. **Review categories**:
   - Clean up unused categories
   - Consolidate similar keywords
   - Improve classification accuracy

2. **Validate correlations**:
   - Check investment correlation balance
   - Manually match any remaining unmatched items

3. **Backup workbook**:
   - Save a copy with date suffix
   - Keep historical versions

---

## Advanced Features

### Custom Classification

To add a new category mapping programmatically:

```vba
Call AddCategoryMapping("MERCHANT NAME", "Category", "Subcategory")
```

### Manual Index Entry

```vba
Call AddIndexEntry("CDI", #1/15/2024#, 11.65)
```

### Correlation Balance Check

```vba
Dim balance As Double
balance = GetCorrelationBalance()
Debug.Print "Unmatched investment value: " & balance
```

---

## Support and Customization

### Customizing Column Mappings

If your bank exports have different column orders, modify the functions in `modImportBanks.bas` and `modImportCards.bas`:

```vba
Private Function GetBankDateColumn(sourceType As DataSource) As Long
    Select Case sourceType
        Case DS_ITAU_BANK: GetBankDateColumn = 1  ' Change this number
        ' ... etc
    End Select
End Function
```

### Adding New Data Sources

To add a new bank or card:

1. Add enum value in `modConfig.bas`:
   ```vba
   Public Enum DataSource
       ' ... existing sources
       DS_NEW_BANK = 11
   End Enum
   ```

2. Add to `GetSourceName` function
3. Add entry in FILES PATHS sheet
4. Update import routines

---

## File Structure Reference

```
my-life-workbook/
├── vba-modules/
│   ├── modConfig.bas
│   ├── modUtilities.bas
│   ├── modImportBanks.bas
│   ├── modImportCards.bas
│   ├── modImportInvestments.bas
│   ├── modClassification.bas
│   ├── modCapitalCost.bas
│   ├── modIndexes.bas
│   ├── modDashboard.bas
│   ├── modHealthCheck.bas
│   └── modMain.bas
├── templates/
│   └── (sample data templates)
└── docs/
    ├── SETUP_GUIDE.md (this file)
    ├── USER_MANUAL.md
    └── TECHNICAL_REFERENCE.md
```

---

## Next Steps

Once your workbook is set up:

1. Read the **USER_MANUAL.md** for detailed usage instructions
2. Check **TECHNICAL_REFERENCE.md** for VBA architecture details
3. Customize categories for your spending patterns
4. Set up automated exports from your financial institutions
5. Establish a monthly routine for data updates

---

## Version Information

- **Workbook Version**: 1.0
- **Excel Compatibility**: Excel for Mac (Office 365, 2019, 2021)
- **VBA Version**: 7.0+
- **Last Updated**: 2024

---

**Congratulations!** Your [MY LIFE] workbook is now ready to use. 🎉
