# [MY LIFE] - Production Financial Management System

A comprehensive Excel-based financial management system for entrepreneurs, family offices, and personal finance management on MacOS.

## 🎯 Overview

**[MY LIFE]** is a production-grade Excel workbook that consolidates all financial information, automates imports, normalizes data, calculates metrics, and presents an executive dashboard. Built specifically for Excel on MacOS using VBA.

## ✨ Key Features

### Data Management
- ✅ **Automated imports** from multiple banks and credit cards
- ✅ **Transaction normalization** across different formats
- ✅ **Investment correlation** with bank movements
- ✅ **Debt tracking** with capital cost calculations
- ✅ **External investments** (OPUS) management

### Intelligence
- 🧠 **Smart classification** using exact and fuzzy matching
- 🔗 **Automatic correlation** of investment transactions
- 📊 **Capital cost calculations** using historical indexes
- ⚡ **Real-time validation** and health checks

### Reporting
- 📈 **Executive dashboard** with KPIs and filters
- 💰 **Consolidated views** of cash, cards, transactions, debts
- 🎨 **Named ranges** for custom reporting
- ✅ **Health check system** for data integrity

### Financial Indexes
- 📉 **CDI** - Brazilian interbank rate
- 📉 **SELIC** - Brazilian base rate
- 📉 **IPCA** - Brazilian inflation index
- 💱 **USD/BRL** - Dollar exchange rate
- 📉 **FED_FUNDS** - US Federal Funds rate

## 📋 Requirements

### Software
- **Microsoft Excel for Mac** (Office 365, 2019, or 2021+)
- **VBA enabled** in Excel preferences
- **MacOS** 10.14 or higher

### Skills
- Basic Excel knowledge
- Ability to export bank/card data
- Understanding of personal finances

## 🚀 Quick Start

### 1. Create Workbook

```bash
# Download the project
git clone [repository-url]
cd my-life-workbook
```

### 2. Set Up Excel

1. Open Excel
2. Create new **Macro-Enabled Workbook** (.xlsm)
3. Save as `MY_LIFE.xlsm`

### 3. Import VBA Modules

1. Press `Option + F11` to open VBA Editor
2. Go to **File > Import File...**
3. Import all `.bas` files from `vba-modules/` folder in order:
   - modConfig.bas
   - modUtilities.bas
   - modImportBanks.bas
   - modImportCards.bas
   - modImportInvestments.bas
   - modClassification.bas
   - modCapitalCost.bas
   - modIndexes.bas
   - modDashboard.bas
   - modHealthCheck.bas
   - modMain.bas

### 4. Initialize

1. Run macro: `InitializeWorkbook`
2. Configure file paths in **FILES PATHS** sheet
3. Set up categories in **CATEGORIES** sheet
4. Import index data in **INDEXES** sheet

### 5. Import Data

1. Export transactions from your banks/cards
2. Save to configured paths
3. Run macro: `RunFullImport`
4. Review **DASHBOARD** sheet

## 📁 Project Structure

```
my-life-workbook/
├── vba-modules/           # VBA source code
│   ├── modConfig.bas      # Configuration management
│   ├── modUtilities.bas   # Common utilities
│   ├── modImportBanks.bas # Bank import logic
│   ├── modImportCards.bas # Card import logic
│   ├── modImportInvestments.bas # Investment correlation
│   ├── modClassification.bas # Transaction classification
│   ├── modCapitalCost.bas # Capital cost calculations
│   ├── modIndexes.bas     # Financial indexes
│   ├── modDashboard.bas   # Dashboard aggregation
│   ├── modHealthCheck.bas # Validation system
│   └── modMain.bas        # Main orchestration
├── templates/             # Sample data templates
├── docs/                  # Documentation
│   ├── SETUP_GUIDE.md     # Detailed setup instructions
│   ├── USER_MANUAL.md     # Complete user guide
│   └── TECHNICAL_REFERENCE.md # VBA architecture docs
└── README.md             # This file
```

## 📊 Worksheet Structure

| Worksheet | Purpose |
|-----------|---------|
| **FILES PATHS** | Configure data source locations |
| **FILES STRUCTURE** | Define expected column structures |
| **BANKS** | Checking account transactions |
| **CARDS** | Credit card transactions |
| **INVESTMENTS** | Investment movements |
| **OPUS** | External investments |
| **DEBTS** | Personal loans and liabilities |
| **INDEXES** | Financial indexes with cumulative factors |
| **CATEGORIES** | Transaction classification rules |
| **DASHBOARD** | Executive view with KPIs |

## 🔧 Core Workflows

### Monthly Import
```
Export Data → Configure Paths → Run Full Import → Review Dashboard
```

### Classification
```
Import → Auto-Classify → Review Unclassified → Add Mappings → Re-Classify
```

### Investment Correlation
```
Import Banks + Investments → Auto-Correlate → Review Unmatched → Manual Match
```

### Capital Cost Updates
```
Import Indexes → Update Cumulative Factors → Update Debts/OPUS → Refresh Dashboard
```

## 🎨 Main Macros

| Macro | Purpose |
|-------|---------|
| `InitializeWorkbook` | Create all worksheets and headers |
| `RunFullImport` | Import and process all data |
| `RunQuickRefresh` | Refresh calculations and dashboard |
| `RefreshDashboard` | Update dashboard data |
| `ClassifyAllTransactions` | Re-classify all transactions |
| `ShowUnclassifiedTransactions` | List unclassified items |
| `RunFullHealthCheck` | Validate data integrity |

## 🛡️ Data Integrity

### Validation Features
- ✅ Workbook structure validation
- ✅ Import success verification
- ✅ Classification completeness check
- ✅ Correlation balance verification
- ✅ Index data freshness check
- ✅ Data type validation

### Health Check System
Run `RunFullHealthCheck` to:
- Verify all imports completed
- Check classification status
- Validate correlations
- Ensure index data is current
- Detect orphan records

## 💡 Best Practices

### Data Management
1. **Consistent formats** - Use same export format every time
2. **Regular backups** - Save monthly copies with date suffix
3. **Clean data** - Remove special characters if causing issues

### Classification
1. **Start broad** - Begin with general categories, refine later
2. **Review monthly** - Check new merchants and add mappings
3. **Use specific keywords** - More specific = better accuracy

### Performance
1. **Archive old data** - Move data older than 2 years to separate workbook
2. **Use Quick Refresh** - Instead of Full Import when only updating calculations
3. **Regular maintenance** - Run health checks monthly

## 🔒 Security

- **Password protection** - Add workbook password in Excel
- **VBA protection** - Protect VBA project to prevent accidental changes
- **Secure storage** - Keep in encrypted folder or secure cloud storage
- **No cloud sync** - Avoid storing sensitive data in public cloud

## 📚 Documentation

Comprehensive documentation included:

- **[SETUP_GUIDE.md](docs/SETUP_GUIDE.md)** - Step-by-step setup instructions
- **[USER_MANUAL.md](docs/USER_MANUAL.md)** - Complete user guide with workflows
- **[TECHNICAL_REFERENCE.md](docs/TECHNICAL_REFERENCE.md)** - VBA architecture and customization

## 🛠️ Customization

### Adding New Banks
1. Add enum in `modConfig.bas`
2. Update `GetSourceName` function
3. Add column mapping functions
4. Configure in FILES PATHS sheet

### Custom Categories
1. Add to CATEGORIES worksheet
2. Use pipe-separated keywords
3. Run classification

### Custom Reports
1. Create new worksheet
2. Use formulas referencing named ranges
3. Add pivot tables or charts
4. Link to dashboard if needed

## ⚠️ Limitations

### MacOS Excel Constraints
- ❌ No ActiveX controls (use Form Controls)
- ❌ Limited web query capabilities (manual index imports)
- ⚠️ Slower VBA execution than Windows
- ⚠️ Some Windows-specific VBA features unavailable

### Design Choices
- Single workbook architecture (easier deployment)
- VBA-only automation (no external dependencies)
- Manual index updates (more reliable on MacOS)
- Form controls for buttons (MacOS compatible)

## 🐛 Troubleshooting

### Common Issues

**Macros won't run**
- Enable macros in Excel preferences
- Trust VBA project access

**Import returns no data**
- Verify file paths in FILES PATHS
- Check source file format
- Ensure CSV has headers

**Classification fails**
- Add keywords to CATEGORIES
- Run `ShowUnclassifiedTransactions`
- Check description text format

**Dashboard shows zeros**
- Run `RefreshDashboard`
- Verify data imported successfully
- Reset filters to "All"

See **USER_MANUAL.md** for detailed troubleshooting.

## 🎯 Roadmap

Future enhancements (contributions welcome):

- [ ] Automated bank API integrations (where available)
- [ ] Machine learning classification
- [ ] Budget vs actual tracking
- [ ] Multi-currency portfolio analysis
- [ ] Tax reporting exports
- [ ] Mobile companion app
- [ ] Cloud sync with encryption

## 📄 License

This project is provided as-is for personal and family office use.

## 🤝 Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create feature branch
3. Add/modify VBA modules
4. Update documentation
5. Submit pull request

## 📞 Support

- **Documentation**: See `docs/` folder
- **Health Check**: Run `RunFullHealthCheck` macro
- **Issues**: Check troubleshooting sections

## ✅ Version

- **Version**: 1.0
- **Excel Compatibility**: Excel for Mac (Office 365, 2019, 2021+)
- **VBA Version**: 7.0+
- **Last Updated**: January 2026

---

**Built for entrepreneurs who need a production-grade financial system without complex software.**

Made with ❤️ for MacOS Excel
