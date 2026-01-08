# [MY LIFE] Quick Reference Guide

## 🚀 Quick Setup (5 Minutes)

1. **Create workbook**: New Excel file → Save as `.xlsm`
2. **Import VBA**: Press `Option+F11` → Import all `.bas` files
3. **Initialize**: Run `InitializeWorkbook` macro
4. **Configure**: Set paths in FILES PATHS sheet
5. **Import**: Run `RunFullImport` macro

## 🎯 Essential Macros

| Macro | Keyboard | Purpose |
|-------|----------|---------|
| `InitializeWorkbook` | - | First-time setup |
| `RunFullImport` | - | Monthly data import |
| `RunQuickRefresh` | - | Update calculations |
| `RefreshDashboard` | - | Update dashboard |
| `RunFullHealthCheck` | - | Validate data |

**Access macros**: Press `Option + F8`

## 📊 Worksheet Quick Guide

| Sheet | What It Does | Your Action |
|-------|--------------|-------------|
| FILES PATHS | File locations | ✏️ Configure once |
| BANKS | Bank transactions | 👀 Review monthly |
| CARDS | Card transactions | 👀 Review monthly |
| INVESTMENTS | Investments | 👀 Check correlations |
| CATEGORIES | Classification rules | ✏️ Add as needed |
| DASHBOARD | Executive view | 👀 Main reporting |

## 🔄 Monthly Workflow

```
1. Export bank/card data
   ↓
2. Save to configured paths
   ↓
3. Run: RunFullImport
   ↓
4. Review: ShowUnclassifiedTransactions
   ↓
5. Add category mappings
   ↓
6. Run: ClassifyAllTransactions
   ↓
7. Check: DASHBOARD
```

## 📝 Data File Formats

### Bank/Investment CSV
```csv
Date,Description,Value
2024-01-15,DESCRIPTION,1000.00
```

### Card CSV
```csv
Card Number,Purchase Date,Category,Description,Installment,Value
****1234,2024-01-15,Shopping,STORE NAME,1/1,100.00
```

### Index CSV
```csv
Date,Value
2024-01-15,11.65
```

## 🎨 Classification

**Add category mapping**:
1. Go to CATEGORIES sheet
2. Add row: `Category | Subcategory | KEYWORD1|KEYWORD2`
3. Run `ClassifyAllTransactions`

**Example**:
```
Food | Restaurants | RESTAURANT|IFOOD|RAPPI|DELIVERY
```

## 🔍 Health Check Status

| Icon | Meaning | Action |
|------|---------|--------|
| ✅ PASS | All good | None needed |
| ⚠️ WARNING | Should fix | Review and address |
| ❌ FAIL | Must fix | Fix immediately |

## 📈 Dashboard KPIs

- **Total Income**: All inflows (positive values)
- **Total Expenses**: All outflows (negative values)
- **Balance**: Income - Expenses

## 🛠️ Common Fixes

**No data imported?**
→ Check FILES PATHS sheet → Verify file exists → Check file format

**Not classifying?**
→ Add keywords to CATEGORIES → Run ClassifyAllTransactions

**Dashboard shows 0?**
→ Run RefreshDashboard → Check filters set to "All"

**Correlation failed?**
→ Check dates ±3 days → Verify opposite signs → Check amounts match

## 💡 Pro Tips

1. **Backup monthly**: Save copy as `MY_LIFE_2024_01.xlsm`
2. **Test with samples**: Use files in `templates/` folder first
3. **Start simple**: Begin with one bank, expand later
4. **Review weekly**: Check dashboard every week
5. **Update indexes**: Monthly minimum, weekly preferred

## 🔐 Security Checklist

- [ ] Password protect workbook
- [ ] Protect VBA project
- [ ] Store in encrypted location
- [ ] Regular backups
- [ ] Never share with passwords

## 📞 Troubleshooting Shortcuts

**Can't run macros?**
→ Excel Preferences → Security → Enable all macros

**VBA Editor won't open?**
→ Excel Preferences → Ribbon → Check "Developer"

**Import slow?**
→ Close other workbooks → Disable auto-calculation temporarily

**Errors in VBA?**
→ Check all modules imported → Verify no compile errors

## 🎓 Learning Path

**Week 1**: Setup and basic import
**Week 2**: Classification mastery
**Week 3**: Investment correlation
**Week 4**: Custom reports and charts

## 📚 Documentation Map

- **Quick Start** → This file
- **Detailed Setup** → SETUP_GUIDE.md
- **Daily Use** → USER_MANUAL.md
- **Customization** → TECHNICAL_REFERENCE.md

## ⌨️ Keyboard Shortcuts

| Mac Shortcut | Action |
|--------------|--------|
| `⌥ F8` | Open Macros |
| `⌥ F11` | VBA Editor |
| `⌘ S` | Save |
| `⌘ ⇧ F` | Toggle Filters |
| `⌘ Home` | Go to A1 |

## 🎯 Performance Tips

**Slow workbook?**
- Archive data > 2 years old
- Use RunQuickRefresh instead of RunFullImport
- Limit open worksheets
- Clear unused named ranges

## 🔄 Update Frequency

| Task | Frequency |
|------|-----------|
| Import transactions | Monthly |
| Update indexes | Weekly |
| Classify new merchants | As needed |
| Health check | Monthly |
| Backup | Monthly |
| Review dashboard | Weekly |

## 📊 Named Ranges (for formulas)

```excel
=Total_Income          ' Total income amount
=Total_Expenses        ' Total expenses amount
=Balance               ' Net balance
```

Use in custom formulas:
```excel
=IF(Balance>0, "Surplus", "Deficit")
```

## 🎨 Custom Buttons

Add macro buttons to any sheet:
1. Developer → Insert → Button
2. Draw on sheet
3. Assign macro
4. Right-click → Edit Text

## 🚨 Error Messages

| Error | Fix |
|-------|-----|
| "File not found" | Check FILES PATHS |
| "Invalid structure" | Run InitializeWorkbook |
| "No data" | Import data first |
| "Macro not found" | Re-import VBA modules |

## ✅ Monthly Checklist

```
□ Export all bank/card statements
□ Save to correct paths
□ Run RunFullImport
□ Review unclassified transactions
□ Add new category mappings
□ Update index data
□ Run RunFullHealthCheck
□ Review dashboard
□ Save monthly backup
□ Document any issues
```

## 🎉 Success Metrics

You're doing it right when:
- ✅ Import completes in < 2 minutes
- ✅ > 90% transactions classified
- ✅ All investments correlated
- ✅ Health check shows all PASS
- ✅ Dashboard updates instantly

## 📱 Mobile Workflow

While workbook is Mac-only, you can:
1. Screenshot dashboard weekly
2. Save to phone for reference
3. Track against budget on-the-go
4. Note transactions to classify later

---

**Remember**: Start small, iterate often, automate gradually.

**Need help?** Check SETUP_GUIDE.md or USER_MANUAL.md
