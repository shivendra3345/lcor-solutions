# CSV Report Viewer - Quick Start Guide

## 5-Minute Setup

### Step 1: Install Dependencies
```bash
cd /Users/shivendrapatel/Documents/projects/lcor-solutions
npm install
```

### Step 2: Verify PnP Initialization
The `ReportsWebPart.ts` already has PnP initialization in the `onInit()` method. No additional setup needed!

### Step 3: Prepare Your CSV File
Create a CSV file in SharePoint:

1. Go to your SharePoint site
2. Navigate to: **Shared Documents > Reports** folder
3. Upload a CSV file with the following format:

```csv
Name,Department,Salary,StartDate
John Smith,Sales,75000,2022-01-15
Jane Doe,Engineering,95000,2021-06-20
Bob Johnson,Marketing,65000,2023-03-10
Alice Williams,Sales,72000,2022-11-05
```

### Step 4: Start Development Server
```bash
npm start
```

The component is already integrated in `Reports.tsx` and will automatically:
- Load CSV files from "Shared Documents" > "Reports" folder
- Display available files in a dropdown
- Allow you to select different chart types
- Visualize your data

## Component Location

The CSV Report Viewer is located at:
```
src/webparts/reports/components/CSVReportViewer.tsx
```

And is already imported and used in:
```
src/webparts/reports/components/Reports.tsx
```

## File Structure

```
src/webparts/reports/
├── components/
│   ├── CSVReportViewer.tsx           (Main component)
│   ├── CSVReportViewer.module.scss   (Component styles)
│   ├── ChartComponent.tsx            (Chart rendering)
│   ├── ChartComponent.module.scss    (Chart styles)
│   ├── Reports.tsx                   (Main page - already integrated)
│   ├── IReportsProps.ts              (Props interface - updated)
│   └── Reports.module.scss
├── services/
│   └── CSVDataService.ts             (CSV fetching & parsing service)
└── ReportsWebPart.ts                 (Web part class - PnP initialized)
```

## Customization Examples

### Example 1: Change Default Library

Edit `Reports.tsx` line 17:

**Current:**
```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  isDarkTheme={isDarkTheme}
/>
```

**Change to:**
```typescript
<CSVReportViewer
  libraryName="Documents"
  folderPath="Reports/2024"
  fileName="sales_report.csv"
  isDarkTheme={isDarkTheme}
/>
```

### Example 2: Add Pre-loaded CSV File

```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  fileName="monthly_sales.csv"
  isDarkTheme={isDarkTheme}
/>
```

### Example 3: Custom Styling

Edit `CSVReportViewer.module.scss` to customize colors, fonts, and layout.

## Available Chart Types

1. **Bar Chart** - Good for comparing values
2. **Line Chart** - Good for trends over time
3. **Pie Chart** - Good for percentages
4. **Doughnut Chart** - Modern pie chart variant

## Testing

### Manual Testing Steps:

1. **Test File Upload:**
   - Upload a CSV to "Shared Documents > Reports"
   - Verify component loads the file

2. **Test Chart Rendering:**
   - Select different X/Y columns
   - Switch between chart types
   - Verify data displays correctly

3. **Test Error Handling:**
   - Enter wrong library name
   - Select empty CSV file
   - Verify error messages display

### Test CSV File

Save this as `test_data.csv` and upload to SharePoint:

```csv
Quarter,Sales,Expenses,Profit
Q1,150000,95000,55000
Q2,165000,100000,65000
Q3,180000,105000,75000
Q4,200000,110000,90000
```

## Troubleshooting

### Issue: "No CSV files found"
- Verify library name matches exactly (including spaces)
- Check user has read permissions
- Ensure CSV files exist in the specified folder

### Issue: Chart not displaying
- Verify Y-axis column contains numbers
- Check that data preview table shows rows
- Try a different chart type

### Issue: Application won't start
- Run `npm install` to ensure all dependencies are installed
- Clear `node_modules` and reinstall: `rm -rf node_modules && npm install`
- Check Node.js version: `node --version` (should be 22.14.0+)

## Build for Production

```bash
npm run build
npm run package-solution
```

The solution package will be created in the `sharepoint/solution/` folder.

## Next Steps

1. ✅ Test with your CSV data
2. ✅ Customize styling in SCSS modules
3. ✅ Add more CSV files to SharePoint
4. ✅ Configure the web part in a SharePoint page
5. ✅ Deploy to production

## Key Files Reference

| File | Purpose |
|------|---------|
| `CSVDataService.ts` | Fetches and parses CSV from SharePoint |
| `CSVReportViewer.tsx` | Main UI component |
| `ChartComponent.tsx` | Chart rendering logic |
| `Reports.tsx` | Integrates CSVReportViewer |
| `CSV_REPORT_VIEWER_GUIDE.md` | Full documentation |

## Performance Tips

- Split large CSV files (>10MB) into smaller chunks
- Use specific folder paths to limit file browsing
- Archive old reports for faster file listings
- Consider pagination for data tables with 1000+ rows

## Support Files

- 📖 `CSV_REPORT_VIEWER_GUIDE.md` - Complete documentation
- 📋 This file - Quick start guide

Enjoy your new CSV Report Viewer! 🎉
