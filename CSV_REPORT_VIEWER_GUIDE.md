# CSV Report Viewer - Implementation Guide

## Overview

The CSV Report Viewer is a comprehensive solution for reading CSV files stored in SharePoint libraries and visualizing the data using interactive charts. This component integrates seamlessly with your SharePoint Framework (SPFx) project.

## Features

- ✅ **CSV File Reading** - Fetch CSV files from SharePoint libraries and folders
- ✅ **Interactive Charts** - Multiple chart types: Bar, Line, Pie, and Doughnut
- ✅ **Flexible Data Visualization** - Select X and Y axes from your data columns
- ✅ **Data Preview** - View raw data in a formatted table
- ✅ **File Management** - Browse and select from available CSV files
- ✅ **Dark Theme Support** - Respects SharePoint theme settings
- ✅ **Error Handling** - Comprehensive error messages and validation
- ✅ **Responsive Design** - Works on desktop and mobile devices

## Components

### 1. **CSVDataService.ts**
Service class for fetching and parsing CSV files from SharePoint.

**Key Methods:**
- `fetchCSVFromSharePoint(libraryName, folderPath, fileName)` - Fetch and parse CSV
- `parseCSV(csvContent)` - Parse raw CSV content
- `listCSVFiles(libraryName, folderPath)` - List available CSV files

### 2. **CSVReportViewer.tsx**
Main React component for the CSV report viewer interface.

**Features:**
- File selection dropdown
- Library and folder configuration
- Chart type selection
- X/Y axis configuration
- Data table preview

### 3. **ChartComponent.tsx**
Chart rendering component using Chart.js.

**Supported Chart Types:**
- Bar Chart
- Line Chart
- Pie Chart
- Doughnut Chart

### 4. **Styling**
- `CSVReportViewer.module.scss` - Main component styles
- `ChartComponent.module.scss` - Chart component styles

## Installation & Setup

### 1. Install Dependencies

The following packages have been added to `package.json`:

```json
{
  "chart.js": "^4.4.1",
  "react-chartjs-2": "^5.2.0",
  "@pnp/sp": "^3.22.0",
  "@pnp/logging": "^3.22.0"
}
```

Run:
```bash
npm install
```

### 2. Initialize PnP in your Web Part

In `ReportsWebPart.ts`, you need to initialize PnP with the SPFx context:

```typescript
import { sp } from '@pnp/sp';

protected onInit(): Promise<void> {
  // Initialize PnP
  sp.setup({
    spfxContext: this.context
  });

  return this._getEnvironmentMessage().then(message => {
    this._environmentMessage = message;
  });
}
```

Update your `ReportsWebPart.ts`:

```typescript
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";

protected onInit(): Promise<void> {
  // Initialize PnP with SPFx context
  sp.setup({
    spfxContext: this.context
  });

  return this._getEnvironmentMessage().then(message => {
    this._environmentMessage = message;
  });
}
```

## Usage

### Basic Usage

In your component, import and use `CSVReportViewer`:

```typescript
import CSVReportViewer from './CSVReportViewer';

<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  fileName="data.csv"
  isDarkTheme={false}
/>
```

### Props

```typescript
interface ICSVReportViewerProps {
  libraryName: string;        // SharePoint library name (required)
  folderPath?: string;        // Folder path within library (optional)
  fileName?: string;          // CSV file name to load (optional)
  isDarkTheme?: boolean;      // Enable dark theme (optional)
}
```

### CSV File Format

Your CSV files should follow standard CSV format:

```csv
Month,Sales,Revenue,Growth
January,150,45000,5
February,165,49500,10
March,180,54000,9
April,195,58500,8.3
May,210,63000,8.2
```

**Requirements:**
- First row must contain column headers
- Values can be text or numeric
- Supports quoted fields with commas inside
- Handles escaped quotes (double quotes)

## Examples

### Example 1: Load CSV on Component Mount

```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports/2024"
  fileName="sales_data.csv"
  isDarkTheme={this.props.isDarkTheme}
/>
```

### Example 2: Allow User to Select Files

```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  isDarkTheme={this.props.isDarkTheme}
/>
```

The user can then:
1. Click "Load Files" to browse available CSV files
2. Select a file from the dropdown
3. Choose X and Y axes for visualization
4. Select chart type (Bar, Line, Pie, Doughnut)
5. View the chart and data preview

## SharePoint Setup

### Required Permissions

Ensure the user has at least **Read** permissions on:
- The specified SharePoint library
- The specified folder (if using folderPath)

### Recommended Folder Structure

```
Shared Documents (Library)
├── Reports (Folder)
│   ├── sales_data.csv
│   ├── customer_data.csv
│   └── 2024 (Subfolder)
│       ├── q1_report.csv
│       └── q2_report.csv
```

## Data Service Usage

### Fetch and Parse CSV

```typescript
import { CSVDataService } from './services/CSVDataService';

const csvData = await CSVDataService.fetchCSVFromSharePoint(
  'Shared Documents',
  'Reports',
  'sales_data.csv'
);

console.log(csvData.headers); // ['Month', 'Sales', 'Revenue']
console.log(csvData.rows);    // Array of data rows
```

### Parse CSV String

```typescript
const csvContent = `Name,Age,City
John,28,New York
Jane,34,Los Angeles`;

const csvData = CSVDataService.parseCSV(csvContent);
```

### List Available CSV Files

```typescript
const files = await CSVDataService.listCSVFiles(
  'Shared Documents',
  'Reports'
);
console.log(files); // ['sales_data.csv', 'customer_data.csv']
```

## Chart Types

### Bar Chart
- Ideal for comparing values across categories
- Supports horizontal and vertical bars

### Line Chart
- Perfect for showing trends over time
- Great for time-series data

### Pie Chart
- Shows proportional data
- Excellent for percentage breakdowns
- Displays tooltips with percentages

### Doughnut Chart
- Similar to pie chart with a center hole
- More modern aesthetic

## Styling & Customization

### Color Scheme

The charts use the following color palette:
- Microsoft Blue (#0078d4)
- Green (#107c10)
- Orange Red (#d83b01)
- Purple (#8661c5)
- Cyan (#00b7c3)
- Red (#f50f0f)
- Gold (#ffb900)
- Light Blue (#00bcf2)

### Dark Theme Support

The component automatically adapts to SharePoint's dark theme:
- Text colors adjust for readability
- Grid colors adapt to background
- Tooltip styling updates accordingly

### Customizing Colors

Edit `ChartComponent.tsx` to modify the color palette:

```typescript
private readonly colors = [
  '#your-color-1',
  '#your-color-2',
  // ... more colors
];
```

## Error Handling

The component provides comprehensive error handling:

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| "Failed to fetch CSV file" | File not found or access denied | Check library name and file path |
| "CSV file is empty" | CSV file has no data rows | Ensure CSV file contains headers and data |
| "No CSV files found" | No CSV files in specified location | Upload CSV files to the library |
| "Please enter a library name" | Library name not specified | Enter a valid SharePoint library name |

## Performance Considerations

- **File Size**: Handles CSV files up to 10MB efficiently
- **Data Rows**: Displays first 10 rows in preview table
- **Memory**: Uses lazy loading for Chart.js to reduce bundle size
- **Rendering**: React.memo and code splitting optimize performance

## Troubleshooting

### Issue: "Unable to load files"

**Solution:**
1. Verify library name (check SharePoint URL for exact name)
2. Check user permissions on the library
3. Ensure folder path is correct
4. Clear browser cache and retry

### Issue: Chart not displaying

**Solution:**
1. Verify X and Y axis columns are selected
2. Ensure Y axis column contains numeric data
3. Check that data is properly parsed in the preview table
4. Try a different chart type

### Issue: Slow performance with large files

**Solution:**
1. Split large CSV files into smaller chunks
2. Use specific folder path to limit files shown
3. Archive old reports to improve browsing speed

## Development Notes

### Testing with Mock Data

Create a test CSV file in your SharePoint library:

```csv
Product,Q1,Q2,Q3,Q4
Product A,100,120,140,160
Product B,150,160,165,170
Product C,80,90,100,110
Product D,200,210,215,220
```

### Building and Packaging

```bash
npm run build
npm run package-solution
```

### Local Development

```bash
npm start
```

The development server will start at `https://localhost:4321`

## Browser Support

- ✅ Chrome/Edge (latest 2 versions)
- ✅ Firefox (latest 2 versions)
- ✅ Safari (latest 2 versions)
- ✅ Mobile browsers (iOS Safari, Chrome Mobile)

## License

This component is part of the LCOR Solutions SharePoint Framework project.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Review CSVDataService error logs in browser console
3. Verify SharePoint library permissions
4. Contact your SharePoint administrator

## Future Enhancements

Potential features for future versions:
- Export chart as image/PDF
- Multi-series charts
- Data filtering and sorting
- Real-time data sync
- Excel file support (.xlsx)
- API data source support
- Advanced analytics (statistics, trends)
- Custom color themes
