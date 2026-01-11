# CSV Report Viewer - SharePoint REST API Configuration

## Overview

The CSV Report Viewer uses SharePoint's REST API to fetch and list files. If you're getting 404 errors, this guide will help you troubleshoot.

## How It Works

### File Fetching

The component uses the SharePoint REST API endpoint:
```
/_api/web/GetFileByServerRelativePath(decodedurl='/LibraryName/FolderPath/FileName')/$value
```

### File Listing

The component uses the REST API endpoint:
```
/_api/web/GetFolderByServerRelativePath(decodedurl='/LibraryName/FolderPath')/Files
```

## Common Issues & Solutions

### Issue 1: Library Name Not Found (404)

**Symptom:** 
```
Error: Failed to list CSV files: HTTP 404
```

**Cause:** Library name doesn't match the exact name in SharePoint

**Solution:**
1. Go to your SharePoint site
2. Click on the library name
3. Copy the exact library name (e.g., "Shared Documents", "Documents", "Reports")
4. Update the component configuration in `Reports.tsx`:

```typescript
<CSVReportViewer
  libraryName="Exact Library Name"  // Must match exactly
  folderPath="Reports"
  isDarkTheme={isDarkTheme}
/>
```

### Issue 2: Folder Path Incorrect (404)

**Symptom:**
```
Error: Failed to fetch CSV file: HTTP 404
```

**Cause:** Folder doesn't exist or path is incorrect

**Solution:**
1. Verify the folder structure in SharePoint
2. Use the exact folder names
3. For nested folders, separate with `/`

**Examples:**
```typescript
// Single folder
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  isDarkTheme={isDarkTheme}
/>

// Nested folders
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports/2024/Q1"
  isDarkTheme={isDarkTheme}
/>

// Root library (no folder)
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath=""  // Empty = root folder
  isDarkTheme={isDarkTheme}
/>
```

### Issue 3: Special Characters in Names

**Symptom:**
```
Error: HTTP 404
```

**Cause:** Library or folder names with special characters (spaces, dashes, etc.)

**Solution:**
The component automatically URL-encodes names, so special characters are handled. However, verify:
- Library/folder names don't have leading/trailing spaces
- Names use standard Unicode characters

**Examples that work:**
- "Shared Documents" ✅
- "Sales Reports" ✅
- "Reports-2024" ✅
- "Q1 Results" ✅

**Examples that may fail:**
- "Reports [Archive]" ⚠️ (brackets may cause issues)
- "Reports (Old)" ⚠️ (parentheses may cause issues)

### Issue 4: File Name Not Found (404)

**Symptom:**
```
Error: Failed to fetch CSV file: HTTP 404
```

**Cause:** CSV file doesn't exist or name is incorrect

**Solution:**
1. Verify the exact file name (including .csv extension)
2. Check file is in the correct folder
3. Try without pre-loading a file and select manually:

```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  // Don't specify fileName - let user select
  isDarkTheme={isDarkTheme}
/>
```

### Issue 5: Permissions Denied (403)

**Symptom:**
```
Error: HTTP 403 - Forbidden
```

**Cause:** User doesn't have permission to access the library/folder

**Solution:**
1. Verify user has at least "Read" permission on the library
2. Check folder permissions are not restricted
3. Ask SharePoint admin to grant permissions

## Debugging

### Enable Console Logging

The component logs API calls to the browser console. To debug:

1. Open Developer Tools (F12)
2. Go to Console tab
3. Look for messages like:
```
Fetching CSV from: https://yoursite.sharepoint.com/_api/web/GetFileByServerRelativePath(...)
Listing files from: https://yoursite.sharepoint.com/_api/web/GetFolderByServerRelativePath(...)
```

4. Copy the URL and test in a new tab or with tools like Postman

### Test API Endpoints

You can test the REST API endpoints directly:

**List files in a folder:**
```
https://yoursite.sharepoint.com/_api/web/GetFolderByServerRelativePath(decodedurl='/Shared%20Documents/Reports')/Files
```

**Get a specific file:**
```
https://yoursite.sharepoint.com/_api/web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/Reports/data.csv')/$value
```

Note: 
- Replace `yoursite` with your actual site URL
- Spaces are encoded as `%20`
- Other special characters are URL-encoded

## SharePoint URL Structure

Different SharePoint site types have different URLs:

### Team Site (Modern)
```
https://tenant.sharepoint.com/sites/sitename
Library: /Shared Documents/Reports
Full path: https://tenant.sharepoint.com/sites/sitename/_api/web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/Reports/file.csv')/$value
```

### Communication Site
```
https://tenant.sharepoint.com/sites/sitename
Library: /Documents/Reports
Full path: https://tenant.sharepoint.com/sites/sitename/_api/web/GetFileByServerRelativePath(decodedurl='/Documents/Reports/file.csv')/$value
```

### Root Site Collection
```
https://tenant.sharepoint.com
Library: /Shared Documents/Reports
Full path: https://tenant.sharepoint.com/_api/web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/Reports/file.csv')/$value
```

## URL Encoding Reference

Special characters are automatically encoded:
- Space → `%20`
- `&` → `%26`
- `#` → `%23`
- `%` → `%25`
- `/` → `/` (kept as-is for path separator)

The component handles this automatically, so you don't need to encode paths.

## Testing Checklist

Before contacting support, verify:

- [ ] Library name matches exactly (case-insensitive but must exist)
- [ ] Folder path exists and is spelled correctly
- [ ] CSV files are in the specified folder
- [ ] Your user account has "Read" permission
- [ ] CSV files have `.csv` extension (lowercase)
- [ ] CSV files are not in a restricted view or subfolder
- [ ] No special characters causing issues in paths
- [ ] Browser console shows the correct API URLs being called

## Example Configurations

### Basic Setup
```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  isDarkTheme={isDarkTheme}
/>
```

CSV files location:
```
https://tenant.sharepoint.com/sites/site/Shared%20Documents/Reports/
```

### Nested Folders
```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports/2024/Q1"
  isDarkTheme={isDarkTheme}
/>
```

CSV files location:
```
https://tenant.sharepoint.com/sites/site/Shared%20Documents/Reports/2024/Q1/
```

### Pre-load Specific File
```typescript
<CSVReportViewer
  libraryName="Shared Documents"
  folderPath="Reports"
  fileName="sales_data.csv"
  isDarkTheme={isDarkTheme}
/>
```

## Contact Support

If you're still getting 404 errors after checking all of the above:

1. **Provide these details:**
   - Your SharePoint site URL
   - Library name (exact)
   - Folder path (if any)
   - CSV file name
   - Your username
   - Error message from console

2. **Check browser console for:**
   - Exact API URL being called
   - HTTP status code
   - Error message details

3. **Try in Postman:**
   - Test the REST API endpoint directly
   - This helps isolate if it's a permissions or path issue

## REST API Reference

### Get File Content
```
GET /_api/web/GetFileByServerRelativePath(decodedurl='/{path}')/$value
```

### List Files in Folder
```
GET /_api/web/GetFolderByServerRelativePath(decodedurl='/{path}')/Files
```

### Filter by File Name
```
GET /_api/web/GetFolderByServerRelativePath(decodedurl='/{path}')/Files?$filter=startswith(Name,'.csv')
```

## More Information

- [SharePoint REST API Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)
- [Files REST API](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-files-and-folders)
- [SPFx HTTP Client Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sp-add-ins/use-the-http-client)

## Quick Fix Checklist

1. Check library name in SharePoint (exact match)
2. Verify folder exists
3. Confirm CSV files are in the folder
4. Check permissions (user has Read access)
5. Look at browser console for actual API URLs
6. Test API URL in new browser tab
7. Try with empty folder path first

Following these steps will resolve most 404 issues!
