import { SPHttpClient } from '@microsoft/sp-http';

export interface CSVRow {
    [key: string]: string | number;
}

export interface CSVData {
    headers: string[];
    rows: CSVRow[];
}

/**
 * Service to fetch and parse CSV files from SharePoint library
 */
export class CSVDataService {
    private static context: any;
    // Simple in-memory cache keyed by server-relative path
    private static _cache: { [serverRelativePath: string]: CSVData } = {};

    /**
     * Initialize the service with SPFx context
     */
    public static initialize(context: any): void {
        CSVDataService.context = context;
    }

    /**
     * Get the REST API URL for a file in SharePoint
     */
    private static getFileRestUrl(libraryName: string, folderPath: string, fileName: string): string {
        const siteUrl = CSVDataService.context.pageContext.web.absoluteUrl;

        // Build server-relative path starting with the web's serverRelativeUrl
        const webServerRel: string = CSVDataService.context.pageContext.web.serverRelativeUrl || '';
        const parts = [webServerRel, libraryName];
        if (folderPath) {
            parts.push(folderPath);
        }
        parts.push(fileName);
        // Join and normalize slashes
        let serverRelativePath = parts.join('/').replace(/\\/g, '/').replace(/\/\//g, '/');

        // Ensure leading slash
        if (!serverRelativePath.startsWith('/')) {
            serverRelativePath = '/' + serverRelativePath;
        }

        // Normalize web root when it's just '/'
        if (webServerRel === '/') {
            // If webServerRel is '/', avoid double prefixing; serverRelativePath already starts with '/'
        }

        // Log computed path for debugging
        console.log('Computed file serverRelativePath:', serverRelativePath);

        // Encode the server relative path for use inside the decodedurl parameter
        const encoded = encodeURIComponent(serverRelativePath);

        // Return REST API endpoint for file content ($value returns the binary/text content)
        return `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encoded}')/$value`;
    }

    /**
     * Build a normalized server-relative path for the given library/folder/file
     */
    public static buildServerRelativePath(libraryName: string, folderPath: string, fileName: string): string {
        const webServerRel: string = CSVDataService.context.pageContext.web.serverRelativeUrl || '';
        const parts = [webServerRel, libraryName];
        if (folderPath) {
            parts.push(folderPath);
        }
        parts.push(fileName);
        let serverRelativePath = parts.join('/').replace(/\\/g, '/').replace(/\/\//g, '/');
        if (!serverRelativePath.startsWith('/')) {
            serverRelativePath = '/' + serverRelativePath;
        }
        return serverRelativePath;
    }

    /**
     * Cache helper: set cached CSV by serverRelativePath
     */
    public static setCachedCSV(serverRelativePath: string, data: CSVData): void {
        if (!serverRelativePath) return;
        CSVDataService._cache[serverRelativePath] = data;
    }

    /**
     * Cache helper: get cached CSV by serverRelativePath
     */
    public static getCachedCSV(serverRelativePath: string): CSVData | undefined {
        if (!serverRelativePath) return undefined;
        return CSVDataService._cache[serverRelativePath];
    }
    /**
     * Fetch CSV file from SharePoint library and parse it
     * @param libraryName - Name of the SharePoint library (e.g., 'Shared Documents')
     * @param folderPath - Path within the library (e.g., 'Reports' or 'Reports/2024')
     * @param fileName - Name of the CSV file (e.g., 'data.csv')
     * @returns Promise containing parsed CSV data with headers and rows
     */
    public static async fetchCSVFromSharePoint(
        libraryName: string,
        folderPath: string,
        fileName: string
    ): Promise<CSVData> {
        try {
            // Build server-relative path similar to getFileRestUrl logic so we can try multiple endpoint variants
            const siteUrl = CSVDataService.context.pageContext.web.absoluteUrl;
            const webServerRel: string = CSVDataService.context.pageContext.web.serverRelativeUrl || '';
            const parts = [webServerRel, libraryName];
            if (folderPath) {
                parts.push(folderPath);
            }
            parts.push(fileName);
            let serverRelativePath = parts.join('/').replace(/\\/g, '/').replace(/\/\//g, '/');
            if (!serverRelativePath.startsWith('/')) {
                serverRelativePath = '/' + serverRelativePath;
            }

            console.log('Computed file serverRelativePath for fetchCSVFromSharePoint:', serverRelativePath);

            const spHttpClient: SPHttpClient = CSVDataService.context.spHttpClient;

            const attempts: { key: string; url: string }[] = [];

            // Build multiple endpoint variants to maximize compatibility
            attempts.push({
                key: 'path-encoded',
                url: `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(serverRelativePath)}')/$value`
            });

            attempts.push({
                key: 'path-raw',
                url: `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativePath}')/$value`
            });

            attempts.push({
                key: 'alias-encoded',
                url: `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='${encodeURIComponent(serverRelativePath)}'`
            });

            attempts.push({
                key: 'alias-raw',
                url: `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='${serverRelativePath}'`
            });

            attempts.push({
                key: 'download-sourceurl',
                url: `${siteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativePath)}`
            });

            const errors: string[] = [];

            for (const attempt of attempts) {
                try {
                    console.log(`Attempting [${attempt.key}]: ${attempt.url}`);
                    const response = await spHttpClient.get(attempt.url, SPHttpClient.configurations.v1);
                    if (!response.ok) {
                        const debugText = await response.text().then(t => t.substring(0, 2000)).catch(() => '');
                        console.warn(`Attempt ${attempt.key} failed: HTTP ${response.status} ${response.statusText}`, debugText);
                        errors.push(`[${attempt.key}] HTTP ${response.status} ${response.statusText} - ${debugText}`);
                        continue;
                    }

                    const content = await response.text();
                    const parsed = this.parseCSV(content);
                    // cache result for this serverRelativePath to avoid duplicate network calls
                    try {
                        CSVDataService.setCachedCSV(serverRelativePath, parsed);
                    } catch (e) {
                        // ignore cache set failures
                    }
                    return parsed;
                } catch (err: any) {
                    console.error(`Attempt ${attempt.key} encountered an error`, err && err.message ? err.message : err);
                    errors.push(`[${attempt.key}] ${err && err.message ? err.message : String(err)}`);
                }
            }

            // If we reach here, all attempts failed
            const summary = errors.join('\n');
            console.error('All fetch attempts failed for fetchCSVFromSharePoint. Summary:', summary);
            throw new Error(`Failed to fetch CSV file from SharePoint. Attempts:\n${summary}`);
        } catch (error) {
            console.error('Error fetching CSV from SharePoint:', error);
            throw new Error(`Failed to fetch CSV file: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    }    /**
     * Parse CSV content string into structured data
     * @param csvContent - Raw CSV content as string
     * @returns Parsed CSV data with headers and rows
     */
    public static parseCSV(csvContent: string): CSVData {
        try {
            const lines = csvContent.trim().split('\n');

            if (lines.length === 0) {
                throw new Error('CSV file is empty');
            }

            // Parse headers (first line)
            const rawHeaders = this.parseCSVLine(lines[0]);
            // Define the canonical header sets we expect (base and optional ExportDate)
            const baseHeaders = ['Property', 'Title', 'Label', 'Value', 'TextData'];
            const expectedHeaders = baseHeaders.concat(['ExportDate']);

            // Create mapping from expected header -> index in rawHeaders (case-insensitive match)
            const indexMap: number[] = expectedHeaders.map(h => {
                const found = rawHeaders.findIndex(rh => rh.trim().toLowerCase() === h.toLowerCase());
                return found >= 0 ? found : -1;
            });

            // Determine whether the first line is a header row by checking for any header-name matches
            const anyMatched = indexMap.some(i => i >= 0);

            let headers: string[] = [];
            let dataStartIndex = 1; // default: data starts on line 1 (second line)

            if (!anyMatched) {
                // No header names matched; assume header-less CSV where columns are positional
                // Choose header list length to match the number of columns present on the first line
                if (rawHeaders.length === baseHeaders.length) {
                    headers = baseHeaders.slice();
                } else if (rawHeaders.length >= expectedHeaders.length) {
                    headers = expectedHeaders.slice();
                } else if (rawHeaders.length < baseHeaders.length && rawHeaders.length > 0) {
                    // Fewer columns than baseHeaders: map as many base headers as available
                    headers = baseHeaders.slice(0, rawHeaders.length);
                } else {
                    // Fallback to baseHeaders
                    headers = baseHeaders.slice();
                }

                // Positional mapping: the nth header maps to index n
                for (let i = 0; i < headers.length; i++) {
                    indexMap[i] = i;
                }

                // Include the first line as data (since it was not a header)
                dataStartIndex = 0;
                console.log('No header row detected; using positional mapping to', headers);
            } else {
                // At least one header matched: use expectedHeaders as the canonical header names
                headers = expectedHeaders.slice();
                console.log('CSV header mapping:', expectedHeaders.map((h, idx) => ({ expected: h, index: indexMap[idx], raw: indexMap[idx] >= 0 ? rawHeaders[indexMap[idx]] : null })));
            }

            // Parse data rows
            const rows: CSVRow[] = [];
            for (let i = dataStartIndex; i < lines.length; i++) {
                const values = this.parseCSVLine(lines[i]);

                // Skip empty lines
                if (values.length === 0 || (values.length === 1 && values[0] === '')) {
                    continue;
                }

                // Create row object following resolved header order
                const row: CSVRow = {};
                for (let j = 0; j < headers.length; j++) {
                    const srcIdx = indexMap[j];
                    const value = (srcIdx >= 0 && srcIdx < values.length ? values[srcIdx] : '') || '';
                    // Try to convert to number if possible
                    row[headers[j]] = isNaN(Number(value)) ? value : Number(value);
                }
                rows.push(row);
            }

            return { headers, rows };
        } catch (error) {
            console.error('Error parsing CSV:', error);
            throw new Error(`Failed to parse CSV: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    }

    /**
     * Parse a single CSV line handling quoted fields
     * @param line - CSV line to parse
     * @returns Array of parsed values
     */
    private static parseCSVLine(line: string): string[] {
        const result: string[] = [];
        let current = '';
        let insideQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];

            if (char === '"') {
                if (insideQuotes && nextChar === '"') {
                    // Escaped quote
                    current += '"';
                    i++; // Skip next quote
                } else {
                    // Toggle quote state
                    insideQuotes = !insideQuotes;
                }
            } else if (char === ',' && !insideQuotes) {
                // Field separator
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }

        // Add last field
        result.push(current.trim());

        return result;
    }

    /**
     * Get list of CSV files from a SharePoint library folder
     * @param libraryName - Name of the SharePoint library
     * @param folderPath - Path within the library
     * @returns Promise containing array of CSV file names
     */
    public static async listCSVFiles(
        libraryName: string,
        folderPath: string = ''
    ): Promise<string[]> {
        try {
            const siteUrl = CSVDataService.context.pageContext.web.absoluteUrl;

            // Build server-relative folder path starting with the web's serverRelativeUrl
            const webServerRel: string = CSVDataService.context.pageContext.web.serverRelativeUrl || '';
            const folderPath_ = folderPath ? `/${folderPath}` : '';
            let serverRelativePath = `${webServerRel}/${libraryName}${folderPath_}`;
            // Normalize slashes
            serverRelativePath = serverRelativePath.replace(/\\/g, '/').replace(/\/\//g, '/');

            // Ensure leading slash
            if (!serverRelativePath.startsWith('/')) {
                serverRelativePath = '/' + serverRelativePath;
            }

            // Log computed server-relative path for debugging
            console.log('Computed folder serverRelativePath:', serverRelativePath);

            // Encode for REST call
            const encodedPath = encodeURIComponent(serverRelativePath);

            // Get folder items using REST API
            const restUrl = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedPath}')/Files?$select=Name`;

            console.log('Listing files from:', restUrl);

            const spHttpClient: SPHttpClient = CSVDataService.context.spHttpClient;
            const response = await spHttpClient.get(restUrl, SPHttpClient.configurations.v1);

            if (!response.ok) {
                throw new Error(`Failed to list files: HTTP ${response.status} - ${response.statusText}`);
            }

            const data = await response.json();

            // Filter for CSV files
            const csvFiles = (data.value || [])
                .filter((item: any) => item.Name && item.Name.toLowerCase().endsWith('.csv'))
                .map((item: any) => item.Name);

            console.log('Found CSV files:', csvFiles);
            return csvFiles;
        } catch (error) {
            console.error('Error listing CSV files:', error);
            throw new Error(`Failed to list CSV files: ${error instanceof Error ? error.message : 'Unknown error'}`);
        }
    }

    /**
     * Extract server-relative path from a SharePoint sharing link
     * Examples of sharing links contain segments like '/:x:/r/sites/TheLoop/Shared%20Documents/...'
     */
    private static extractServerRelativePathFromLink(link: string): string {
        try {
            const u = new URL(link);
            const rawPath = u.pathname || '';

            // Look for typical site prefixes
            const sitesIdx = rawPath.indexOf('/sites/');
            const teamsIdx = rawPath.indexOf('/teams/');

            let startIdx = -1;
            if (sitesIdx !== -1) {
                startIdx = sitesIdx;
            } else if (teamsIdx !== -1) {
                startIdx = teamsIdx;
            } else {
                // Fallback: look for Shared Documents
                const sdIdx = rawPath.indexOf('/Shared%20Documents');
                if (sdIdx !== -1) {
                    startIdx = sdIdx;
                }
            }

            let serverRelative = startIdx !== -1 ? rawPath.substring(startIdx) : rawPath;

            // decode URL encoded characters
            serverRelative = decodeURIComponent(serverRelative);

            // Ensure leading slash
            if (!serverRelative.startsWith('/')) {
                serverRelative = '/' + serverRelative;
            }

            // Ensure the server-relative path starts with the current web's serverRelativeUrl
            try {
                const webServerRel: string = CSVDataService.context?.pageContext?.web?.serverRelativeUrl || '';
                if (webServerRel && webServerRel !== '/' && !serverRelative.startsWith(webServerRel)) {
                    // If serverRelative already contains the web path later in the string, avoid duplicating
                    if (serverRelative.indexOf(webServerRel) === -1) {
                        // Prefix the web server relative URL
                        serverRelative = (webServerRel.endsWith('/') ? webServerRel.slice(0, -1) : webServerRel) + serverRelative;
                    }
                }
            } catch (e) {
                // If context isn't available for some reason, just proceed with the decoded path
            }

            return serverRelative;
        } catch (e) {
            console.warn('Failed to parse link, using raw link path as fallback', e);
            return link;
        }
    }

    /**
     * Fetch CSV by a full SharePoint sharing link (handles the /:x:/r/ style sharing links)
     */
    public static async fetchCSVFromLink(link: string): Promise<CSVData> {
        const siteUrl = CSVDataService.context.pageContext.web.absoluteUrl;
        const serverRelativePath = CSVDataService.extractServerRelativePathFromLink(link);

        const spHttpClient: SPHttpClient = CSVDataService.context.spHttpClient;

        // Try multiple endpoint/encoding variants to handle different link formats
        const attempts: { key: string; url: string }[] = [];

        // Alias endpoint (encoded value)
        attempts.push({
            key: 'alias-encoded',
            url: `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='${encodeURIComponent(serverRelativePath)}'`
        });

        // Alias endpoint (raw value)
        attempts.push({
            key: 'alias-raw',
            url: `${siteUrl}/_api/web/GetFileByServerRelativeUrl(@v)/$value?@v='${serverRelativePath}'`
        });

        // Path endpoint (encoded)
        attempts.push({
            key: 'path-encoded',
            url: `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodeURIComponent(serverRelativePath)}')/$value`
        });

        // Path endpoint (raw)
        attempts.push({
            key: 'path-raw',
            url: `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${serverRelativePath}')/$value`
        });

        // As a last resort, try using the /_layouts/15/download.aspx?SourceUrl= approach
        attempts.push({
            key: 'download-sourceurl',
            url: `${siteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativePath)}`
        });

        const errors: string[] = [];

        for (const attempt of attempts) {
            try {
                console.log(`Attempting [${attempt.key}]: ${attempt.url}`);
                const response = await spHttpClient.get(attempt.url, SPHttpClient.configurations.v1);
                const status = response.status;
                const statusText = response.statusText;

                // Capture small debug body for non-ok responses
                if (!response.ok) {
                    const debugText = await response.text().then(t => t.substring(0, 2000)).catch(() => '');
                    console.warn(`Attempt ${attempt.key} failed: HTTP ${status} ${statusText}`, debugText);
                    errors.push(`[${attempt.key}] HTTP ${status} ${statusText} - ${debugText}`);
                    // continue to next attempt
                    continue;
                }

                // Success: parse content
                const content = await response.text();
                try {
                    return this.parseCSV(content);
                } catch (parseErr) {
                    console.error('CSV parse failed for successful HTTP response', parseErr);
                    throw parseErr;
                }
            } catch (err: any) {
                console.error(`Attempt ${attempt.key} encountered an error`, err && err.message ? err.message : err);
                errors.push(`[${attempt.key}] ${err && err.message ? err.message : String(err)}`);
            }
        }

        // If we reach here, all attempts failed
        const summary = errors.join('\n');
        console.error('All fetch attempts failed for link. Summary:', summary);
        throw new Error(`Failed to fetch CSV from link. Attempts:\n${summary}`);
    }
}
