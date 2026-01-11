import * as React from 'react';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Dropdown, IDropdownOption, TextField, PrimaryButton, DefaultButton, Stack, StackItem, Pivot, PivotItem } from '@fluentui/react';
import styles from './CSVReportViewer.module.scss';
import { CSVDataService, CSVData, CSVRow } from '../services/CSVDataService';

// Dynamic import for Chart.js components
const ChartComponent = React.lazy(() => import('./ChartComponent'));

export interface ICSVReportViewerProps {
    libraryName: string;
    folderPath?: string;
    fileName?: string;
    isDarkTheme?: boolean;
    // from web part: per-chart visibility map (title -> boolean)
    chartVisibilities?: { [title: string]: boolean };
    // from web part: per-chart hide axis names map (title -> boolean)
    hideAxisNames?: { [title: string]: boolean };
    // from web part: per-chart labels (sanitized key -> label)
    chartLabels?: { [title: string]: string };
    // optional SPFx context passed from web part
    context?: any;
}

interface ICSVReportViewerState {
    data: CSVData | null;
    loading: boolean;
    error: string | null;
    selectedFile: string | null;
    availableFiles: string[];
    libraryName: string;
    folderPath: string;
    fileName: string;
    selectedProperty: string | null;
    perTitleChartTypes: { [title: string]: 'bar' | 'line' | 'pie' | 'doughnut' };
    perTitleVisibility: { [title: string]: boolean };
    chartType: 'bar' | 'line' | 'pie' | 'doughnut';
}

export default class CSVReportViewer extends React.Component<ICSVReportViewerProps, ICSVReportViewerState> {
    private sanitizeTitleKey(title: string): string {
        return String(title).replace(/[^a-zA-Z0-9]/g, '_');
    }

    private getLabelForTitle(title: string): string {
        const key = this.sanitizeTitleKey(title);
        if (this.props.chartLabels && typeof this.props.chartLabels[key] === 'string' && this.props.chartLabels[key].trim().length > 0) {
            return this.props.chartLabels[key];
        }
        return title;
    }

    private getPropertyDetails(prop: string): Array<{ label: string; value: string }> {
        const result: Array<{ label: string; value: string }> = [];
        const d = this.state.data;
        if (!d || !d.rows || !d.headers) return result;

        // Normalize requested property
        const requestedProp = String(prop || '').trim();

        const detailsRows = d.rows.filter(r => {
            const title = String(r['Title'] || r['title'] || '').trim().toLowerCase();
            const property = String(r['Property'] || r['property'] || '').trim();
            // Consider header-less or varied casing: check includes 'property' and either 'detail' or 'data'
            const isDetail = title === 'property details'
                || title === 'property detail'
                || title === 'property data'
                || (title.indexOf('property') >= 0 && (title.indexOf('detail') >= 0 || title.indexOf('data') >= 0));
            return isDetail && property === requestedProp;
        });

        // Preserve the order of rows as they appear in the CSV for display.
        detailsRows.forEach(r => {
            const rawLabel = String(r['Label'] || r['label'] || r['LabelName'] || r['labelName'] || '').trim();
            const normalizedLabel = rawLabel || String(r['Label'] || r['label'] || '').trim();
            const textData = String(r['TextData'] || r['Text'] || r['text'] || '').trim();
            const valueData = String(r['Value'] || r['ValueData'] || r['Number'] || r['value'] || '').trim();

            const value = textData || valueData || '';
            const labelForDisplay = normalizedLabel ? normalizedLabel.replace(/\s+/g, ' ').trim() : '';
            if (labelForDisplay && value) {
                result.push({ label: this.toTitleCase(labelForDisplay), value });
            } else if (value) {
                // Fallback label when none provided — use 'Value'
                result.push({ label: 'Value', value });
            }
        });

        return result;
    }

    private toTitleCase(str: string): string {
        return str.split(' ').map(s => s.charAt(0).toUpperCase() + s.slice(1)).join(' ');
    }

    private pickPropertyFields(details: Array<{ label: string; value: string }>, wanted: string[]): Array<{ label: string; value: string }> {
        if (!details || details.length === 0) return [];
        const map = new Map<string, { label: string; value: string }>();
        details.forEach(d => {
            map.set(d.label.toLowerCase(), d);
        });
        const result: Array<{ label: string; value: string }> = [];
        wanted.forEach(w => {
            const found = map.get(w.toLowerCase());
            if (found) result.push(found);
        });
        return result;
    }

    private getSimplePropertyDetails(prop: string): { name?: string; location?: string; unitCount?: string } {
        const details = this.getPropertyDetails(prop);
        const result: { name?: string; location?: string; unitCount?: string } = {};
        if (!details || details.length === 0) return result;
        details.forEach(d => {
            const label = (d.label || '').toLowerCase();
            const value = d.value || '';
            if (!value) return;
            if (label === 'name' || label.indexOf('name') >= 0) {
                if (!result.name) result.name = value;
            } else if (label === 'location' || label.indexOf('location') >= 0) {
                if (!result.location) result.location = value;
            } else if (label === 'unit count' || label === 'unitcount' || label.indexOf('unit') >= 0) {
                if (!result.unitCount) result.unitCount = value;
            }
        });
        return result;
    }
    constructor(props: ICSVReportViewerProps) {
        super(props);

        this.state = {
            data: null,
            loading: false,
            error: null,
            selectedFile: props.fileName || null,
            availableFiles: [],
            libraryName: props.libraryName,
            folderPath: props.folderPath || '',
            fileName: props.fileName || '',

            selectedProperty: null,
            perTitleChartTypes: {},
            perTitleVisibility: {},
            chartType: 'bar'
        };
    }

    public componentDidMount(): void {
        // Load available files on mount
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.loadAvailableFiles();
    }

    private loadAvailableFiles = async (): Promise<void> => {
        try {
            this.setState({ loading: true, error: null });
            const files = await CSVDataService.listCSVFiles(
                this.state.libraryName,
                this.state.folderPath
            );
            this.setState({ availableFiles: files, loading: false });

            // Auto-select first file if none selected and files exist
            if ((!this.state.selectedFile || this.state.selectedFile === null) && files && files.length > 0) {
                const first = files[0];
                // eslint-disable-next-line @typescript-eslint/no-floating-promises
                this.loadCSVFile(first);
                return;
            }

            // If a fileName was provided, load it automatically (fallback)
            if (this.state.fileName) {
                // eslint-disable-next-line @typescript-eslint/no-floating-promises
                this.loadCSVFile(this.state.fileName);
            }
        } catch (error) {
            this.setState({
                error: error instanceof Error ? error.message : 'Failed to load files',
                loading: false
            });
        }
    };

    private loadCSVFile = async (fileName: string): Promise<void> => {
        try {
            this.setState({ loading: true, error: null, selectedFile: fileName });
            const csvData = await CSVDataService.fetchCSVFromSharePoint(
                this.state.libraryName,
                this.state.folderPath,
                fileName
            );

            // Cache the fetched CSV so the web part property pane can use it without another network call
            try {
                const serverRel = CSVDataService.buildServerRelativePath(this.state.libraryName, this.state.folderPath, fileName);
                CSVDataService.setCachedCSV(serverRel, csvData);
            } catch (e) {
                // ignore cache set failures
            }

            // Auto-select first two columns for initial chart
            const xAxis = csvData.headers[0] || null;
            const yAxis = csvData.headers[1] || null;

            this.setState({
                data: csvData,
                loading: false,

            });

            // Auto-select first property when data is loaded and none selected
            try {
                const props = Array.from(new Set(csvData.rows.map(r => String(r['Property'] || '').trim()))).filter(p => p);
                if ((!this.state.selectedProperty || this.state.selectedProperty === null) && props.length > 0) {
                    this.setState({ selectedProperty: props[0] });
                }
            } catch (e) {
                // ignore
            }

            // initialize per-title visibility from webpart props if provided
            try {
                const titleHeader = (csvData.headers && csvData.headers.length > 1) ? csvData.headers[1] : 'Title';
                const titles = Array.from(new Set(csvData.rows.map(r => String(r[titleHeader] || '').trim()))).filter(t => t);
                const visMap: { [title: string]: boolean } = {};
                titles.forEach(t => {
                    if (this.props.chartVisibilities && typeof this.props.chartVisibilities[t] !== 'undefined') {
                        visMap[t] = !!this.props.chartVisibilities[t];
                    } else {
                        visMap[t] = true;
                    }
                });
                this.setState({ perTitleVisibility: visMap });
            } catch (e) {
                // ignore
            }
        } catch (error) {
            this.setState({
                error: error instanceof Error ? error.message : 'Failed to load CSV file',
                loading: false,
                data: null
            });
        }
    };

    private handleFileChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option && typeof option.key === 'string') {
            // eslint-disable-next-line @typescript-eslint/no-floating-promises
            this.loadCSVFile(option.key);
        }
    };

    private handleLibraryChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ libraryName: newValue || '' });
    };

    private handleFolderChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({ folderPath: newValue || '' });
    };



    private handlePropertyChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            this.setState({ selectedProperty: option.key as string });
        }
    };

    private handleChartTypeChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (option) {
            this.setState({ chartType: option.key as 'bar' | 'line' | 'pie' | 'doughnut' });
        }
    };

    private handlePerTitleChartTypeChange = (title: string, _event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
        if (!option) return;
        this.setState(prev => ({
            perTitleChartTypes: { ...prev.perTitleChartTypes, [title]: option.key as 'bar' | 'line' | 'pie' | 'doughnut' }
        }));
    };

    private handlePerTitleVisibilityChange = (title: string, checked?: boolean): void => {
        this.setState(prev => ({
            perTitleVisibility: { ...prev.perTitleVisibility, [title]: !!checked }
        }));
    };

    private handlePrintAll = (): void => {
        try {
            const container = document.querySelector(`.${styles.chartContainer}`) as HTMLElement | null;
            if (!container) {
                this.setState({ error: 'No charts available to print' });
                return;
            }

            // Select chart wrappers that we annotated with data-title and respect visibility settings
            const wrappers = Array.from(container.querySelectorAll('[data-title]')) as HTMLElement[];
            const visibleWrappers = wrappers.filter(w => {
                const title = w.getAttribute('data-title') || '';
                if (this.props.chartVisibilities && this.props.chartVisibilities[title] === false) return false;
                if (this.state.perTitleVisibility && this.state.perTitleVisibility[title] === false) return false;
                return true;
            });

            if (visibleWrappers.length === 0) {
                this.setState({ error: 'No visible chart canvases found to print' });
                return;
            }

            const parts: string[] = [];
            visibleWrappers.forEach((wrapper) => {
                const c = wrapper.querySelector('canvas') as HTMLCanvasElement | null;
                if (!c) return;
                const rawTitle = wrapper.getAttribute('data-title') || '';
                const title = this.getLabelForTitle(rawTitle);
                const dataUrl = c.toDataURL('image/png');
                parts.push(`<div style="page-break-inside:avoid;margin-bottom:24px;"><h2 style='font-family:Arial,Helvetica,sans-serif;'>${title || ''}</h2><img src='${dataUrl}' style='max-width:100%;height:auto;' /></div>`);
            });

            // If a property is selected, include Property Data block before charts
            let headerHtml = '';
            if (this.state.selectedProperty) {
                const pd = this.getSimplePropertyDetails(this.state.selectedProperty as string);
                const lines: string[] = [];
                if (pd.name) lines.push(`<div><strong>Name:</strong> ${pd.name}</div>`);
                if (pd.location) lines.push(`<div><strong>Location:</strong> ${pd.location}</div>`);
                if (pd.unitCount) lines.push(`<div><strong>Unit Count:</strong> ${pd.unitCount}</div>`);
                if (lines.length > 0) {
                    headerHtml = `<div style="margin-bottom:16px;"><h3 style='font-family:Arial,Helvetica,sans-serif;margin:0 0 8px 0;'>Property Data</h3>${lines.join('')}</div>`;
                }
            }

            const html = `<!doctype html><html><head><meta charset='utf-8'><title>Charts</title><style>body{font-family:Arial,Helvetica,sans-serif;padding:16px} img{display:block;margin:8px 0} @media print { img{max-width:100%} }</style></head><body>${headerHtml}${parts.join('')}</body></html>`;

            const printWindow = window.open('', '_blank');
            if (!printWindow) {
                this.setState({ error: 'Unable to open print window (popup blocked?)' });
                return;
            }
            printWindow.document.open();
            printWindow.document.write(html);
            printWindow.document.close();
            printWindow.focus();
            // Delay slightly to allow images to load
            setTimeout(() => {
                try {
                    printWindow.print();
                } catch (e) {
                    console.warn('Print failed', e);
                }
            }, 600);
        } catch (e) {
            console.error('Print all charts error', e);
            this.setState({ error: e instanceof Error ? e.message : 'Print failed' });
        }
    };

    private handleDownloadPDF = async (): Promise<void> => {
        try {
            const container = document.querySelector(`.${styles.chartContainer}`) as HTMLElement | null;
            if (!container) {
                this.setState({ error: 'No charts available to export' });
                return;
            }

            // Select only visible chart wrappers and find their canvases
            const wrappers = Array.from(container.querySelectorAll('[data-title]')) as HTMLElement[];
            const visibleWrappers = wrappers.filter(w => {
                const title = w.getAttribute('data-title') || '';
                if (this.props.chartVisibilities && this.props.chartVisibilities[title] === false) return false;
                if (this.state.perTitleVisibility && this.state.perTitleVisibility[title] === false) return false;
                return true;
            });

            const items: Array<{ canvas: HTMLCanvasElement; title: string }> = visibleWrappers.map(w => {
                const c = w.querySelector('canvas') as HTMLCanvasElement | null;
                const rawTitle = w.getAttribute('data-title') || '';
                const title = this.getLabelForTitle(rawTitle);
                return c ? { canvas: c, title } : null;
            }).filter(Boolean) as Array<{ canvas: HTMLCanvasElement; title: string }>;

            if (items.length === 0) {
                this.setState({ error: 'No chart canvases found to export' });
                return;
            }

            // Prefer dynamic import of the local `jspdf` package (clean bundling).
            // If dynamic import fails for any reason, fall back to any UMD global already present on the page.
            let jsPDFModule: any = null;
            try {
                jsPDFModule = await import('jspdf');
            } catch (e) {
                // dynamic import may fail in some environments; we'll fall back to globals
                console.warn('Dynamic import of jspdf failed, falling back to global (if available).', e);
            }

            const jsPDFLib: any = jsPDFModule?.jsPDF || jsPDFModule || (window as any).jspdf?.jsPDF || (window as any).jsPDF || (window as any).jspdf;
            if (!jsPDFLib) {
                console.error('jspdf not available (local package or global)');
                this.setState({ error: 'Required library "jspdf" is not available. Run "npm install jspdf" and rebuild.' });
                return;
            }

            const jsPDFConstructor = jsPDFLib.jsPDF ? jsPDFLib.jsPDF : jsPDFLib;
            const pdf = new jsPDFConstructor({ unit: 'mm', format: 'a4', orientation: 'portrait' });
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = pdf.internal.pageSize.getHeight();

            // Build heading using ExportDate month and year (fall back to current month/year)
            let monthName = '';
            let yearNum: number | null = null;
            try {
                const csv = this.state.data;
                if (csv && csv.rows && csv.rows.length > 0) {
                    const headerLower = csv.headers.map(h => h.toLowerCase());
                    const idx = headerLower.findIndex(h => h.indexOf('export') >= 0 || h.indexOf('exportdate') >= 0 || h.indexOf('export_date') >= 0);
                    const val = idx >= 0 ? csv.rows[0][csv.headers[idx]] : csv.rows[0]['ExportDate'] || csv.rows[0]['Export Date'];
                    const d = val ? new Date(String(val)) : new Date();
                    if (!isNaN(d.getTime())) {
                        monthName = d.toLocaleString('en-US', { month: 'long' });
                        yearNum = d.getFullYear();
                    }
                }
            } catch (e) {
                // ignore and fallback
            }
            if (!monthName) {
                const now = new Date();
                monthName = now.toLocaleString('en-US', { month: 'long' });
                yearNum = now.getFullYear();
            } else if (!yearNum) {
                yearNum = new Date().getFullYear();
            }

            const headingText = `LCOR demographics report (${monthName} - ${yearNum})`;

            // Layout charts in two columns per row to fit more charts on a single page.
            const margin = 10; // mm
            const gap = 8; // mm between columns/rows
            const colWidth = (pageWidth - margin * 2 - gap) / 2;

            // Draw heading on first page
            pdf.setFontSize(16);
            pdf.text(headingText, pageWidth / 2, 14, { align: 'center' });

            let currentY = 14 + 6; // start below heading

            // If a property is selected, render the Property Data block before charts
            if (this.state.selectedProperty) {
                const pd = this.getSimplePropertyDetails(this.state.selectedProperty as string);
                if (pd && (pd.name || pd.location || pd.unitCount)) {
                    // Title
                    pdf.setFontSize(12);
                    pdf.text('Property Data', pageWidth / 2, currentY, { align: 'center' });
                    currentY += 6;

                    // Lines left-aligned
                    pdf.setFontSize(10);
                    const lines = [];
                    if (pd.name) lines.push(`Name: ${pd.name}`);
                    if (pd.location) lines.push(`Location: ${pd.location}`);
                    if (pd.unitCount) lines.push(`Unit Count: ${pd.unitCount}`);
                    lines.forEach(line => {
                        if (currentY + 6 + margin > pageHeight) {
                            pdf.addPage();
                            pdf.setFontSize(16);
                            pdf.text(headingText, pageWidth / 2, 14, { align: 'center' });
                            currentY = 14 + 6;
                        }
                        pdf.text(line, margin, currentY);
                        currentY += 6;
                    });

                    // Add a small gap before charts
                    currentY += 4;
                }
            }

            for (let i = 0; i < items.length; i += 2) {
                const item1 = items[i];
                const c1 = item1.canvas;
                const imgData1 = c1.toDataURL('image/png');

                const img1 = new Image();
                await new Promise<void>((resolve) => {
                    img1.onload = () => resolve();
                    img1.onerror = () => resolve();
                    img1.src = imgData1;
                });

                const ratio1 = img1.width && img1.height ? img1.width / img1.height : 1;
                const h1 = colWidth / ratio1;

                const title1 = item1.title || '';

                let img2: HTMLImageElement | null = null;
                let h2 = 0;
                let imgData2: string | null = null;

                if (i + 1 < items.length) {
                    const item2 = items[i + 1];
                    const c2 = item2.canvas;
                    imgData2 = c2.toDataURL('image/png');
                    img2 = new Image();
                    await new Promise<void>((resolve) => {
                        img2!.onload = () => resolve();
                        img2!.onerror = () => resolve();
                        img2!.src = imgData2!;
                    });
                    const ratio2 = img2.width && img2.height ? img2.width / img2.height : 1;
                    h2 = colWidth / ratio2;
                }

                const titleHeight = 6; // mm for title text
                const rowHeight = Math.max(h1 + titleHeight, h2 ? (h2 + titleHeight) : 0);

                // If this row doesn't fit, create a new page
                if (currentY + rowHeight + margin > pageHeight) {
                    pdf.addPage();
                    // redraw heading on new page
                    pdf.setFontSize(16);
                    pdf.text(headingText, pageWidth / 2, 14, { align: 'center' });
                    currentY = 14 + 6;
                }

                // Draw first image (left column) with title
                const x1 = margin;
                const y1 = currentY;
                pdf.setFontSize(12);
                pdf.text(String(title1), x1, y1 + 4);
                pdf.addImage(imgData1, 'PNG', x1, y1 + 6, colWidth, h1);

                // Draw second image (right column) if present
                if (imgData2 && img2) {
                    const title2 = items[i + 1].title || '';
                    const x2 = margin + colWidth + gap;
                    const y2 = currentY;
                    pdf.setFontSize(12);
                    pdf.text(String(title2), x2, y2 + 4);
                    pdf.addImage(imgData2, 'PNG', x2, y2 + 6, colWidth, h2);
                }

                currentY += rowHeight + gap;
            }

            pdf.save('charts.pdf');
        } catch (e: any) {
            console.error('Error creating PDF', e);
            this.setState({ error: e && e.message ? e.message : 'Failed to create PDF' });
        }
    };

    private handleRefresh = (): void => {
        // eslint-disable-next-line @typescript-eslint/no-floating-promises
        this.loadAvailableFiles();
    };

    private handleLoadFiles = async (): Promise<void> => {
        const { libraryName, folderPath } = this.state;
        if (!libraryName) {
            this.setState({ error: 'Please enter a library name' });
            return;
        }
        this.setState({ libraryName, folderPath });
        await this.loadAvailableFiles();
    };

    public render(): React.ReactElement<ICSVReportViewerProps> {
        const {
            data,
            loading,
            error,
            availableFiles,
            selectedFile,
            libraryName,
            folderPath,
            chartType
        } = this.state;

        const fileOptions: IDropdownOption[] = availableFiles.map(file => ({
            key: file,
            text: file
        }));

        const propertyOptions: IDropdownOption[] = data
            ? Array.from(new Set(data.rows.map(r => String(r['Property'] || '').trim()))).filter(p => p).map(p => ({ key: p, text: p }))
            : [];

        const chartTypeOptions: IDropdownOption[] = [
            { key: 'bar', text: 'Bar Chart' },
            { key: 'line', text: 'Line Chart' },
            { key: 'pie', text: 'Pie Chart' },
            { key: 'doughnut', text: 'Doughnut Chart' }
        ];

        return (
            <div className={styles.csvReportViewer}>
                <h2>CSV Report Viewer</h2>

                {/* Error Message */}
                {error && (
                    <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: null })}>
                        {error}
                    </MessageBar>
                )}

                {/* Loading Spinner */}
                {loading && (
                    <div className={styles.spinnerContainer}>
                        <Spinner size={SpinnerSize.large} label="Loading..." />
                    </div>
                )}

                {/* Tabs */}
                <Pivot aria-label="Report tabs">
                    <PivotItem headerText="Select & Charts" itemKey="select">
                        {/* Top row: File, Property, Chart Type */}
                        <div className={styles.selectionSection}>
                            <Stack horizontal tokens={{ childrenGap: 12 }} wrap verticalAlign="end">
                                <StackItem grow styles={{ root: { minWidth: 220 } }}>
                                    <Dropdown
                                        label="Select CSV File"
                                        options={fileOptions}
                                        selectedKey={selectedFile}
                                        onChange={this.handleFileChange}
                                        placeholder="Choose a CSV file"
                                    />
                                </StackItem>

                                <StackItem grow styles={{ root: { minWidth: 200 } }}>
                                    <Dropdown
                                        label="Property"
                                        options={propertyOptions}
                                        selectedKey={this.state.selectedProperty}
                                        onChange={this.handlePropertyChange}
                                        placeholder="Select property to group by"
                                    />
                                </StackItem>

                                <StackItem styles={{ root: { minWidth: 160 } }}>
                                    <Dropdown
                                        label="Chart Type"
                                        options={chartTypeOptions}
                                        selectedKey={chartType}
                                        onChange={this.handleChartTypeChange}
                                    />
                                </StackItem>
                            </Stack>
                        </div>

                        {/* Chart Configuration header + actions */}
                        {data && data.rows.length > 0 && (
                            <div className={styles.chartSection}>
                                <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" styles={{ root: { justifyContent: 'space-between' } }}>
                                    <StackItem>
                                        <h3 style={{ margin: 0 }}>Chart Configuration</h3>
                                    </StackItem>
                                    <StackItem>
                                        <Stack horizontal tokens={{ childrenGap: 8 }}>
                                            <PrimaryButton
                                                text="Download PDF"
                                                onClick={() => void this.handleDownloadPDF()}
                                                disabled={loading || !data}
                                            />
                                            <DefaultButton
                                                text="Print All Charts"
                                                onClick={() => this.handlePrintAll()}
                                                disabled={loading || !data}
                                            />
                                        </Stack>
                                    </StackItem>
                                </Stack>

                                {/* Property Data: card-style summary */}
                                {this.state.selectedProperty && (
                                    (() => {
                                        const pd = this.getSimplePropertyDetails(this.state.selectedProperty as string);
                                        const hasAny = !!(pd.name || pd.location || pd.unitCount);
                                        if (!hasAny) return null;
                                        return (
                                            <div className={styles.propertyCard}>
                                                <div className={styles.propertyCardHeader}>Property Data</div>
                                                <div className={styles.propertyCardBody}>
                                                    {pd.name && <div><strong>Name:</strong> {pd.name}</div>}
                                                    {pd.location && <div><strong>Location:</strong> {pd.location}</div>}
                                                    {pd.unitCount && <div><strong>Unit Count:</strong> {pd.unitCount}</div>}
                                                </div>
                                            </div>
                                        );
                                    })()
                                )}

                                <div className={styles.chartContainer}>
                                    {/* existing chart rendering logic copied here */}
                                    {this.state.selectedProperty ? (
                                        (() => {
                                            const prop = this.state.selectedProperty as string;
                                            const titles = Array.from(new Set(data.rows.filter(r => String(r['Property']) === prop).map(r => String(r['Title'] || '').trim()))).filter(t => t);

                                            if (titles.length === 0) {
                                                return <MessageBar messageBarType={MessageBarType.warning}>No Titles found for property '{prop}'.</MessageBar>;
                                            }

                                            return (
                                                <div className={(styles as any).multipleCharts}>
                                                    {titles.map((title, idx) => {
                                                        const key = this.sanitizeTitleKey(title);
                                                        if (this.props.chartVisibilities && this.props.chartVisibilities[key] === false) {
                                                            return null;
                                                        }
                                                        if (this.state.perTitleVisibility && this.state.perTitleVisibility[title] === false) {
                                                            return null;
                                                        }
                                                        const filteredData: CSVData = {
                                                            headers: data.headers,
                                                            rows: data.rows.filter(r => String(r['Property']) === prop && String(r['Title'] || '').trim() === title)
                                                        };
                                                        const headerLower = data.headers.map(h => h.toLowerCase());
                                                        const hasLabel = headerLower.indexOf('label') >= 0;
                                                        const hasValue = headerLower.indexOf('value') >= 0;

                                                        const xAxisToUse = hasLabel
                                                            ? data.headers[headerLower.indexOf('label')]
                                                            : (data.headers[0] || '');

                                                        const yAxisToUse = hasValue
                                                            ? data.headers[headerLower.indexOf('value')]
                                                            : (data.headers[1] || data.headers[0] || '');

                                                        const perTitleType = this.state.perTitleChartTypes[title] || chartType;

                                                        // Use sanitized key to look up label overrides in web part properties
                                                        const labelFromProps = this.props.chartLabels && typeof this.props.chartLabels[key] === 'string' && this.props.chartLabels[key].trim().length > 0
                                                            ? this.props.chartLabels[key]
                                                            : title;

                                                        return (
                                                            <div key={idx} className={(styles as any).singleChartWrapper} data-title={title}>
                                                                <div className={(styles as any).singleChartHeader}>
                                                                    <div className={(styles as any).singleChartTitle}>{labelFromProps}</div>
                                                                    <div className={(styles as any).singleChartControls}>
                                                                        <Dropdown
                                                                            options={chartTypeOptions}
                                                                            selectedKey={perTitleType}
                                                                            onChange={(e, opt) => this.handlePerTitleChartTypeChange(title, e, opt)}
                                                                            styles={{ root: { width: 140 } }}
                                                                        />
                                                                    </div>
                                                                </div>
                                                                <React.Suspense fallback={<Spinner size={SpinnerSize.small} label="Loading..." />}>
                                                                    <ChartComponent
                                                                        data={filteredData}
                                                                        xAxis={xAxisToUse}
                                                                        yAxis={yAxisToUse}
                                                                        chartType={perTitleType}
                                                                        chartTitle={labelFromProps}
                                                                        isDarkTheme={this.props.isDarkTheme}
                                                                        hideAxisNames={this.props.hideAxisNames ? !!this.props.hideAxisNames[key] : false}
                                                                    />
                                                                </React.Suspense>
                                                            </div>
                                                        );
                                                    })}
                                                </div>
                                            );
                                        })()
                                    ) : (
                                        <MessageBar messageBarType={MessageBarType.info}>Select a Property to view per-Title charts.</MessageBar>
                                    )}
                                </div>
                            </div>
                        )}
                    </PivotItem>

                    <PivotItem headerText="Configuration" itemKey="configuration">
                        <div className={styles.configSection}>
                            <h3>Configuration</h3>
                            <Stack tokens={{ childrenGap: 15 }}>
                                <StackItem>
                                    <TextField
                                        label="SharePoint Library Name"
                                        value={libraryName}
                                        onChange={this.handleLibraryChange}
                                        placeholder="e.g., Shared Documents"
                                        disabled={loading}
                                    />
                                </StackItem>

                                <StackItem>
                                    <TextField
                                        label="Folder Path (optional)"
                                        value={folderPath}
                                        onChange={this.handleFolderChange}
                                        placeholder="e.g., Reports/2024"
                                        disabled={loading}
                                    />
                                </StackItem>

                                <StackItem>
                                    <Stack horizontal tokens={{ childrenGap: 10 }}>
                                        <StackItem grow>
                                            <PrimaryButton
                                                text="Load Files"
                                                onClick={this.handleLoadFiles}
                                                disabled={loading || !libraryName}
                                            />
                                        </StackItem>
                                        <StackItem>
                                            <DefaultButton
                                                text="Refresh"
                                                onClick={this.handleRefresh}
                                                disabled={loading || availableFiles.length === 0}
                                            />
                                        </StackItem>
                                    </Stack>
                                </StackItem>
                            </Stack>
                        </div>
                    </PivotItem>

                    <PivotItem headerText="Data Preview" itemKey="preview">
                        {/* Data Table Preview */}
                        {data ? (
                            <div className={styles.dataTableSection}>
                                <h3>Data Preview ({data.rows.length} rows)</h3>
                                <div className={styles.tableWrapper}>
                                    <table className={styles.dataTable}>
                                        <thead>
                                            <tr>
                                                {data.headers.map((header, idx) => (
                                                    <th key={idx}>{header}</th>
                                                ))}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {data.rows.map((row, idx) => (
                                                <tr key={idx}>
                                                    {data.headers.map((header, colIdx) => (
                                                        <td key={colIdx}>{row[header]}</td>
                                                    ))}
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        ) : (
                            <MessageBar messageBarType={MessageBarType.info}>No data loaded yet.</MessageBar>
                        )}
                    </PivotItem>
                </Pivot>

                {/* Error Message */}
                {error && (
                    <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: null })}>
                        {error}
                    </MessageBar>
                )}

                {/* Loading Spinner */}
                {loading && (
                    <div className={styles.spinnerContainer}>
                        <Spinner size={SpinnerSize.large} label="Loading..." />
                    </div>
                )}

                {/* File Selection */}


                {/* Chart Configuration and Display */}


                {!loading && data && data.rows.length === 0 && (
                    <MessageBar messageBarType={MessageBarType.warning}>
                        The CSV file is empty. Please check the file and try again.
                    </MessageBar>
                )}

                {!loading && !data && availableFiles.length === 0 && !error && (
                    <MessageBar messageBarType={MessageBarType.info}>
                        No CSV files found. Please check the library name and folder path.
                    </MessageBar>
                )}
            </div>
        );
    }
}
