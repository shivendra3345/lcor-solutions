export interface IReportsProps {
  description: string;
  libraryName?: string;
  folderPath?: string;
  fileName?: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context?: any;
  // per-chart visibility and axis-name hiding maps (title -> boolean)
  chartVisibilities?: { [title: string]: boolean };
  hideAxisNames?: { [title: string]: boolean };
  // per-chart label overrides (sanitized key -> label)
  chartLabels?: { [title: string]: string };
  // configurable report title shown as the page heading
  reportTitle?: string;
}
