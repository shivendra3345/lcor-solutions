import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneButton
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ReportsWebPartStrings';
import Reports from './components/Reports';
import { IReportsProps } from './components/IReportsProps';
import { CSVDataService } from './services/CSVDataService';

export interface IReportsWebPartProps {
  description: string;
  libraryName?: string;
  folderPath?: string;
  fileName?: string;
  chartVisibilities?: { [title: string]: boolean };
  hideAxisNames?: { [title: string]: boolean };
}

export default class ReportsWebPart extends BaseClientSideWebPart<IReportsWebPartProps> {
  private _dynamicToggleFields: any[] = [];

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IReportsProps> = React.createElement(
      Reports,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        libraryName: this.properties.libraryName,
        folderPath: this.properties.folderPath,
        fileName: this.properties.fileName,
        chartVisibilities: this.properties.chartVisibilities,
        hideAxisNames: this.properties.hideAxisNames
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    // Initialize CSVDataService with SPFx context
    CSVDataService.initialize(this.context);

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('libraryName', {
                  label: 'Library name',
                  placeholder: 'e.g., Shared Documents'
                }),
                PropertyPaneTextField('folderPath', {
                  label: 'Folder path (optional)',
                  placeholder: 'e.g., Reports/2024'
                }),
                PropertyPaneTextField('fileName', {
                  label: 'CSV file name (e.g., data.csv)'
                }),
                PropertyPaneButton('refreshTitles', {
                  text: 'Refresh Titles',
                  buttonType: 0,
                  onClick: async () => {
                    // allow manual refresh of dynamic toggles for the property pane
                    await this.buildDynamicToggles();
                    this.context.propertyPane.refresh();
                  }
                })
              ]
            },
            {
              groupName: 'Chart Visibility',
              groupFields: [
                // dynamic per-chart toggles will be appended here; if none exist show a helper button
                ...(this._dynamicToggleFields.length > 0 ? this._dynamicToggleFields : [
                  PropertyPaneButton('refreshTitlesPlaceholder', {
                    text: 'No titles found. Click to refresh',
                    buttonType: 0,
                    onClick: async () => {
                      await this.buildDynamicToggles();
                      this.context.propertyPane.refresh();
                    }
                  })
                ])
              ]
            }
          ]
        }
      ]
    };
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this.buildDynamicToggles();
  }

  private async buildDynamicToggles(): Promise<void> {
    // Build dynamic toggles for each Title in the selected CSV file (if available)
    this._dynamicToggleFields = [];

    const lib = this.properties.libraryName || 'Shared Documents';
    const folder = this.properties.folderPath || '';
    const file = this.properties.fileName || '';

    if (!file) {
      // nothing to build
      return;
    }

    try {
      // Try to use any cached CSV that may have been fetched by the client-side Reports component
      const serverRel = CSVDataService.buildServerRelativePath(lib, folder, file);
      let csv = CSVDataService.getCachedCSV(serverRel);
      if (!csv) {
        csv = await CSVDataService.fetchCSVFromSharePoint(lib, folder, file);
      }
      const titleHeader = (csv.headers && csv.headers.length > 1) ? csv.headers[1] : 'Title';
      const titles = Array.from(new Set(csv.rows.map(r => String(r[titleHeader] || '').trim()))).filter(t => t);

      // Create toggles bound directly to web part properties using dot-paths
      if (titles.length === 0) {
        console.warn('No Titles found in CSV; attempting to use persisted visibility keys if available.');
      }

      // Prefer the current CSV titles, but fall back to any persisted keys if CSV is empty
      const persistedKeys = Object.keys(this.properties.chartVisibilities || {});
      const sourceTitles = titles.length > 0 ? titles : (persistedKeys.length > 0 ? persistedKeys.map(k => k) : []);

      sourceTitles.forEach((titleOrKey) => {
        // If we got a persisted key (which is sanitized), try to show a readable label
        const isSanitizedKey = !!this.properties.chartVisibilities && !!this.properties.chartVisibilities[titleOrKey as string] && titles.indexOf(titleOrKey as string) === -1;
        const displayLabel = isSanitizedKey ? (titleOrKey as string).replace(/_/g, ' ') : (titleOrKey as string);
        const key = isSanitizedKey ? (titleOrKey as string) : this._sanitizeTitleKey(titleOrKey as string);

        const visProp = `chartVisibilities.${key}`;
        const axisProp = `hideAxisNames.${key}`;

        this._dynamicToggleFields.push(
          PropertyPaneToggle(visProp, {
            label: `${displayLabel} - Visible`,
            onText: 'Shown',
            offText: 'Hidden'
          })
        );

        this._dynamicToggleFields.push(
          PropertyPaneToggle(axisProp, {
            label: `${displayLabel} - Hide axis names`,
            onText: 'Yes',
            offText: 'No'
          })
        );
      });
    } catch (e) {
      console.warn('Could not build property pane dynamic fields', e);

      // If fetching failed, but there are persisted keys, still build toggles from them so users can control visibility
      const persisted = Object.keys(this.properties.chartVisibilities || {});
      if (persisted.length > 0) {
        persisted.forEach((key) => {
          const visProp = `chartVisibilities.${key}`;
          const axisProp = `hideAxisNames.${key}`;
          const displayLabel = key.replace(/_/g, ' ');

          this._dynamicToggleFields.push(
            PropertyPaneToggle(visProp, {
              label: `${displayLabel} - Visible`,
              onText: 'Shown',
              offText: 'Hidden'
            })
          );

          this._dynamicToggleFields.push(
            PropertyPaneToggle(axisProp, {
              label: `${displayLabel} - Hide axis names`,
              onText: 'Yes',
              offText: 'No'
            })
          );
        });
      }
    }
  }

  private _sanitizeTitleKey(title: string): string {
    // create a safe key for property paths: keep letters/numbers and replace others with underscore
    return title.replace(/[^a-zA-Z0-9]/g, '_');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    // If the CSV file name changed in the property pane, rebuild toggles automatically
    if (propertyPath === 'fileName') {
      // build and refresh asynchronously
      void this.buildDynamicToggles().then(() => this.context.propertyPane.refresh());
    }
    // If toggles use property-paths like `chartVisibilities.<key>` or `hideAxisNames.<key>`, map them back
    if (propertyPath && propertyPath.indexOf('chartVisibilities.') === 0) {
      const key = propertyPath.replace('chartVisibilities.', '');
      this.properties.chartVisibilities = this.properties.chartVisibilities || {};
      this.properties.chartVisibilities[key] = !!newValue;
      this.render();
      return;
    }

    if (propertyPath && propertyPath.indexOf('hideAxisNames.') === 0) {
      const key = propertyPath.replace('hideAxisNames.', '');
      this.properties.hideAxisNames = this.properties.hideAxisNames || {};
      this.properties.hideAxisNames[key] = !!newValue;
      this.render();
      return;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
