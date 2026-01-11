import * as React from 'react';
import styles from './Reports.module.scss';
import type { IReportsProps } from './IReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import CSVReportViewer from './CSVReportViewer';

export default class Reports extends React.Component<IReportsProps> {
  public render(): React.ReactElement<IReportsProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.reports} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* CSV Report Viewer Component */}
        <CSVReportViewer
          libraryName={this.props['libraryName'] || 'Shared Documents'}
          folderPath={this.props['folderPath'] || 'Reports'}
          fileName={this.props['fileName']}
          isDarkTheme={isDarkTheme}
          // pass per-chart settings from web part
          chartVisibilities={this.props.chartVisibilities}
          hideAxisNames={this.props.hideAxisNames}
          context={this.props.context}
        />


        <div>

        </div>
      </section>
    );
  }
}

