import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'OneDriveRecordingsWebPartStrings';
import OneDriveRecordings from './components/OneDriveRecordings';
import { IOneDriveRecordingsProps } from './components/IOneDriveRecordingsProps';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import OneDriveRecordingStyles from './OneDriveRecordings.module.scss';

export interface IOneDriveRecordingsWebPartProps {
  description: string;
}

export default class OneDriveRecordingsWebPart extends BaseClientSideWebPart<IOneDriveRecordingsWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        // get information about the current user from the Microsoft Graph
        client
          .api('/communications/callRecords')
          .version('beta')
          .top(10)
          .orderby('endDateTime desc')
          .get((error, callRecords: any, rawResponse?: any) => {
            this.domElement.innerHTML = `
      <section class="${OneDriveRecordingStyles.section}">
      <div>
            <h1 class"${OneDriveRecordingStyles.welcome}">Recordings</h1>
            <h3 class="${OneDriveRecordingStyles.welcome}">Below is the list of calls records.</h3>
          </div>    
      </section>`;

            // List the latest call records based on what we got from the Graph
            this._renderCallList(callRecords.value);
          });
      });
  }

  private _renderCallList(
    calls: MicrosoftGraph.CallRecords.CallRecord[]
  ): void {
    let html: string = '';
    for (let index = 0; index < calls.length; index++) {
      html += `
      <ul class="${OneDriveRecordingStyles.list}>
      <li class="${OneDriveRecordingStyles.listItem}>Call Recording ${
        index + 1
      } - ${escape(calls[index].endDateTime)}</li>
      </ul>`;
    }
  }

  // Add the calls to the placeholder
  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty(
      '--linkHovered',
      semanticColors.linkHovered
    );
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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
