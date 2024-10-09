/* eslint-disable @typescript-eslint/typedef */
import * as React from 'react'
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import "@pnp/graph/groups";

import { initializeIcons } from '@fluentui/react/lib/Icons';
import "@pnp/sp/webs";
import "@pnp/sp/search";
import IDataProvider from './components/SharingView/DataProvider';
import DataProvider from './components/SharingView/DataProvider';
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';

import {
  Logger,
  ConsoleListener,
  LogLevel
} from "@pnp/logging";
import { IPropertyPaneConfiguration, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { ISharingWebPartContext } from './model';
import { SharingWebPartContext } from './hooks/SharingWebPartContext';
import SharingApp from './components/SharingApp';

const LOG_SOURCE: string = 'Microsoft-Governance-Sharing';

export interface ISharingWebPartProps {
  description: string;
  debugMode: boolean;
}

export default class SharingWebPart extends BaseClientSideWebPart<ISharingWebPartProps> {
  private dataProvider: IDataProvider;

  protected async onInit(): Promise<void> {
    // load the filetype icons and other icons
    initializeIcons(undefined, { disableWarnings: true });
    initializeFileTypeIcons();

    // setting up the logging framework
    Logger.subscribe(ConsoleListener(LOG_SOURCE));
    Logger.activeLogLevel = (this.properties.debugMode) ? LogLevel.Verbose : LogLevel.Warning;

    // if you don't want to send telemetry data to PnP, you can opt-out here (see https://github.com/pnp/telemetry-js for details on what is being sent)
    // const telemetry = PnPTelemetry.getInstance();
    // telemetry.optOut();

    // loading the data provider to get access to the REST/Search API
    this.dataProvider = new DataProvider(this.context);
  }

  public render(): void {
    const sharingWebPartContextValue: ISharingWebPartContext = {
      pageLimit: 15,
      webpartContext: this.context,
      isTeams: this.context.sdks.microsoftTeams ? true : false,
      dataProvider: this.dataProvider
    };

    // Put the context value with Provider
    const element: React.ReactElement = React.createElement(
      SharingWebPartContext.Provider,
      {
        value: sharingWebPartContextValue
      },
      React.createElement(SharingApp)
    );

    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDom.render(element, this.domElement);



    // const element: React.ReactElement<ISharingViewProps> = React.createElement(
    //   SharingViewSingle,
    //   {
    //     pageLimit: 15,
    //     context: this.context,
    //     isTeams: isTeams,
    //     dataProvider: this.dataProvider
    //   }
    // );
    // // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    // ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: this.properties.description
          },
          groups: [
            {
              groupName: "Configuration",
              groupFields: [
                PropertyPaneToggle('debugMode', {
                  label: "Enable debug mode",
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }
}
