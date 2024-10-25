
import * as React from 'react'
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';
import { Providers } from "@microsoft/mgt-element";
import { SharePointProvider } from "@microsoft/mgt-sharepoint-provider";
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

  protected async onInit(): Promise<void> {
    // load the filetype icons and other icons
    initializeIcons(undefined, { disableWarnings: true });
    initializeFileTypeIcons();

    // setting up the logging framework
    Logger.subscribe(ConsoleListener(LOG_SOURCE));
    Logger.activeLogLevel = (this.properties.debugMode) ? LogLevel.Verbose : LogLevel.Warning;

    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }

    // if you don't want to send telemetry data to PnP, you can opt-out here (see https://github.com/pnp/telemetry-js for details on what is being sent)
    // const telemetry = PnPTelemetry.getInstance();
    // telemetry.optOut();
  }

  public render(): void {
    const sharingWebPartContextValue: ISharingWebPartContext = {
      pageLimit: 15,
      webpartContext: this.context,
      isTeams: this.context.sdks.microsoftTeams ? true : false
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
    //     isTeams: false,
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
