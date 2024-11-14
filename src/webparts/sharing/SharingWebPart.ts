
import * as React from 'react'
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { initializeFileTypeIcons } from '@fluentui/react-file-type-icons';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { ISharingWebPartContext } from './model';
import { SharingWebPartContext } from './hooks/SharingWebPartContext';
import SharingApp from './components/SharingApp';

export interface ISharingWebPartProps {
  webpartTitle: string;
  debugMode: boolean;
  preQuery: string;
}

export default class SharingWebPart extends BaseClientSideWebPart<ISharingWebPartProps> {

  protected async onInit(): Promise<void> {
    // load the filetype icons and other icons
    initializeIcons(undefined, { disableWarnings: true });
    initializeFileTypeIcons();
  }

  public render(): void {
    const sharingWebPartContextValue: ISharingWebPartContext = {
      pageLimit: 15,
      webpartContext: this.context,
      isTeams: this.context.sdks.microsoftTeams ? true : false,
      webpartProperties: this.properties
    };
    // Put the context & webpart propertey values in Provider
    const element: React.ReactElement = React.createElement(
      SharingWebPartContext.Provider,
      {
        value: sharingWebPartContextValue
      },
      React.createElement(SharingApp)
    );

    // eslint-disable-next-line @microsoft/spfx/pair-react-dom-render-unmount
    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Configuration",
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: "Webpart title"
                }),
                PropertyPaneTextField('preQuery', {
                  label: "Query",
                  multiline: true
                }),
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
