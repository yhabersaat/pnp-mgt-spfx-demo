import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PnPMgtSpFxDemoWebPartStrings';
import PnPMgtSpFxDemo from './components/PnPMgtSpFxDemo';
import { IPnPMgtSpFxDemoProps } from './components/IPnPMgtSpFxDemoProps';
import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

export interface IPnPMgtSpFxDemoWebPartProps {
  description: string;
}

export default class PnPMgtSpFxDemoWebPart extends BaseClientSideWebPart<IPnPMgtSpFxDemoWebPartProps> {

  protected async onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IPnPMgtSpFxDemoProps> = React.createElement(
      PnPMgtSpFxDemo,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
