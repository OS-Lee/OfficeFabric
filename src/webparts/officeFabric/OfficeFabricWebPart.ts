import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'OfficeFabricWebPartStrings';
import OfficeFabric from './components/OfficeFabric';
import { IOfficeFabricProps } from './components/IOfficeFabricProps';

export interface IOfficeFabricWebPartProps {
  description: string;
}

export default class OfficeFabricWebPart extends BaseClientSideWebPart<IOfficeFabricWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IOfficeFabricProps > = React.createElement(
      OfficeFabric,
      {
        description: this.properties.description,
        context:this.context
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
