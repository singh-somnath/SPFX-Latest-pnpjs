import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxPnpjxWebPartStrings';
import SpfxPnpjx from './components/SpfxPnpjx';
import { ISpfxPnpjxProps } from './components/ISpfxPnpjxProps';
import {WebPartContext} from '@microsoft/sp-webpart-base'

export interface ISpfxPnpjxWebPartProps {
  description: string;


}

export default class SpfxPnpjxWebPart extends BaseClientSideWebPart<ISpfxPnpjxWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxPnpjxProps> = React.createElement(
      SpfxPnpjx,
      {
        description: this.properties.description,
        spcontext:this.context
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
