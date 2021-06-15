import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PnpCrudWebPartStrings';
import PnpCrud from './components/PnpCrud';
import { IPnpCrudProps } from './components/IPnpCrudProps';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";

export interface IPnpCrudWebPartProps {
  description: string;
}

export default class PnpCrudWebPart extends BaseClientSideWebPart<IPnpCrudWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPnpCrudProps> = React.createElement(
      PnpCrud,
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
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
  

  
    return super.onInit();
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
