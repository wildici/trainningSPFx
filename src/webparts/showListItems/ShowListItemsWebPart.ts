import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ShowListItemsWebPartStrings';
import ShowListItems from './components/ShowListItems';
import { IShowListItemsProps } from './components/IShowListItemsProps';
import { setup as pnpSetup } from "@pnp/common";


export interface IShowListItemsWebPartProps {
  description: string;
  listName: string;
}

export default class ShowListItemsWebPart extends BaseClientSideWebPart<IShowListItemsWebPartProps > {

// ...
protected onInit(): Promise<void> {
  return super.onInit().then(_ => {
    // other init code may be present
    pnpSetup({
      spfxContext: this.context
    });
  });
}
// ...

  public render(): void {
    const element: React.ReactElement<IShowListItemsProps> = React.createElement(
      ShowListItems,
      {
        description: this.properties.description,
        listName: this.properties.listName
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
                }),
                PropertyPaneTextField('listName', {
                  label: 'Nome da lista'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
