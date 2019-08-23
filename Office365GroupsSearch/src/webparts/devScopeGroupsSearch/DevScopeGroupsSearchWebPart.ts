import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext, PropertyPaneDropdown } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'DevScopeGroupsSearchWebPartStrings';
import DevScopeGroupsSearch from './components/DevScopeGroupsSearch';
import { IDevScopeGroupsSearchProps } from './components/IDevScopeGroupsSearchProps';

export interface IDevScopeGroupsSearchWebPartProps {
  description: string;
  context: WebPartContext;
  itemnumberproperty:"string";
  orderbyfieldproperty: "string";
  ordermodefieldproperty: "string";
}

export default class DevScopeGroupsSearchWebPart extends BaseClientSideWebPart<IDevScopeGroupsSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDevScopeGroupsSearchProps > = React.createElement(
      DevScopeGroupsSearch,
      {
        context: this.context,
        itemnumberproperty: this.properties.itemnumberproperty,
        orderbyfieldproperty: this.properties.orderbyfieldproperty,
        ordermodefieldproperty: this.properties.ordermodefieldproperty
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
                PropertyPaneTextField('itemnumberproperty', {
                  label: "Item Number",
                  value: "999",
                  maxLength: 3
                }),
                PropertyPaneDropdown('orderbyfieldproperty', {
                  label: 'Order By',
                  options: [
                    { key: '0', text: 'Group' },
                    { key: '1', text: 'Email' }
                  ], 
                  selectedKey: "0"
                }),
                PropertyPaneDropdown('ordermodefieldproperty', {
                  label: 'Ascending or Descending',
                  options: [
                    { key: 'asc', text: 'Ascending' },
                    { key: 'desc', text: 'Descending' }
                  ],
                  selectedKey: "asc"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
