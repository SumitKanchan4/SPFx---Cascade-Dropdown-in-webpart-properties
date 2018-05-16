import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'DemoCascadeWebPartStrings';
import DemoCascade from './components/DemoCascade';
import { IDemoCascadeProps } from './components/IDemoCascadeProps';
import { SPCommonOperations } from 'spfxhelper';

export interface IDemoCascadeWebPartProps {
  description: string;
}

export default class DemoCascadeWebPart extends BaseClientSideWebPart<IDemoCascadeWebPartProps> {

  private ddlListOptions: IPropertyPaneDropdownOption[] = [];
  private ddlViewsOptions: IPropertyPaneDropdownOption[] = [];
  private listsLoaded: boolean = false;

  private get oSPComOps(): SPCommonOperations {
    return SPCommonOperations.getInstance(this.context.spHttpClient as any, this.context.pageContext.web.absoluteUrl);
  }

  private get webUrl():string{
    return this.context.pageContext.web.absoluteUrl;
  }

  public render(): void {
    const element: React.ReactElement<IDemoCascadeProps> = React.createElement(
      DemoCascade,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get getAllLists(): Promise<IPropertyPaneDropdownOption[]> {

    let ddlOptions: IPropertyPaneDropdownOption[] = [];
    return this.oSPComOps.queryGETResquest(`${this.webUrl}/_api/web/lists`).then(resp => {
      if (resp.ok) {
        resp.result.value.forEach(lists => {
          
          ddlOptions.push({
            key: lists.Title,
            text: lists.Title
          });
        });
      }

      return Promise.resolve(ddlOptions);
    });
  }

  protected getAllViews(selectedList: string): Promise<IPropertyPaneDropdownOption[]> {
    let ddlViewOptions: IPropertyPaneDropdownOption[] = [];
    return this.oSPComOps.queryGETResquest(`${this.webUrl}/_api/web/lists/getByTitle('${selectedList}')/Views`).then(respViews => {
      if (respViews.ok) {
        respViews.result.value.forEach(views => {
          ddlViewOptions.push({
            key: views.Title,
            text: views.Title
          });
        });
      }

      return Promise.resolve(ddlViewOptions);
    });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

    if (propertyPath == `selectedList`) {
      this.getAllViews(newValue).then(resView => {
        this.ddlViewsOptions = resView;
        this.context.propertyPane.refresh();
      });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    // Check if the values are recieved
    if (!this.listsLoaded) {
      // Call the method to get all the lists Titles
      this.getAllLists.then(resp => {

        // Fill the values in the variable assigned
        this.ddlListOptions = resp;

        // update the flag so it is not called again
        this.listsLoaded = true;

        // Refresh the property pane, to reflect the changes
        this.context.propertyPane.refresh();
      });
    }

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
                PropertyPaneDropdown(`selectedList`, {
                  label: `select List`,
                  options: this.ddlListOptions
                }),
                PropertyPaneDropdown(`selectedView`,{
                  label: `select view`,
                  options: this.ddlViewsOptions,
                  disabled: this.ddlListOptions.length == 0
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
