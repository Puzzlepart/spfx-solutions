import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'PageNavigationWebPartStrings';
import PageNavigation from './components/PageNavigation';
import { IPageNavigationProps } from './components/IPageNavigationProps';
import pnp from "sp-pnp-js";
import { IODataListItem } from '@microsoft/sp-odata-types';

export interface IPageNavigationWebPartProps {
  lookupField: string;
  topLevelPage: number;
}

export default class PageNavigationWebPart extends BaseClientSideWebPart<IPageNavigationWebPartProps> {
  private pageDropdownOptions: IPropertyPaneDropdownOption[];
  private pageDropdownDisabled: boolean;
  private lookupFieldOptions: IPropertyPaneDropdownOption[];
  private lookupFieldOptionsDisabled: boolean;
  public render(): void {
    const element: React.ReactElement<IPageNavigationProps > = React.createElement(
      PageNavigation,
      {
        lookupField: this.properties.lookupField,
        listServerRelativeUrl: this.context.pageContext.list.serverRelativeUrl,
        topLevelPage: this.properties.topLevelPage,
        serverRequestPath: this.context.pageContext.legacyPageContext.serverRequestPath,
        currentPage: this.context.pageContext.listItem
      }
    );
    ReactDom.render(element, this.domElement);
  }
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    try {
      await this.loadLookupDropdown();
      await this.loadPagesDropdown();
    } catch (error) {
      throw error;
    }
  }
  private async loadLookupDropdown() {
    this.lookupFieldOptionsDisabled = !this.lookupFieldOptions;
    if (!this.lookupFieldOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lookupField');
      let options = await this.fetchLookupValues();
      this.lookupFieldOptionsDisabled = false;
      this.lookupFieldOptions = options;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }
  }
  private async loadPagesDropdown() {
    this.pageDropdownDisabled = !this.pageDropdownOptions;
    if (!this.pageDropdownOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lookupField');
      let response = await this.fetchPagesOptions();
      this.pageDropdownDisabled = false;
      this.pageDropdownOptions = response;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }
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
                PropertyPaneDropdown('lookupField', {
                  label: strings.lookupFieldLabel,
                  options: this.lookupFieldOptions,
                }),
                PropertyPaneDropdown('topLevelPage', {
                  label: strings.topLevelPageFieldLabel,
                  options: this.pageDropdownOptions,
                }),
                PropertyPaneToggle('isRootExpanded', {
                  label: strings.isRootExpanded,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
  private async fetchPagesOptions(): Promise<IPropertyPaneDropdownOption[]> {
    try {
      let items = await pnp.sp.web.getList(this.context.pageContext.list.serverRelativeUrl).items.get();
      let options: Array<IPropertyPaneDropdownOption> = items.map((item: IODataListItem) => {
        return { key: item.ID, text: item.Title };
      });
      return options;
    } catch (error) {
      throw error;
    }
  }
  private async fetchLookupValues(): Promise<IPropertyPaneDropdownOption[]> {
    try {
      const odataFilter = "TypeAsString eq 'Lookup' and Hidden eq false and AllowMultipleValues eq false and LookupList ne ''";
      let fields = await pnp.sp.web.getList(this.context.pageContext.list.serverRelativeUrl).fields.filter(odataFilter).get();
      let options: Array<IPropertyPaneDropdownOption> = fields.map((field) => {
        return { key: field.InternalName, text: field.Title };
      });
      return options;
    } catch (error) {
      throw error;
    }
  }
}
