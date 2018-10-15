import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'BreadcrumbWebPartStrings';
import SitePagesBreadcrumb from './components/SitePagesBreadcrumb';
import { IBreadcrumbProps } from './components/IBreadcrumbProps';
import pnp from "sp-pnp-js";

export interface IBreadcrumbWebPartProps {
  description: string;
  lookupField: IPropertyPaneDropdownOption;
}

export default class BreadcrumbWebPart extends BaseClientSideWebPart<IBreadcrumbWebPartProps> {
  private lookupFieldOptions: IPropertyPaneDropdownOption[];
  private lookupFieldOptionsDisabled: boolean;

  public render(): void {
    const element: React.ReactElement<IBreadcrumbProps> = React.createElement(
      SitePagesBreadcrumb,
      {
        description: this.properties.description,
        currentPage: this.context.pageContext.listItem,
        listServerRelativeUrl: this.context.pageContext.list.serverRelativeUrl,
        lookupField: this.properties.lookupField
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
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    try {
      await this.loadLookupDropdown();
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
                PropertyPaneDropdown('lookupField', {
                  label: strings.lookupFieldLabel,
                  options: this.lookupFieldOptions,
                })
              ]
            }
          ]
        }
      ]
    };
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
