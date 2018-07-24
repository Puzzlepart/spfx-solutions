import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { IODataList } from '@microsoft/sp-odata-types';
import * as pnp from "sp-pnp-js";
import * as strings from 'TilesWebPartStrings';
import Tiles from './components/Tiles';
import { ITilesProps } from './components/ITilesProps';

export interface ITilesWebPartProps {
  list: string;
  descriptionField: string;
  backgroundImageField: string;
  fallbackImageUrl: string;
  newTabField: string;
  orderByField: string;
  linkField: string;
  count: number;
  imageWidth: number;
  imageHeight: number;
  textPadding: number;
  relativeSiteUrl: string;
  tileType: string;
  tileTypeField: string;
  showAdvanced: boolean;
}

export default class TilesWebPart extends BaseClientSideWebPart<ITilesWebPartProps> {

  private listOptions: IPropertyPaneDropdownOption[];
  private tileTypeFieldOptions: IPropertyPaneDropdownOption[];
  private tileTypeOptions: IPropertyPaneDropdownOption[];
  private tileTypeFieldDropdownDisabled: boolean;
  private tileTypeDropdownDisabled: boolean;
  private listsDropdownDisabled: boolean;
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<ITilesProps> = React.createElement(
      Tiles,
      {
        list: this.properties.list,
        title: "Title",
        descriptionField: this.properties.descriptionField,
        backgroundImageField: this.properties.backgroundImageField,
        fallbackImageUrl: this.properties.fallbackImageUrl,
        newTabField: this.properties.newTabField,
        linkField: this.properties.linkField,
        orderByField: this.properties.orderByField,
        count: this.properties.count,
        imageWidth: this.properties.imageWidth,
        imageHeight: this.properties.imageHeight,
        textPadding: this.properties.textPadding,
        webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        tileType: this.properties.tileType,
        tileTypeField: this.properties.tileTypeField
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  protected async onPropertyPaneConfigurationStart() {
    try {
      await this.loadListDropdown();
      await this.loadTileTypeFieldDropdown();
      await this.loadTileTypeDropdown();
    } catch (error) {
      throw error;
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'list' &&
      newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      const previousTileTypeFieldValue: string = this.properties.tileTypeField;
      this.properties.tileTypeField = undefined;
      this.onPropertyPaneFieldChanged('tileTypeField', previousTileTypeFieldValue, this.properties.tileTypeField);
      this.tileTypeFieldDropdownDisabled = true;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'tileTypeField');
      this.fetchChoiceFieldTypes()
        .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
          this.tileTypeFieldOptions = itemOptions;
          this.tileTypeFieldDropdownDisabled = false;
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
          this.context.propertyPane.refresh();
        });
    }
    else if (propertyPath === 'tileTypeField' &&
      newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      const previousItem: string = this.properties.tileType;
      this.properties.tileType = undefined;
      this.onPropertyPaneFieldChanged('tileType', previousItem, this.properties.tileType);
      this.tileTypeDropdownDisabled = true;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'tileType');
      this.fetchFileTypeOptions()
        .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
          this.tileTypeOptions = itemOptions;
          this.tileTypeOptions.push({ text: "<Ingen verdi>", key: "" });
          this.tileTypeDropdownDisabled = false;
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
          this.context.propertyPane.refresh();
        });
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }
  private async loadListDropdown() {
    this.listsDropdownDisabled = !this.listOptions;
    if (!this.listOptions) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'list');
      let response = await this.fetchListOptions();
      this.listOptions = response;
      this.listsDropdownDisabled = false;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }
  }

  private async loadTileTypeFieldDropdown() {
    this.tileTypeFieldDropdownDisabled = !this.tileTypeFieldOptions;
    if (!this.tileTypeFieldOptions && this.properties.list) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'tileTypeField');
      let response = await this.fetchChoiceFieldTypes();
      this.tileTypeFieldOptions = response;
      this.tileTypeFieldDropdownDisabled = false;
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }
  }

  private async loadTileTypeDropdown() {
    this.tileTypeDropdownDisabled = !this.tileTypeOptions;
    if (!this.tileTypeOptions && this.properties.list && this.properties.tileTypeField) {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'tileType');
      let response = await this.fetchFileTypeOptions();
      this.tileTypeOptions = response;
      this.tileTypeOptions.push({ text: "<Ingen verdi>", key: "" });
      this.tileTypeDropdownDisabled = false;
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

              groupName: "Innstillinger",
              groupFields: [
                PropertyPaneDropdown('list', {
                  label: 'Liste',
                  options: this.listOptions,
                  disabled: this.listsDropdownDisabled
                }),
                PropertyPaneDropdown('tileTypeField', {
                  label: 'Felt for flistype',
                  options: this.tileTypeFieldOptions,
                  disabled: this.tileTypeFieldDropdownDisabled
                }),
                PropertyPaneDropdown('tileType', {
                  label: 'Flistype',
                  options: this.tileTypeOptions,
                  disabled: this.tileTypeDropdownDisabled
                }),
                PropertyPaneToggle('showAdvanced', {
                  label: 'Avanserte innstillinger',
                  offText: "Vis",
                  onText: "Skjul"
                }),
              ]
            },
            {
              isCollapsed: !this.properties.showAdvanced,
              groupFields: [
                PropertyPaneTextField('descriptionField', {
                  label: strings.DescriptionFieldLabel,
                  description: strings.DescriptionFieldDescription,
                }),
                PropertyPaneTextField('backgroundImageField', {
                  label: strings.BackgroundImageFieldLabel,
                  description: strings.BackgroundImageFieldDescription,
                }),
                PropertyPaneTextField('fallbackImageUrl', {
                  label: strings.FallbackImageUrlLabel,
                  description: strings.FallbackImageUrlDescription,
                }),
                PropertyPaneTextField('newTabField', {
                  label: strings.NewTabFieldLabel,
                  description: strings.NewTabFieldDescription,
                }),
                PropertyPaneTextField('linkField', {
                  label: strings.LinkFieldLabel,
                  description: strings.LinkFieldDescription,
                }),
                PropertyPaneTextField('orderByField', {
                  label: strings.OrderByFieldLabel,
                  description: strings.OrderByFieldDescription,
                }),
              ]
            },
            {
              isCollapsed: !this.properties.showAdvanced,
              groupFields: [
                PropertyPaneSlider('count', {
                  label: strings.CountFieldLabel,
                  min: 1,
                  max: 20
                }),
                PropertyPaneSlider('imageWidth', {
                  label: strings.ImageWidthFieldLabel,
                  min: 100,
                  max: 500
                }),
                PropertyPaneSlider('imageHeight', {
                  label: strings.ImageHeightFieldLabel,
                  min: 100,
                  max: 500
                }),
                PropertyPaneSlider('textPadding', {
                  label: strings.TextPaddingFieldLabel,
                  min: 2,
                  max: 20
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  private async fetchListOptions(): Promise<IPropertyPaneDropdownOption[]> {
    try {
      const results = await pnp.sp.web.lists.filter('BaseTemplate eq 100 and Hidden eq false').get();
      return results.map((item: IODataList, index) => {
        return { text: item.Title, key: item.Title, index: index };
      });
    } catch (error) {
      throw error;
    }
  }

  private async fetchChoiceFieldTypes(): Promise<IPropertyPaneDropdownOption[]> {
    let fields = await pnp.sp.web.lists.getByTitle(this.properties.list).fields.filter("TypeAsString eq 'Choice'").get();
    let options: Array<IPropertyPaneDropdownOption> = fields.map((field) => {
      return { key: field.InternalName, text: field.Title };
    });
    return options;
  }

  private async fetchFileTypeOptions(): Promise<IPropertyPaneDropdownOption[]> {
    let field = await pnp.sp.web.lists.getByTitle(this.properties.list).fields.getByInternalNameOrTitle(this.properties.tileTypeField).get();
    let options: Array<IPropertyPaneDropdownOption> = field.Choices.map((choice) => {
      return { key: choice, text: choice };
    });
    return options;
  }

  private validateFields(value: string): string {
    return (value) ? "" : "Vennligst fyll inn verdi.";
  }
}
