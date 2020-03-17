import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneCheckbox } from "@microsoft/sp-property-pane";
import * as strings from 'SimpleSearchBoxWebPartStrings';
import SimpleSearchBox from './components/SimpleSearchBox';
import { ISimpleSearchBoxProps } from './components/ISimpleSearchBoxProps';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ISimpleSearchBoxWebPartProps {
  searchurl: string;
  title: string;
  openInNewTab: boolean;
  placeholder: string;
}

export default class SearchCentreWebPart extends BaseClientSideWebPart<ISimpleSearchBoxWebPartProps> {

  private _themeProvider: ThemeProvider;
private _themeVariant: IReadonlyTheme | undefined;
  
protected onInit(): Promise<void> {
    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit();
}

/**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
}
  
  public render(): void {
    const element: React.ReactElement<ISimpleSearchBoxProps> = React.createElement(
      SimpleSearchBox,
      {
        searchurl: this.properties.searchurl,
        title: this.properties.title,
        openInNewTab: this.properties.openInNewTab,
        displayMode: this.displayMode,
        placeholder: this.properties.placeholder,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        themeVariant: this._themeVariant,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('searchurl', {
                  label: strings.UrlFieldLabel
                }),
                PropertyPaneTextField('placeholder', {
                  label: strings.PlaceholderFieldLabel
                }),
                PropertyPaneCheckbox('openInNewTab', {
                  text: strings.OpenInNewTabLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
