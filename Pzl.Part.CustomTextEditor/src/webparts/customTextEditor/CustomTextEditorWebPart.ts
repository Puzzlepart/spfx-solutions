import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneCheckbox, PropertyPaneChoiceGroup } from "@microsoft/sp-property-pane";
import * as strings from 'CustomTextEditorWebPartStrings';
import CustomTextEditor from './components/CustomTextEditor';
import { TextBoxStyle } from "./components/TextBoxStyle";
import { ICustomTextEditorProps } from './components/ICustomTextEditorProps';
import { ICustomTextEditorWebPartProps } from './ICustomTextEditorWebPartProps';

/*
Nothing really special in this class, just integartes it with sharepoint.
*/
export default class CustomTextEditorWebPart extends BaseClientSideWebPart<ICustomTextEditorWebPartProps> {
    private _colorPickerComponent: any;

    public render(): void {
        const element: React.ReactElement<ICustomTextEditorProps> = React.createElement(
            CustomTextEditor,
            {
                title: this.properties.title,
                displayMode: this.displayMode,
                setTitle: this.setTitle.bind(this),
                saveRteContent: this.setRteContentProp.bind(this),
                isReadMode: DisplayMode.Read === this.displayMode,
                content: this.properties.Content,
                textBoxStyle: this.properties.textBoxStyle,
                backgroundColor: this.properties.backgroundColor,
                headerExpandColor: this.properties.headerExpandColor,
                underlineLinks: typeof this.properties.underlineLinks === 'undefined' ? true : this.properties.underlineLinks
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected get propertiesMetadata(): IWebPartPropertiesMetadata {
        return { searchableContent: { isHtmlString: true } };
    }

    protected async loadPropertyPaneResources(): Promise<void> {
        let component = await import(
            /* webpackChunkName: 'color-picker' */
            '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker'
        );
        this._colorPickerComponent = component;
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        let propertyControls: IPropertyPaneField<any>[] = [];
        propertyControls.push(PropertyPaneChoiceGroup("textBoxStyle", {
            options: [
                { text: strings.StandardLabel, key: TextBoxStyle.Regular, iconProps: { officeFabricIconFontName: "TextBox" } },
                { text: strings.StandardLabelFade, key: TextBoxStyle.RegularFade, iconProps: { officeFabricIconFontName: "AddNotes" } },
                { text: strings.AccordionLabel, key: TextBoxStyle.Accordion, iconProps: { officeFabricIconFontName: "Dropdown" } },
                { text: strings.BackgroundLabel, key: TextBoxStyle.WithBackgroundColor, iconProps: { officeFabricIconFontName: "BackgroundColor" } }
            ],
        }));

        if (this.properties.textBoxStyle === TextBoxStyle.WithBackgroundColor) {
            propertyControls.push(this._colorPickerComponent.PropertyFieldColorPicker("backgroundColor", {
                label: 'Color',
                selectedColor: this.properties.backgroundColor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                alphaSliderHidden: false,
                style: this._colorPickerComponent.PropertyFieldColorPickerStyle.Inline,
                iconName: 'Precipitation',
                key: 'colorFieldId'
            }));
        }

        if (this.properties.textBoxStyle === TextBoxStyle.Accordion) {
            propertyControls.push(this._colorPickerComponent.PropertyFieldColorPicker("headerExpandColor", {
                label: 'Color',
                selectedColor: this.properties.headerExpandColor,
                onPropertyChange: this.onPropertyPaneFieldChanged,
                properties: this.properties,
                disabled: false,
                alphaSliderHidden: false,
                style: this._colorPickerComponent.PropertyFieldColorPickerStyle.Inline,
                iconName: 'Precipitation',
                key: 'colorFieldIdHeadline'
            }));
        }
        propertyControls.push(PropertyPaneCheckbox("underlineLinks", { text: strings.LinkUnderline, checked: this.properties.underlineLinks }));

        return {
            pages: [
                {
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: propertyControls
                        }
                    ]
                }
            ]
        };
    }

    private setRteContentProp(content: string): void {
        this.properties.Content = content;
        this.properties.searchableContent = `${this.properties.title}|${content}`;
    }

    private setTitle(title: string) {
        this.properties.title = title;
        this.properties.searchableContent = `${this.properties.title}|${this.properties.Content}`;
    }
}