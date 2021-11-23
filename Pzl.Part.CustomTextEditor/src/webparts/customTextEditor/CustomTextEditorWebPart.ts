import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, IPropertyPaneField, PropertyPaneChoiceGroup, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import * as strings from 'CustomTextEditorWebPartStrings';
import CustomTextEditor from './components/CustomTextEditor';
import { TextBoxStyle } from "./components/TextBoxStyle";
import { ICustomTextEditorProps } from './components/ICustomTextEditorProps';
import { ICustomTextEditorWebPartProps } from './ICustomTextEditorWebPartProps';
import {
    ThemeProvider,
    ThemeChangedEventArgs,
    IReadonlyTheme
} from '@microsoft/sp-component-base';

/*
Nothing really special in this class, just integartes it with sharepoint.
*/

export default class CustomTextEditorWebPart extends BaseClientSideWebPart<ICustomTextEditorWebPartProps> {

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

    public colorOptions = [
        {
            'key': 'factbox',
            'text': 'Faktaboks',
        },
        {
            'key': 'changelog',
            'text': 'Endringslogg',
        },
        {
            'key': 'aside',
            'text': 'Tilleggsopplysninger (spørsmål, kontaktinfo)',
        },
        {
            'key': 'other',
            'text': 'Generell boks',
        },
        {
            'key': 'none',
            'text': 'Ingen bakgrunnsfarge',
        },
    ];

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
                backgroundColorChoice: this.properties.backgroundColorChoice
                    ? this.properties.backgroundColorChoice
                    : this.properties.backgroundColor
                        ? 'other'
                        : 'none',
                useBorder: this.properties.useBorder,
                useBottomBorder: this.properties.useBottomBorder,
                themeVariant: this._themeVariant,
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

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        let propertyControls: IPropertyPaneField<any>[] = [];
        propertyControls.push(PropertyPaneChoiceGroup("textBoxStyle", {
            options: [
                { text: strings.StandardLabel, key: TextBoxStyle.Regular, iconProps: { officeFabricIconFontName: "TextBox" } },
                { text: strings.AccordionLabel, key: TextBoxStyle.Accordion, iconProps: { officeFabricIconFontName: "Dropdown" } },
                { text: strings.BackgroundLabel, key: TextBoxStyle.WithBackgroundColor, iconProps: { officeFabricIconFontName: "BackgroundColor" } }
            ],
        }));

        switch (this.properties.textBoxStyle) {
            case TextBoxStyle.Regular:
                propertyControls.push(PropertyPaneToggle('useBorder', {label: "Bruk ramme på tekstboksen"}));
            break;

            case TextBoxStyle.Accordion:
                propertyControls.push(PropertyPaneToggle('useBottomBorder', {label: "Vis skillelinje mellom trekkspill"}));
            break;

            case TextBoxStyle.WithBackgroundColor:
                propertyControls.push(PropertyPaneChoiceGroup('backgroundColorChoice', {label: "Bakgrunnsfarge", options: this.colorOptions}));
            break;

            default: break;
        }

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

    /**
     * Update the current theme variant reference and re-render.
     *
     * @param args The new theme
     */
     private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
        this._themeVariant = args.theme;
        this.render();
    }

}
