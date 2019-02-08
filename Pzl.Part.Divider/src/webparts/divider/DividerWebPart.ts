import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneSlider,
    PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { loadStyles } from '@microsoft/load-themed-styles';
import PropertyPaneLogo from './PropertyPaneLogo';
import styles from './DividerWebPart.module.scss';
import * as strings from 'DividerWebPartStrings';


export interface IDividerWebPartProps {
    width: number;
    color: string;
}

export default class DividerWebPart extends BaseClientSideWebPart<IDividerWebPartProps> {

    public render(): void {
        let color = this.properties.color;
        if (!color || color.trim().length === 0) {
            color = "[theme: neutralTertiaryAlt, default: #c8c8c8]";
        }

        // Enclose theme colors with quotes
        if (color.toLocaleLowerCase().indexOf("theme") !== -1) {
            color = '"' + color.trim() + '"';
        }

        let className = "pzl" + this.makeId();
        let propClass = `.${className} {
            background-color: ${color};
            width: ${this.properties.width}%;
        }`;

        let cssString = ``;
        if (this.displayMode == DisplayMode.Edit) {
            cssString = `margin-bottom:50px`;
        }

        loadStyles(propClass);
        this.domElement.innerHTML = `<hr aria-hidden="true" role="presentation" class="${className}" style="${cssString}">`;
    }

    protected renderLogo(domElement: HTMLElement) {
        domElement.innerHTML = `
      <div style="margin-top: 30px">
        <div style="float:right">Author: <a href="mailto:mikael.svenson@puzzlepart.com" tabindex="-1">Mikael Svenson</a></div>
        <div style="float:right"><a href="https://www.puzzlepart.com/" target="_blank"><img src="//www.puzzlepart.com/wp-content/uploads/2017/08/Pzl-LogoType-200.png" onerror="this.style.display = \'none\'";"></a></div>
      </div>`;
    }

    private makeId() {
        let text = "";
        let possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

        for (let i = 0; i < 5; i++)
            text += possible.charAt(Math.floor(Math.random() * possible.length));

        return text;
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
                                PropertyPaneSlider('width', {
                                    label: strings.WidthFieldLabel,
                                    min: 10,
                                    max: 100,
                                    showValue: true,
                                    step: 5,
                                    value: this.properties.width
                                }),
                                PropertyPaneTextField('color', {
                                    label: strings.ColorFieldLabel,
                                    description: strings.ColorFieldDescription,
                                    value: this.properties.color
                                }),
                                new PropertyPaneLogo()
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
