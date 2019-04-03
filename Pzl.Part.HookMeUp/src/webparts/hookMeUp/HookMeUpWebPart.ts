import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import PropertyPaneLogo from './PropertyPaneLogo';
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'HookMeUpWebPartStrings';

export interface IHookMeUpWebPartProps {
    anchorId: string;
}

export default class HookMeUpWebPart extends BaseClientSideWebPart<IHookMeUpWebPartProps> {

    public render(): void {
        if (this.displayMode === DisplayMode.Edit) {
            this.domElement.innerHTML = `
                <div id="${this.properties.anchorId}" style="margin-bottom:50px">
                    <b>Anchor:</b> ${this.properties.anchorId} <i>(${strings.NotVisible})</i>
                </div>
            `;
        } else {
            let element: HTMLElement = this.domElement.parentElement;
            // check up to 5 levels up for padding
            for (let i: number = 0; i < 5; i++) {
                const style: CSSStyleDeclaration = window.getComputedStyle(element);
                const hasPadding: boolean = style.paddingTop !== '0px';
                if (hasPadding) {
                    element.style.paddingTop = '0';
                    element.style.paddingBottom = '0';
                    element.style.marginTop = '0';
                    element.style.marginBottom = '0';
                }
                if (element.className === 'ControlZone') {
                    break;
                }
                element = element.parentElement;
            }
            this.domElement.innerHTML = this.domElement.innerHTML = `<span id="${this.properties.anchorId}"></span>`;
            if (this.displayMode == DisplayMode.Read && document.location.hash === '#' + this.properties.anchorId) {
                this.domElement.scrollIntoView();
            }
        }
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
                                PropertyPaneTextField('anchorId', {
                                    label: strings.AnchorId,
                                    description: strings.Description,
                                    value: this.properties.anchorId
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
