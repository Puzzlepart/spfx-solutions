import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneCustomField
} from '@microsoft/sp-webpart-base';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { HttpClient } from '@microsoft/sp-http';

import styles from './DailyDilbertWebPart.module.scss';

export interface IDailyDilbertWebPartProps {
    FlowUrl: string;
}

export default class DailyDilbertWebPart extends BaseClientSideWebPart<IDailyDilbertWebPartProps> {

    public async getComic() {
        let cache = window.localStorage.getItem("SPFXDilbert");
        // Subtract 12h to make sure we're in a valid day before a new comic is released
        let date = new Date(new Date().toISOString());
        date.setHours(date.getHours() - 12);
        let today = date.getFullYear() + "-" + this.pad(date.getMonth() + 1) + '-' + this.pad(date.getDate());
        let dilbert: any;
        if (cache) {
            let data = JSON.parse(cache);
            if (data.Date === today) {
                dilbert = data;
            }
        }
        if (!dilbert) {
            this.domElement.innerHTML = "Loading Dilbert...";
            console.log("Fetching todays Dilbert URL - " + today);
            let result = await this.context.httpClient.get(this.properties.FlowUrl, HttpClient.configurations.v1);
            if (result.ok) {
                dilbert = await result.json();
                window.localStorage.setItem("SPFXDilbert", JSON.stringify(dilbert));
            } else {
                this.domElement.innerHTML = 'Someting went wrong...go to <a target="_blank" href="http://dilbert.com/">http://dilbert.com</a> instead'; 
            }
        }
        if (dilbert) {
            this.domElement.innerHTML = `
            <div class="${styles.dailyDilbert}">
                <h3 class="${styles.title}">${dilbert.Title.trim()}</h3>
                <img src="${dilbert.Url.trim()}" title="${dilbert.Description.trim()}" width="100%">
            </div>
        `;
        }
    }

    public render(): void {
        if (isEmpty(this.properties.FlowUrl)) {
            this.domElement.innerHTML = "Missing Flow URL to fetch latest Dilbert comic. Please edit the web part.";
        } else {
            this.getComic();
        }
    }

    protected renderLogo(domElement: HTMLElement) {
        domElement.innerHTML = `
      <div style="margin-top: 30px">
        <div style="float:right">Author: <a href="mailto:mikael.svenson@puzzlepart.com" tabindex="-1">Mikael Svenson</a></div>
        <div style="float:right"><a href="https://www.puzzlepart.com/" target="_blank"><img src="//www.puzzlepart.com/wp-content/uploads/2017/08/Pzl-LogoType-200.png" onerror="this.style.display = \'none\'";"></a></div>
      </div>`;
    }

    protected pad(number: number) {
        if (number.toString().length === 1) {
            return "0" + number;
        }
        return number.toString();
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
                                PropertyPaneTextField('FlowUrl', {
                                    label: "Flow URL",
                                    value: this.properties.FlowUrl
                                }),
                                PropertyPaneCustomField({
                                    onRender: this.renderLogo,
                                    key: "logo"
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
