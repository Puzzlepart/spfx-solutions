import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import styles from './HighLightHomeFieldCustomizer.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';

const LOG_SOURCE: string = 'HighLightHomeFieldCustomizer';
export interface IHighLightHomeFieldCustomizerProperties {
}

let WELCOME_PAGE: string = '';

export default class HighLightHomeFieldCustomizer
    extends BaseFieldCustomizer<IHighLightHomeFieldCustomizerProperties> {

    @override
    public async onInit(): Promise<void> {
        let result = await this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/rootfolder?$select=welcomepage`,
            SPHttpClient.configurations.v1);
        let welcomePage = await result.json();
        WELCOME_PAGE = welcomePage.WelcomePage.replace("SitePages/", "");
        Log.info(LOG_SOURCE, `Welcome page is: ${WELCOME_PAGE}`);
    }

    @override
    public async onRenderCell(event: IFieldCustomizerCellEventParameters): Promise<void> {
        event.domElement.className = styles.ext;
        let fileLeafRef = event.listItem.getValueByName("FileLeafRef");
        if (WELCOME_PAGE === fileLeafRef) {
            Log.info(LOG_SOURCE, 'Highlighting welcome page');
            event.domElement.innerHTML = `<i class="${styles.homeIcon}" aria-hidden="true"></i>`;
            let row = this.getClosest(event.domElement, ".ms-DetailsRow-fields") as HTMLDivElement;
            if (row) {
                row.style.fontWeight = "800";
            }
        } else {
            event.domElement.innerHTML = '<img class="FileTypeIcon-icon" title="aspx File" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAUtJREFUeNpiYBgFlAFGQgqmTJnyn0Z2L8jJyUkkpIhpAAMnAej5+YQUsRBrWnZ2NlVcNXXqVHRHMuALyQELQQ0NDaJCcsAc6OzsTJQjBzINEuVIloFyHFpahDsSiBMHTQgSA0gOwQP3voHpg/e/ofBh9P92DbJKAxwhisWBFdcdoCwwPfH1ezAnp/LGoAnB/cic2z8FSTJwy5WXJKn30REftDUJUWDUgTTPxTMDJcB0mpkAivisUx/AdPr6F6MhiBfAQg49xGAhO5oGCSmAhRx6GqR12hs+IYgeUvQOyeFTDuIKSXRxQnXryAtBGEBv58FyNzo4efIkSQ4wNzcfIbmY2NxKKERGbhokFoymQSBwRO40qbK/ryelb0LtNIjpwA7NA7AeJojInzKlHrm7SKjbOeBp0EGJC4Wudx5tzVA0ljLaHhwFlAKAAAMApTB5hrOUthUAAAAASUVORK5CYII=" style="width: 20px; height: 20px;">';
        }
    }

    @override
    public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
        super.onDisposeCell(event);
    }

    private getClosest(elem, selector) {
        if (!Element.prototype.matches) {
            Element.prototype.matches =
                (<any>Element.prototype).matchesSelector ||
                (<any>Element.prototype).mozMatchesSelector ||
                Element.prototype.msMatchesSelector ||
                (<any>Element.prototype).oMatchesSelector ||
                Element.prototype.webkitMatchesSelector ||
                function (s) {
                    var matches = (this.document || this.ownerDocument).querySelectorAll(s),
                        i = matches.length;
                    while (--i >= 0 && matches.item(i) !== this) { }
                    return i > -1;
                };
        }

        // Get closest match
        for (; elem && elem !== document; elem = elem.parentNode) {
            if (elem.matches(selector)) return elem;
        }

        return null;
    }
}
