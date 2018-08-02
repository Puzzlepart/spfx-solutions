import { override } from '@microsoft/decorators';
import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from '@microsoft/sp-listview-extensibility';
import { Web, sp } from "@pnp/sp";

import styles from './HighLightHomeFieldCustomizer.module.scss';

export interface IHighLightHomeFieldCustomizerProperties {
}

let WELCOME_PAGE: string = '';

export default class HighLightHomeFieldCustomizer
    extends BaseFieldCustomizer<IHighLightHomeFieldCustomizerProperties> {

    @override
    public async onInit(): Promise<void> {
        sp.configure({
            headers: {
                'Accept': 'application/json;odata=minimalmetadata;charset=utf-8',
                'Content-Type': 'application/json;odata=minimalmetadata'
            }
        });

        let web = new Web(this.context.pageContext.web.absoluteUrl);

        let welcomePage = await web.rootFolder.select("welcomepage").get();
        WELCOME_PAGE = welcomePage.WelcomePage.replace("SitePages/", "");
    }

    @override
    public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
        event.domElement.innerText = event.fieldValue;
        if (WELCOME_PAGE === event.fieldValue) {
            event.domElement.classList.add(styles.homeIcon);
        }
    }

    @override
    public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
        super.onDisposeCell(event);
    }
}
