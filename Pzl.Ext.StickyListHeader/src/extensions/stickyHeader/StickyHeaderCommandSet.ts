import { override } from '@microsoft/decorators';
import {
    BaseListViewCommandSet,
} from '@microsoft/sp-listview-extensibility';

export interface IStickyHeaderCommandSetProperties {
}

const LOG_SOURCE: string = 'StickyHeaderCommandSet';

export default class StickyHeaderCommandSet extends BaseListViewCommandSet<IStickyHeaderCommandSetProperties> {

    @override
    public onInit(): Promise<void> {
        let listHeader: HTMLDivElement = document.querySelector("div[role='grid'] div") as HTMLDivElement;
        if (listHeader && listHeader.style.position === "") {
            listHeader.style.position = "sticky";
            listHeader.style.zIndex = "1000";
            listHeader.style.top = "0px";
            if (console.log) {
                console.log("Making list header sticky courtesy of Puzzlepart");
            }
        }
        return Promise.resolve();
    }
}
