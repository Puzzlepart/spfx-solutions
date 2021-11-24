import { IGlobalNavigationDataFetchSpListParams, IGlobalNavigationDataFetchJsonParams } from "./GlobalNavigationDataFetch";
import { Alignment } from "../TextAlignment";

export default interface IGlobalNavigationApplicationCustomizerProperties {
    dataSource: {
        spList?: IGlobalNavigationDataFetchSpListParams;
        json?: IGlobalNavigationDataFetchJsonParams;
    };
    serviceAnnouncements?: {
        serverRelativeWebUrl: string;
        listUrl: string;
        settingsListUrl: string;
        discardForSessionOnly: boolean;
        textAlignment: Alignment;
        boldText: boolean;
    };
}

