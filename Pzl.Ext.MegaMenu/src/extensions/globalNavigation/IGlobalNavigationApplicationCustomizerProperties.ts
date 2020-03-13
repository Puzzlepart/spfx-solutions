import { IGlobalNavigationDataFetchSpListParams, IGlobalNavigationDataFetchJsonParams, IGlobalNavigationDataFetchTaxonomyParams } from "./GlobalNavigationDataFetch";
import { Alignment } from "../TextAlignment";

export default interface IGlobalNavigationApplicationCustomizerProperties {
    dataSource: {
        spList?: IGlobalNavigationDataFetchSpListParams;
        json?: IGlobalNavigationDataFetchJsonParams;
        taxonomy?: IGlobalNavigationDataFetchTaxonomyParams;
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

