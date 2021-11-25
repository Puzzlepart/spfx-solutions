import { GlobalNavigationDataFetchBase } from "../GlobalNavigationDataFetch";
import { NavigationSettings } from "../SettingsHelper/SettingsHelper";

export default interface IGlobalNavigationProps extends React.HTMLAttributes<HTMLElement> {
    dataFetch: GlobalNavigationDataFetchBase;
    errorText: string;
    settings: NavigationSettings;
    currentSiteUrl: string;
}
