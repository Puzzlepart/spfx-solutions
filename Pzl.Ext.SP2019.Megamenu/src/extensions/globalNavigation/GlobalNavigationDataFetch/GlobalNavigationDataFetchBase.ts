import { INavigationElementProps } from '../GlobalNavigation/NavigationElement';

export default class GlobalNavigationDataFetchBase {
    public async fetch(): Promise<INavigationElementProps[]> {
        return [];
    }
}