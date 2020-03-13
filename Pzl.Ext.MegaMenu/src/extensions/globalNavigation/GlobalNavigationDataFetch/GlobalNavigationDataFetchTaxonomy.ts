import { PnPClientStorage, dateAdd } from "@pnp/common";
import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from "spfx-jsom";
import { INavigationElementProps } from '../GlobalNavigation/NavigationElement';
import GlobalNavigationDataFetchBase from "./GlobalNavigationDataFetchBase";
import { override } from "@microsoft/decorators";

export interface IGlobalNavigationDataFetchTaxonomyParams {
    termGroupId: string;
    termSetSortProperty?: string;
    storageConfig: {
        key: string;
        expirationMinutes: number;
    };
}

export default class GlobalNavigationDataFetchTaxonomy extends GlobalNavigationDataFetchBase {
    private _params: IGlobalNavigationDataFetchTaxonomyParams;
    private _siteUrl: string;
    private _storage: PnPClientStorage;

    /**
     * Constructor
     * 
     * @param {IGlobalNavigationDataFetchTaxonomyParams} params Parameters
     * @param {string} siteUrl Site URL
     */
    constructor(params: IGlobalNavigationDataFetchTaxonomyParams, siteUrl: string) {
        super();
        this._params = params;
        this._siteUrl = siteUrl;
        this._storage = new PnPClientStorage();
    }

    /**
     * Override fetch() in GlobalNavigationDataFetchBase
     */
    @override
    public async fetch(): Promise<INavigationElementProps[]> {
        const { termGroupId, termSetSortProperty, storageConfig } = this._params;

        const jsomContext: JsomContext = await initSpfxJsom(this._siteUrl, { loadTaxonomy: true, loadPublishing: true });
        const taxSession = SP.Taxonomy.TaxonomySession.getTaxonomySession(jsomContext.clientContext);
        const termStore = taxSession.getDefaultSiteCollectionTermStore();

        let navigationElements: INavigationElementProps[] = await this._storage.local.getOrPut(storageConfig.key, async () => {
            let items: INavigationElementProps[] = [];
            const termGroup = termStore.getGroup(new SP.Guid(termGroupId));
            const termSets = termGroup.get_termSets();
            jsomContext.clientContext.load(termSets);
            await ExecuteJsomQuery(jsomContext);
            let terms = [];
            for (let i = 0; i < termSets.get_count(); i++) {
                terms.push(SP.Publishing.Navigation.NavigationTermSet.getAsResolvedByWeb(jsomContext.clientContext, termSets.get_item(i), jsomContext.clientContext.get_web(), "GlobalNavigationTaxonomyProvider").getAllTerms());
                jsomContext.clientContext.load(terms[i], 'Include(Title,SimpleLinkUrl)');
            }
            await ExecuteJsomQuery(jsomContext);
            for (let i = 0; i < terms.length; i++) {
                let order = i;
                const ts = termSets.get_item(i);
                if (termSetSortProperty) {
                    let property = ts.get_customProperties()[termSetSortProperty];
                    if (property) {
                        order = parseInt(property, 10);
                    }
                }
                let item: INavigationElementProps = {
                    header: ts.get_name(),
                    links: [],
                    order,
                };
                item.links = terms[i].get_data().map(t => ({
                    text: t.get_title().get_value(),
                    url: t.get_simpleLinkUrl(),
                }));
                items.push(item);
            }
            return items;
        }, dateAdd(new Date(), "minute", storageConfig.expirationMinutes));
        return navigationElements;
    }
}