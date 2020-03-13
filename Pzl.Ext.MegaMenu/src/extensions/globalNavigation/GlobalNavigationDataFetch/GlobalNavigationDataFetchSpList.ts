import { Web, List } from "@pnp/sp";
import * as urljoin from 'url-join';
import GlobalNavigationDataFetchBase from "./GlobalNavigationDataFetchBase";
import { INavigationElementProps } from '../GlobalNavigation/NavigationElement';
import { INavigationLinkProps } from "../GlobalNavigation/NavigationElement/NavigationLink";
import { override } from "@microsoft/decorators";
//http://www.dotnetmafia.com/blogs/dotnettipoftheday/archive/2018/10/01/typeerror-object-doesn-t-support-property-or-method-from-with-pnpjs-and-internet-explorer-11.aspx
//https://github.com/pnp/pnpjs/issues/237
import "core-js/modules/es6.promise";
import "core-js/modules/es6.array.iterator.js";
import "core-js/modules/es6.array.from.js";
import "whatwg-fetch";
import "es6-map/implement";

export interface IGlobalNavigationDataFetchSpListParams {
    serverRelativeWebUrl: string;
    linksListUrl: string;
    settingsListUrl: string;
    headersListLookupFieldName: string;
    searchQueryFieldName: string;
    urlFieldName: string;
    headersOrderFieldName: string;
    hasHeaderNavLinks: boolean;
}

export default class GlobalNavigationDataFetchSpList extends GlobalNavigationDataFetchBase {
    private _params: IGlobalNavigationDataFetchSpListParams;
    private _spWeb: Web;
    private _linksList: List;

    /**
     * Constructor
     *
     * @param {IGlobalNavigationDataFetchSpListParams} params Parameters
     */
    constructor(params: IGlobalNavigationDataFetchSpListParams) {
        super();
        this._params = params;
        var webUrl = urljoin(`${document.location.protocol}//${document.location.hostname}`, this._params.serverRelativeWebUrl);
        var listUrl = urljoin(this._params.serverRelativeWebUrl, this._params.linksListUrl);
        this._spWeb = new Web(webUrl);
        this._linksList = this._spWeb.getList(listUrl);
    }

    /**
     * Override fetch() in GlobalNavigationDataFetchBase
     */
    @override
    public async fetch() {
        let fields: any = ["Title", this._params.urlFieldName, this._params.headersOrderFieldName,
            `${this._params.headersListLookupFieldName}/Title`,
            `${this._params.headersListLookupFieldName}/${this._params.headersOrderFieldName}`];
        if (this._params.hasHeaderNavLinks) {
            fields.push(`${this._params.headersListLookupFieldName}/${this._params.urlFieldName}`);
        }

        const linksListItems = await this._linksList.items.select(fields).expand(this._params.headersListLookupFieldName).orderBy(`${this._params.headersListLookupFieldName}/${this._params.headersOrderFieldName},${this._params.headersListLookupFieldName}/Title,${this._params.headersOrderFieldName},Title`).usingCaching().get();
        const navigationElements = linksListItems.reduce((arr, linkItem) => {
            const navLink: INavigationLinkProps = {
                text: linkItem.Title,
                url: linkItem[this._params.urlFieldName],
            };
            const headerItemLookup = linkItem[this._params.headersListLookupFieldName];
            if (!headerItemLookup) {
                return arr;
            }
            const [element] = arr.filter(({ header }) => header === headerItemLookup.Title);
            if (element) {
                element.links.push(navLink);
            } else {
                let headerItem: INavigationElementProps = { header: headerItemLookup.Title, order: headerItemLookup.PzlNavOrder, links: [navLink] };
                if (this._params.hasHeaderNavLinks) {
                    headerItem.headerLink = headerItemLookup[this._params.urlFieldName];
                }
                arr.push(headerItem);
            }
            return arr;
        }, []);
        return navigationElements;
    }
}
