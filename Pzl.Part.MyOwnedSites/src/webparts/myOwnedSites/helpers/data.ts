import { IGraphBatchRequestObject, IGraphBatchResponseObject, IGraphGroup, IGraphSiteProperties, ISiteResponse, ISite, ISiteListPage } from "../models/types";
import { MSGraphClientV3 } from '@microsoft/sp-http'
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/search";
import { SearchResults, SearchQueryBuilder, ISearchResult } from "@pnp/sp/search";

export const getOwnedGroupSites = async (context: WebPartContext, loadedPages: ISiteListPage[], nextPage: string | undefined): Promise<ISiteResponse> => {
    const graphClient = await context.msGraphClientFactory.getClient('3') as MSGraphClientV3
    const url = nextPage ?? "https://graph.microsoft.com/v1.0/me/ownedObjects/microsoft.graph.group?$select=mailNickname,id,displayName,description&$top=10";
    const res = await graphClient.api(url).get()

    const nextPageUrl = res["@odata.nextLink"];
    const results: unknown[] = res.value;
    const batchRequest: IGraphBatchRequestObject = {
        requests: []
    }
    const groupsMap = new Map<string, ISite>();

    results.forEach((group: IGraphGroup) => {
        groupsMap.set(group.id, { displayName: group.displayName, description: group.description });

        batchRequest.requests.push({
            id: group.id,
            method: "GET",
            url: `/groups/${group.id}/sites/root?$select=createdDateTime,webUrl`
        });
    });

    const batchResponse: IGraphBatchResponseObject = await graphClient.api("https://graph.microsoft.com/v1.0/$batch").post(batchRequest);

    const previousPage = loadedPages[loadedPages.length - 1]?.page || 0;
    const page: ISiteListPage = { page: previousPage + 1, sites: [] };

    batchResponse.responses.forEach(res => {
        const group = groupsMap.get(res.id);
        if (!group) {
            return;
        }

        const siteProperties: IGraphSiteProperties = {
            createdDateTime: new Date(res.body.createdDateTime),
            webUrl: res.body.webUrl
        };

        group.url = siteProperties.webUrl;
        group.createdDate = siteProperties.createdDateTime.toLocaleDateString('nb-NO');
        page.sites.push(group);
    });

    return { nextPage: nextPageUrl, pages: [...loadedPages, page] };
};

export const getCreatedSites = async (client: SPFI, user: string, loadedPages: ISiteListPage[], startRow: number): Promise<ISiteResponse> => {
    console.log(startRow);
    const previousPage = loadedPages[loadedPages.length - 1]?.page || 0;
    const page: ISiteListPage = { page: previousPage + 1, sites: [] };

    const q = SearchQueryBuilder().text(`* contentclass:sts_site People:${user}`)
    .selectProperties('Title,SPWebUrl,Description,Created')
    .rowsPerPage(10).startRow(startRow);
    const result: SearchResults = await client.search(q);
    const sites: ISite[] = result.PrimarySearchResults.map((site: ISearchResult) => {
        return {
            displayName: site.Title,
            description: site.Description,
            //eslint-disable-next-line @typescript-eslint/no-explicit-any
            createdDate: new Date((site as any).Created).toLocaleDateString('nb-NO'),
            url: site.SPWebUrl
        };
    });

    page.sites = sites;

    return { pages: [...loadedPages, page] };
};