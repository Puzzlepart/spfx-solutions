import { IGraphBatchRequestObject, IGraphBatchResponseObject, IGraphGroup, IGraphSiteProperties, IGraphSiteResponse, ISite, ISiteListPage } from "../models/types";
import { MSGraphClientV3 } from '@microsoft/sp-http'
import { WebPartContext } from "@microsoft/sp-webpart-base";

export const getOwnedGroupSites = async (context: WebPartContext, loadedPages: ISiteListPage[], nextPage: string | undefined): Promise<IGraphSiteResponse> => {
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
        group.createdDate = siteProperties.createdDateTime.toLocaleDateString();
        page.sites.push(group);
    });

    return { nextPage: nextPageUrl, pages: [...loadedPages, page] };
};