export interface ISiteListPage {
    page: number;
    sites: ISite[];
}

export interface ISite {
    displayName: string;
    description: string;
    createdDate?: string;
    url?: string;
}

export interface IGraphGroup {
    id: string;
    displayName: string;
    description: string;
}

export interface IGraphSiteResponse {
    pages: ISiteListPage[];
    nextPage: string;
}

export interface IGraphBatchRequestObject {
    requests: IGraphBatchRequest[];
}

export interface IGraphBatchRequest {
    id: string;
    method: string;
    url: string;
}

export interface IGraphSiteResponseBody {
    webUrl: string;
    createdDateTime: string;
}

export interface IGraphBatchResponse {
    id: string;
    body: IGraphSiteResponseBody;
}

export interface IGraphBatchResponseObject {
    responses: IGraphBatchResponse[];
}

export interface IGraphSiteProperties {
    createdDateTime: Date;
    webUrl: string;
}