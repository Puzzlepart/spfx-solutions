import { AadTokenProvider, AadHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import IComment from '../interfaces/IComment';

export interface IYammerService {
    getWebLink(): Promise<any>;
    getWebLinkMessages(id: string): Promise<any>;
    getCommunities(): Promise<any>;
    getCurrentUser(): Promise<any>;
    getUser(id: string): Promise<any>;
    postComment(comment: IComment): Promise<any>;
}

export class YammerService implements IYammerService {

    private readonly api: string = "https://api.yammer.com/api/v1";

    constructor(private tokenProvider: AadTokenProvider, private httpClient: AadHttpClient) { }

    public async getWebLink(): Promise<any> {
        let response = await this.httpClient.get(
            `${this.api}/open_graph_objects?url=${window.location.href}`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            return await response.json();
        } else {
            throw Error(response.statusText);
        }
    }

    public async getWebLinkMessages(id: string): Promise<any> {
        let response = await this.httpClient.get(
            `${this.api}/messages/open_graph_objects/${id}.json?threaded=true`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            return await response.json();
        } else {
            throw Error(response.statusText);
        }
    }

    public async getCommunities(): Promise<any> {

        let user = await this.getCurrentUser();

        let response = await this.httpClient.get(
            `${this.api}/groups/for_user/${user.id}.json`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            return await response.json();
        } else if (response.status === 404) {
            return null;
        } else {
            throw Error(response.statusText);
        }
    }

    public async getUser(id: string): Promise<any> {

        const cachedUser = sessionStorage.getItem(`Yammer.User.${id}`);
        if (cachedUser) {
            return JSON.parse(cachedUser);
        }

        let response = await this.httpClient.get(
            `${this.api}/users/${id}.json`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            const json = await response.json();
            let user = { id: json.id, name: json.full_name, email: json.email };
            sessionStorage.setItem(`Yammer.User.${id}`, JSON.stringify(user));
            return user;
        } else {
            throw Error(response.statusText);
        }
    }

    public async getCurrentUser(): Promise<any> {
        return this.getUser('current');
    }

    public async postComment(comment: IComment): Promise<any> {

        const headers: HeadersInit = new Headers();
        headers.append('content-type', 'application/json');

        // See https://developer.yammer.com/docs/messages-json-post
        let message = !comment.replyToId ? {
            body: comment.text,
            group_id: comment.groupId,
            og_url: window.location.href,
            og_title: document.title,
            // TODO og_description: Lookup using pageContext.listItem.id
            // TODO og_image: `https://${pageContext.site.absoluteUrl}/_layouts/15/getpreview.ashx?path=${pageContext.site.serverRequestPath}`
            // TODO og_site_name: `${pageContext.web.title}`
        } : {
            body: comment.text,
            replied_to_id: comment.replyToId
        };

        let response = await this.httpClient.post(
            `${this.api}/messages.json`,
            AadHttpClient.configurations.v1,
            {
                headers: headers,
                body: JSON.stringify(message)
            }
        );

        if (response.ok) {
            let result = await response.json();
            console.log('reponse: ' + result);
            return result;
        } else {
            throw Error(response.statusText);
        }
    }
}