import { MessageBarBase } from '@microsoft/office-ui-fabric-react-bundle';
import { AadTokenProvider, AadHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import IComment from '../interfaces/IComment';
import IUser from '../interfaces/IUser';

export interface IYammerService {
    getWebLink(): Promise<any>;
    getWebLinkMessages(id: string): Promise<string[]>;
    getMessagesInThread(id: string): Promise<IComment[]>;
    getCommunities(): Promise<any>;
    getCurrentUser(): Promise<IUser>;
    getUser(id: string): Promise<IUser>;
    postComment(comment: IComment): Promise<any>;
    buildHierarchy(comments: IComment[]): IComment;
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

    public async getWebLinkMessages(id: string): Promise<string[]> {
        let response = await this.httpClient.get(
            `${this.api}/messages/open_graph_objects/${id}.json?threaded=true`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            const result = await response.json();
            result.references.forEach(reference => {
                if (reference.type === 'user') {
                    let user = { id: reference.id, name: reference.full_name, email: reference.email };
                    sessionStorage.setItem(`Yammer.User.${id}`, JSON.stringify(user));
                }
            });
            return result.messages.map(message => {
                return message.thread_id;
            });
        } else {
            throw Error(response.statusText);
        }
    }

    private async getComment(message) {
        const user = await this.getUser(message.sender_id);
        return {
            id: message.id,
            text: message.body.parsed,
            created: message.created_at,
            groupId: message.group_id,
            replyToId: message.replied_to_id,
            user: user
        };
    }

    public async getMessagesInThread(id: string): Promise<IComment[]> {
        let response = await this.httpClient.get(
            `${this.api}/messages/in_thread/${id}.json`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            let json = await response.json();
            return Promise.all(json.messages.map(message => this.getComment(message)));
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

    public async getUser(id: string): Promise<IUser> {

        const cachedUser = sessionStorage.getItem(`Yammer.User.${id}`);
        if (cachedUser) {
            return JSON.parse(cachedUser);
        }

        let response = await this.httpClient.get(
            `${this.api}/users/${id}.json`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            const json = await response.json();
            let user = { id: json.id, name: json.full_name, email: json.email, url: json.web_url };
            sessionStorage.setItem(`Yammer.User.${id}`, JSON.stringify(user));
            return user;
        } else {
            throw Error(response.statusText);
        }
    }

    public async getCurrentUser(): Promise<IUser> {
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
            og_site_name: window.location.hostname
            // TODO og_description: Lookup using pageContext.listItem.id
            // TODO og_image: `https://${pageContext.site.absoluteUrl}/_layouts/15/getpreview.ashx?path=${pageContext.site.serverRequestPath}`
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

    public buildHierarchy(comments: IComment[]): IComment {

        let rootComment;
        let hashMap = {};

        // Build a map of comments
        //   key = comment.id
        //   value = comment
        comments.forEach(comment => {
            hashMap[comment.id] = comment;
            hashMap[comment.id].replies = [];
        });

        for (var id in hashMap) {
            let comment = hashMap[id];
            if (comment.replyToId) {
                // The comment is a reply to another comment, add it as a child to it's parent
                hashMap[comment.replyToId].replies.push(comment);
            } else {
                // This is the root comment
                rootComment = comment;
            }
        }
        return rootComment;
    }
}