import { AadTokenProvider, AadHttpClient } from '@microsoft/sp-http';

export interface IYammerService {
    getWebLink(): Promise<any>;
    getCommunities():Promise <any>;
    getCurrentUser():Promise <any>;
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

    public async getCommunities():Promise <any> {
        
        let user = await this.getCurrentUser();

        let response = await this.httpClient.get(
            `${this.api}/groups/for_user/${user.id}.json`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            return await response.json();
        } else {
            throw Error(response.statusText);
        }
    }

    public async getCurrentUser():Promise <any> {
        let response = await this.httpClient.get(
            `${this.api}/users/current.json`,
            AadHttpClient.configurations.v1);
        if (response.ok) {
            return await response.json();
        } else {
            throw Error(response.statusText);
        }
    }
}