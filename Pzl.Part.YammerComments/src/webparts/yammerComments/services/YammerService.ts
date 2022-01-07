import { AadTokenProvider, AadHttpClient } from '@microsoft/sp-http';

export interface IYammerService {
    getWebLink(): Promise<any>;
}

export default class YammerService implements IYammerService {

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
}