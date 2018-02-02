import { GraphHttpClient, GraphHttpClientResponse } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';

export class MSGraph {
    public static async Get(graphClient: GraphHttpClient, url: string) {
        let response: GraphHttpClientResponse = await graphClient.get(url, GraphHttpClient.configurations.v1);
        // Check that the request was successful
        if (response.ok) {
            return await response.json();
        }
        else {
            // Reject with the error message
            let error = new Error(response.statusText);
            Log.error("Graph call", error);
            throw error;
        }
    }

    public static async Put(graphClient: GraphHttpClient, url: string, payload: any): Promise<boolean> {
        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-Type", "application/json");

        let response: GraphHttpClientResponse = await graphClient.fetch(url, GraphHttpClient.configurations.v1, {
            body: payload,
            method: "PUT",
            headers: requestHeaders
        });
        // Check that the request was successful
        if (response.ok) {
            return true;
        }
        else {
            // Reject with the error message
            let error = new Error(response.statusText);
            Log.error("Graph call", error);
            throw error;
        }
    }
}
