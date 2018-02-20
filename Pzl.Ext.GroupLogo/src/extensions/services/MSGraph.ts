import { GraphHttpClient, GraphHttpClientResponse } from '@microsoft/sp-http';
import { Log } from '@microsoft/sp-core-library';

export class MSGraph {
    public static async Patch(graphClient: GraphHttpClient, url: string, payload: any): Promise<boolean> {
        const requestHeaders: Headers = new Headers();
        requestHeaders.append("Content-Type", "image/jpeg");
        
        let response: GraphHttpClientResponse = await graphClient.fetch(url, GraphHttpClient.configurations.v1, {
            body: payload,
            method: "PATCH",
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
