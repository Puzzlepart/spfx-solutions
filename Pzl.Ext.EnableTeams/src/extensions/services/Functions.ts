import { MSGraph } from '.';
import { GraphHttpClient } from '@microsoft/sp-http';
import { Web, Site } from '@pnp/sp';

export class Functions {
    public static async CreateTeam(graphHttpClient: GraphHttpClient, groupId: string, siteUrl: string) {
        console.log("Creating team");
        let payload: any = {
            "memberSettings": {
                "allowCreateUpdateChannels": true
            },
            "messagingSettings": {
                "allowUserEditMessages": true,
                "allowUserDeleteMessages": true
            },
            "funSettings": {
                "allowGiphy": true,
                "giphyContentRating": "strict"
            }
        };
        await MSGraph.Put(graphHttpClient, `beta/groups/${groupId}/team`, JSON.stringify(payload));
        let teamsUri;
        while (true) {
            let endPointInfo = await MSGraph.Get(graphHttpClient, `beta/groups/${groupId}/endpoints`);
            if (endPointInfo && endPointInfo.value && endPointInfo.value.length > 0) {
                let info = endPointInfo.value.find(element => { return element.providerName === 'Microsoft Teams'; });
                if (info) {
                    console.log("Adding teams link");
                    teamsUri = info.uri;
                    let currentWeb = new Web(siteUrl);
                    await currentWeb.navigation.quicklaunch.add("Teams", info.uri);
                    break;
                }
            } else {
                console.log("Waiting for teams to be ready");
                await this.Timeout(500);
            }
        }
        return teamsUri;
    }

    private static Timeout(ms: number): Promise<any> {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    public static async RemoveCustomizer(siteUrl: string, componentId: string) {
        // Remove custom action from current sute
        let site = new Site(siteUrl);
        let customActions = await site.userCustomActions.get();
        for (let i = 0; i < customActions.length; i++) {
            var instance = customActions[i];
            if (instance.ClientSideComponentId === componentId) {
                await site.userCustomActions.getById(instance.Id).delete();
                console.log("Teams creation extension removed");
                window.location.href = window.location.href;
                break;
            }
        }
    }
}
