import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { Web, SiteUserProps } from '@pnp/sp';

export interface IMoveEveryoneApplicationCustomizerProperties {
    force: string;
}

export default class MoveEveryoneApplicationCustomizer extends BaseApplicationCustomizer<IMoveEveryoneApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        this.DoWork();
        return Promise.resolve();
    }

    private async DoWork() {
        let isGroupOwner = this.context.pageContext.legacyPageContext.isSiteAdmin;
        if (!isGroupOwner) return;
        let force = this.properties.force.toLowerCase() === 'true';
        let currentWeb = new Web(this.context.pageContext.web.absoluteUrl);
        let memberGroupUsers = await currentWeb.associatedMemberGroup.users.get();
        let siteUsers = await currentWeb.siteUsers.get();
        const everyoneIdent = "c:0-.f|rolemanager|spo-grid-all-users/";
        for (let i = 0; i < siteUsers.length; i++) {
            var user: SiteUserProps = siteUsers[i];
            if (user.LoginName.indexOf(everyoneIdent) === -1) continue;
            if (force) {
                await currentWeb.associatedVisitorGroup.users.add(user.LoginName);
                console.log("Moved everyone to visitors");
            }

            for (var j = 0; j < memberGroupUsers.length; j++) {
                var member = memberGroupUsers[j];
                if (member.LoginName == user.LoginName) {
                    await currentWeb.associatedMemberGroup.users.removeByLoginName(member.LoginName);
                    if (!force) {
                        await currentWeb.associatedVisitorGroup.users.add(member.LoginName);
                        console.log("Moved everyone to visitors");
                    }
                    break;
                }
            }
        }
    }
}
