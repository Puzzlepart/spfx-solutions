import { Guid, ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import '@pnp/sp/items/list';
import { ICamlQuery } from "@pnp/sp/lists";
import { ClientsidePageFromFile } from "@pnp/sp/clientside-pages";

export interface IPageService {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    getPage(): Promise<any>;
    savePage(): Promise<void>;
}

export class PageService {
    public static readonly serviceKey: ServiceKey<IPageService> = ServiceKey.create<IPageService>("PageService", PageService);
    private _listId: Guid;
    private _itemId: number;
    private _sp: SPFI;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            try {
                const pageContext = serviceScope.consume(PageContext.serviceKey);
                this._listId = pageContext.list.id;
                this._itemId = pageContext.listItem?.id;
                this._sp = spfi().using(SPFx({ pageContext }));
            } catch (e) {
                // We don't have the pageContext.listItem until the page is saved for the first time
            }
        });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    public async getPage(): Promise<any> {
        const query: ICamlQuery = {
            ViewXml: `
            <View>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='ID' />
                            <Value Type='Number'>${this._itemId}</Value>
                        </Eq>
                    </Where>
                </Query>
            </View>
            `
        }
        const [item] = await this._sp.web.lists.getById(this._listId.toString()).getItemsByCAMLQuery(query);
        return item;
    }

    public async savePage(): Promise<void> {

        const page = await ClientsidePageFromFile(
            this._sp.web.getFileByServerRelativePath(window.location.pathname));
        await page.save(true);
    }
}