import GlobalNavigationDataFetchBase from "./GlobalNavigationDataFetchBase";
import { override } from "@microsoft/decorators";

export interface IGlobalNavigationDataFetchJsonParams {
    jsonPath: string;
}

export default class GlobalNavigationDataFetchJson extends GlobalNavigationDataFetchBase {
    public _params: IGlobalNavigationDataFetchJsonParams;

    /**
     * Constructor
     * 
     * @param {IGlobalNavigationDataFetchJsonParams} params Parameters
     */
    constructor(params: IGlobalNavigationDataFetchJsonParams) {
        super();
        this._params = params;
    }

    /**
     * Override fetch() in GlobalNavigationDataFetchBase
     */
    @override
    public async fetch() {
        try {
            const response = await fetch(this._params.jsonPath, { credentials: "include" });
            const json = await response.json();
            return json;
        } catch (err) {
            throw err;
        }
    }
}