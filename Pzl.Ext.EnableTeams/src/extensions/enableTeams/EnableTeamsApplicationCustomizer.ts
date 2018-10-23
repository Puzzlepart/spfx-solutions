import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { MSGraph, Functions } from '../services';
import TeamsButton from './TeamsButton';
import CreateTeamsDialog from './CreateTeamsDialog';
import './AppCustomizer.scss';

const LOG_SOURCE: string = 'EnableTeamsApplicationCustomizer';

export interface IEnableTeamsApplicationCustomizerProperties {
    autoCreate: boolean;
    shouldRedirect: boolean;
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class EnableTeamsApplicationCustomizer
    extends BaseApplicationCustomizer<IEnableTeamsApplicationCustomizerProperties> {

    private _topPlaceholder: PlaceholderContent | undefined;

    @override
    public onInit(): Promise<void> {
        this.arrayPolyfill();
        if (typeof console == "undefined" || typeof console.log == "undefined") var console = { log: function () { } };
        // Extra check for siteadmin to ensure it's run by a Group owner
        let isSiteAdmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
        if (isSiteAdmin) {
            let autoCreate = false;
            if (typeof this.properties.autoCreate !== 'undefined') {
                autoCreate = this.properties.autoCreate.toString().toLowerCase() === 'true';
            }
            let shouldRedirect = false;
            if (typeof this.properties.shouldRedirect !== 'undefined') {
                shouldRedirect = this.properties.shouldRedirect.toString().toLowerCase() === 'true';
            }

            this.context.placeholderProvider.changedEvent.add(this, () => { this.DoWork(autoCreate, shouldRedirect); });
        }
        return Promise.resolve();
    }

    private async DoWork(autoCreate: boolean, shouldRedirect: boolean, ) {
        let failed = false;
        let teamsUri;
        let hasTeam = false;
        let groupId = this.context.pageContext.legacyPageContext.groupId;

        let endPointInfo = await MSGraph.Get(this.context.graphHttpClient, `beta/groups/${groupId}/endpoints`);
        if (endPointInfo && endPointInfo.value) {
            let info = endPointInfo.value.find(element => { return element.providerName === 'Microsoft Teams'; });
            hasTeam = info != null;
        }

        const dialog: CreateTeamsDialog = new CreateTeamsDialog();
        try {
            if (!hasTeam && autoCreate) {
                dialog.message = "Please wait while we set up Microsoft Teams and add a navigation link...";
                dialog.show();
                teamsUri = await Functions.CreateTeam(this.context.graphHttpClient, groupId, this.context.pageContext.site.absoluteUrl);
                hasTeam = true;
            }
        } catch (error) {
            failed = true;
            Log.error(LOG_SOURCE, error);
        } finally {
            dialog.close();
        }

        if (!failed && hasTeam) {
            await Functions.RemoveCustomizer(this.context.pageContext.site.absoluteUrl, this.componentId);
            if (shouldRedirect) {
                document.location.href = teamsUri;
            }
        }
        if (!this._topPlaceholder && !autoCreate) {
            this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {});
            if (!this._topPlaceholder || !this._topPlaceholder.domElement) {
                Log.info(LOG_SOURCE, 'The expected placeholder (Top) was not found.');
                return;
            }
            Log.info(LOG_SOURCE, 'The expected placeholder (Top) was found. Rendering <NavigationContainer />');

            let buttonProps = { graphClient: this.context.graphHttpClient, groupId: groupId, siteUrl: this.context.pageContext.site.absoluteUrl, componentId: this.componentId, shouldRedirect: shouldRedirect };
            ReactDOM.render(React.createElement(TeamsButton, buttonProps, null), this._topPlaceholder.domElement);
        }

    }

    private arrayPolyfill() {
        if (!(<any>Array.prototype).find) {
            Object.defineProperty(Array.prototype, 'find', {
                value: function (predicate) {
                    // 1. Let O be ? ToObject(this value).
                    if (this == null) {
                        throw new TypeError('"this" is null or not defined');
                    }

                    var o = Object(this);

                    // 2. Let len be ? ToLength(? Get(O, "length")).
                    var len = o.length >>> 0;

                    // 3. If IsCallable(predicate) is false, throw a TypeError exception.
                    if (typeof predicate !== 'function') {
                        throw new TypeError('predicate must be a function');
                    }

                    // 4. If thisArg was supplied, let T be thisArg; else let T be undefined.
                    var thisArg = arguments[1];

                    // 5. Let k be 0.
                    var k = 0;

                    // 6. Repeat, while k < len
                    while (k < len) {
                        // a. Let Pk be ! ToString(k).
                        // b. Let kValue be ? Get(O, Pk).
                        // c. Let testResult be ToBoolean(? Call(predicate, T, « kValue, k, O »)).
                        // d. If testResult is true, return kValue.
                        var kValue = o[k];
                        if (predicate.call(thisArg, kValue, k, o)) {
                            return kValue;
                        }
                        // e. Increase k by 1.
                        k++;
                    }

                    // 7. Return undefined.
                    return undefined;
                },
                configurable: true,
                writable: true
            });
        }
    }
}
