import styles from './AppCustomizer.module.scss';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer,
    PlaceholderContent,
    PlaceholderName
} from '@microsoft/sp-application-base';
import { loadStyles } from '@microsoft/load-themed-styles';
require('mutationobserver-shim');

const LOG_SOURCE: string = '[ModernSiteStatusbarApplicationCustomizer]';

export interface IOffice365GroupStatusbarApplicationCustomizerProperties {
    messageId: string;
}

// Need to keep global scoped variables due to MutationObserver
let _topPlaceholder: PlaceholderContent | undefined;
let _observer: MutationObserver;

/** A Custom Action which can be run during execution of a Client Side Application */
export default class Office365GroupStatusbarApplicationCustomizer
    extends BaseApplicationCustomizer<IOffice365GroupStatusbarApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        this.context.placeholderProvider.changedEvent.add(this, this.render);
        return Promise.resolve<void>();
    }

    private async render(): Promise<void> {
        // Make sure all messages are shown on one line
        loadStyles(`#${this.properties.messageId} > DIV {float:left;margin-right:15px;}`);

        let targetNode = document.getElementById(this.properties.messageId);
        if (!targetNode) {
            // Make sure message element is present - if not create it
            targetNode = document.createElement("DIV");
            targetNode.id = this.properties.messageId;
            document.body.appendChild(targetNode);
        }

        if (!_topPlaceholder) {
            _topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });

            // Hide placeholder if there is no text in it
            _topPlaceholder.domElement.style.cssText = targetNode.innerText.length > 0 ? "" : "display:none"

            _topPlaceholder.domElement.innerHTML = `
                <div class="${styles.msgContainer}">
                    <div class="${styles.top} ${styles.message}" id="PzlOuterContainer">
                    </div>
                </div>`;
            let outerContainer = document.getElementById("PzlOuterContainer");
            outerContainer.appendChild(targetNode);
            targetNode.style.cssText = ""; // show node in case it was created by another extension
        }

        if (!_observer) {
            // Register observable
            let config = { childList: true, subtree: true, attributes: false, characterData: true };

            // Callback function to execute when mutations are observed
            console.log("Hooking observer for " + this.properties.messageId);
            _observer = new MutationObserver(this.callback);
            _observer.observe(_topPlaceholder.domElement, config);
        }
    }

    private callback(mutationsList) {
        for (let mutation of mutationsList) {
            let record: MutationRecord = mutation;
            let element: HTMLElement = record.target as HTMLElement;

            _topPlaceholder.domElement.style.cssText = element.innerText.length > 0 ? "" : "display:none"
            break;
        }
        console.log("Status changed");
    }

    private _onDispose(): void {
        if (_observer) {
            _observer.disconnect();
        }
        console.log(LOG_SOURCE + ' Disposed custom top placeholders.');
    }
}
