import styles from './AppCustomizer.module.scss';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseApplicationCustomizer,
    PlaceholderContent,
    PlaceholderName
} from '@microsoft/sp-application-base';
require('mutationobserver-shim');

const LOG_SOURCE: string = '[Office365GroupStatusbarApplicationCustomizer]';

export interface IOffice365GroupStatusbarApplicationCustomizerProperties {
    messageId: string;
}

// Need to keep global scoped variables due to MutationObserver
let _messageId: string;
let _renderContent: () => void;
let _topPlaceholder: PlaceholderContent | undefined;
let _observer: MutationObserver;

/** A Custom Action which can be run during execution of a Client Side Application */
export default class Office365GroupStatusbarApplicationCustomizer
    extends BaseApplicationCustomizer<IOffice365GroupStatusbarApplicationCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        this.context.placeholderProvider.changedEvent.add(this, this.render);
        // Call render method for generating the HTML elements.
        //this.render();
        return Promise.resolve<void>();
    }

    private async render(): Promise<void> {
        if (!_messageId && this.properties.messageId) {
            _messageId = this.properties.messageId;
        }
        if (!_renderContent) {
            _renderContent = this.renderContent;
        }

        if (!window[_messageId]) {
            // Make sure message element is present
            let div = document.createElement("DIV");
            div.id = _messageId;
            div.style.cssText = "display:none";
            document.body.appendChild(div);
        }

        if (!_topPlaceholder) {
            _topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });
        }

        if (!_observer) {
            // Register observable
            let targetNode = window[_messageId];
            let config = { childList: true };

            // Callback function to execute when mutations are observed
            _observer = new MutationObserver(this.callback);
            _observer.observe(targetNode, config);
        }

        _renderContent();
    }

    private renderContent(): void {
        if (!_topPlaceholder) {
            console.error('The expected placeholder (Top) was not found.');
            return;
        }
        console.log('Rendering content');

        if (_topPlaceholder.domElement && window[_messageId] && window[_messageId].innerHTML.length > 0) {
            _topPlaceholder.domElement.innerHTML = `
                <div class="${styles.msgContainer}">
                    <div class="${styles.top} ${styles.message}">
                        ${window[_messageId].innerHTML}
                    </div>
                </div>`;
        } else {
            _topPlaceholder.domElement.innerHTML = '';
        }
    }

    private callback(mutationsList) {
        for (var mutation of mutationsList) {
            if (mutation.type == 'childList') {
                _renderContent();
            }
        }
    }

    private _onDispose(): void {
        if (_observer) {
            _observer.disconnect();
        }
        console.log(LOG_SOURCE + ' Disposed custom top placeholders.');
    }
}
