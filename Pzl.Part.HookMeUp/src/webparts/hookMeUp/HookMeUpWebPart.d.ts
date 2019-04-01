import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface IHookMeUpWebPartProps {
    anchorId: string;
}
export default class HookMeUpWebPart extends BaseClientSideWebPart<IHookMeUpWebPartProps> {
    render(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
