import * as React from 'react';
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import styles from './Styling.module.scss';
import * as strings from 'ClassificationNotificationApplicationCustomizerStrings';
import { Dialog } from '@microsoft/sp-dialog';

export interface IClassificationNotificationProps {
    context: IWebPartContext;
    classifications: string[];
}

export interface IClassificationNotificationState {
    classifications: string;
}
export default class ClassificationNotification extends React.PureComponent<IClassificationNotificationProps, IClassificationNotificationState> {
    constructor(props) {
        super(props);
    }

    public componentDidMount(): void {
    }

    public render(): React.ReactElement<null> {
        if (this.props.classifications.indexOf(this.props.context.pageContext.legacyPageContext.siteClassification) != -1) {
            return (
                <div className={styles.notification}>
                    <Icon iconName='Lock' className={styles.warningIcon} /><span className={styles.warningAdjust}>{strings.Notification} <span className={styles.classificationColor}>{this.props.context.pageContext.legacyPageContext.siteClassification}</span></span>
                </div>
            );
        }
        return <span />;
    }
}
