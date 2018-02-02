import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
    DialogContent
} from 'office-ui-fabric-react';

interface ICreateTeamsDialogContentProps {
    message: string;
}

class CreateTeamsDialogContent extends React.Component<ICreateTeamsDialogContentProps, {}> {
    constructor(props) {
        super(props);
    }

    public render(): JSX.Element {
        return <DialogContent
            title='Creating Microsoft Teams'
            subText={this.props.message}
            showCloseButton={false}
        >
        </DialogContent>;
    }

}

export default class CreateTeamsDialog extends BaseDialog {
    public message: string;
    public colorCode: string;

    public render(): void {
        ReactDOM.render(<CreateTeamsDialogContent
            message={this.message}
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: true
        };
    } 
}