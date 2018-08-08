import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Label } from 'office-ui-fabric-react/lib/Label';
import {
    DialogContent
} from 'office-ui-fabric-react';

interface IWaitDialogContentProps {
    message: string;
    error: string;
    title: string;
}

class WaitDialogContent extends React.Component<IWaitDialogContentProps, {}> {
    constructor(props) {
        super(props);
    }

    public render(): JSX.Element {
        return (<div style={{ width: "400px" }}>
            <DialogContent
                title={this.props.title}
                subText={this.props.message}
                showCloseButton={false}
            >
                <Label>
                    {this.props.error}
                </Label>
            </DialogContent>
        </div>);
    }
}


export default class WaitDialog extends BaseDialog {
    public message: string;
    public title: string;
    public error: string;

    public render(): void {
        ReactDOM.render(<WaitDialogContent
            message={this.message}
            title={this.title}
            error={this.error}
        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: true
        };
    }
}