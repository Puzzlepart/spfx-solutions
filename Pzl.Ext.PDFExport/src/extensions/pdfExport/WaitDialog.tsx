import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DialogContent } from 'office-ui-fabric-react/lib/Dialog';

interface IWaitDialogContentProps {
    message: string;
    error: string;
    title: string;
    showClose: boolean;
    closeCallback: () => void;
}

class WaitDialogContent extends React.Component<IWaitDialogContentProps, {}> {
    constructor(props) {
        super(props);
    }

    public render(): JSX.Element {
        let logo = require('./pzl-logo-black-transparent.png');

        return (<div style={{ width: "400px" }}>
            <DialogContent
                title={this.props.title}
                subText={this.props.message}
                showCloseButton={this.props.showClose}
                onDismiss={this.props.closeCallback}
            >
                <Label>
                    {this.props.error}
                </Label>
                <div style={{ fontSize: '0.8em', float: 'right' }}>
                    <a href="https://www.puzzlepart.com">
                        Powered by
                    <br />
                        <img src={logo} style={{ width: '100px' }} />
                    </a>
                </div>
            </DialogContent>
        </div>);
    }
}


export default class WaitDialog extends BaseDialog {
    public message: string;
    public title: string;
    public error: string;
    public showClose: boolean = false;

    constructor(props) {
        super(props);
        this.closeDialog = this.closeDialog.bind(this);
    }

    public render(): void {
        ReactDOM.render(<WaitDialogContent
            message={this.message}
            title={this.title}
            error={this.error}
            showClose={this.showClose}
            closeCallback={this.closeDialog}
        />, this.domElement);
    }

    private closeDialog() {
        this.close();
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: true,
        };
    }
}