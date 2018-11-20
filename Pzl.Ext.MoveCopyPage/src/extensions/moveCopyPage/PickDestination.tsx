import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export interface IPickDestinationProps {
    //getObject(done: (data: any, elapsedTime: number) => void): void;
    start(targetSite: string): void;
    fileUrls: string[];
}

export interface IPickDestinationState {
    targetSite: string;
}


export default class PickDestination extends React.Component<IPickDestinationProps, IPickDestinationState> {
    constructor(props) {
        super(props);
        this.start = this.start.bind(this);
    }

    public render(): JSX.Element {
        return (<Panel
            isOpen={true}
            type={PanelType.smallFixedFar}
            headerText="Pick destination site"
            closeButtonAriaLabel="Close"
        >
            <PrimaryButton onClick={this.start}>Click me</PrimaryButton>

        </Panel>);
    }

    private start() {
        this.props.start(this.state.targetSite);
    }

    public updateLog(){

    }
}