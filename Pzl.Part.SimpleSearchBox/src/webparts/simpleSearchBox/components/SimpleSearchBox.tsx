
import * as React from 'react';
import styles from './SimpleSearchBox.module.scss';
import { ISimpleSearchBoxProps } from './ISimpleSearchBoxProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SearchBox } from 'office-ui-fabric-react';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

export default class SimpleSearchBox extends React.Component<ISimpleSearchBoxProps, {}> {

  constructor(props) {
    super(props);
  }
  public render() {
    return (
      <div>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty} />
        <SearchBox labelText={this.props.placeholder} onSearch={(query) => this.executeSearch(query)} />
      </div>
    );
  }
  private executeSearch(query) {
    window.location.href = this.props.searchurl + encodeURIComponent(query);
  }
}
