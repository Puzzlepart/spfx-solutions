
import * as React from 'react';
import { ISimpleSearchBoxProps } from './ISimpleSearchBoxProps';
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
    window.open(this.props.searchurl + encodeURIComponent(query), this.props.openInNewTab ? '_blank' : '_parent');
  }
}
