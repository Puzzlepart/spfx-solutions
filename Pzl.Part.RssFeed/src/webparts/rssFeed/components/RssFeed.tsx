import * as React from 'react';
import styles from './RssFeed.module.scss';
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IRssFeedProps } from './IRssFeedProps';
import { IRssFeedState } from './IRssFeedState';
import * as pnp from 'sp-pnp-js';
import * as strings from 'RssFeedWebPartStrings';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';

export default class RssFeed extends React.Component<IRssFeedProps, IRssFeedState> {
  constructor(props: IRssFeedProps) {
    super(props);
    this.state = { items: [] };
  }

  public componentDidMount() {
    this.fetchData();
  }

  public componentWillUpdate(nextProps) {
    if (nextProps != this.props) {
      this.fetchData();
    }
  }

  private async fetchData() {
    try {
      let json = await pnp.storage.local.getOrPut(`rssfeed-${this.props.context.pageContext.web.serverRelativeUrl}-${this.props.instanceId}`, async () => {
        const response = await fetch(`https://api.rss2json.com/v1/api.json?rss_url=${this.props.rssFeedUrl}&api_key=${this.props.apiKey}`);
        return await response.json();
      }, pnp.util.dateAdd(new Date(), "minute", this.props.cacheDuration));
      this.setState({ items: (json.items) ? json.items.splice(0, this.props.itemsCount) : [] });
    } catch (error) {
      throw error;
    }
  }

  public render(): React.ReactElement<IRssFeedProps> {
    return (
      <div className={styles.rssFeed}>
        <div className={styles.container}>
          <div className={styles.headerContainer}>
            <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} />
            {this.props.seeAllUrl && <span className={styles.showAll}><a href={this.props.seeAllUrl}>{strings.SeeAllText}</a></span>}
          </div>
          <ul className={styles.itemsList}>
            {(this.state.items) ? this.state.items.map(({ title, pubDate, link }, index) => (
              <a target="_blank" key={`listItem_${index}`} className={styles.listItem} href={link.replace(/&amp;/g, '&')}>
                <Icon iconName={this.props.officeUIFabricIcon} className={styles.icon} />
                <div>
                  <div className={`${styles.listItemTitle}`}>
                    {title}
                  </div>
                  {pubDate ? <div className={`${styles.listItemPubDate} ms-font-xs`}>{strings.View_PublishLabel} {moment(pubDate).format("DD.MM.YYYY")}</div> : null}
                </div>
              </a>
            )) : null}
          </ul>
        </div>
      </div>
    );
  }
}
