import * as React from 'react';
import styles from './RssFeed.module.scss';
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { IRssFeedProps } from './IRssFeedProps';
import { IRssFeedState } from './IRssFeedState';
import { PnPClientStorage } from "@pnp/common";
import * as strings from 'RssFeedWebPartStrings';
import { Text } from 'office-ui-fabric-react/lib/Text';

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
      const storage = new PnPClientStorage();
      let now = new Date();
      let json = await storage.local.getOrPut(`rssfeed-${this.props.context.pageContext.web.serverRelativeUrl}-${this.props.instanceId}`, async () => {
        const apiKey = this.props.apiKey && this.props.apiKey !== undefined ? `&api_key=${this.props.apiKey}` : '';
        const response = await fetch(`https://api.rss2json.com/v1/api.json?rss_url=${encodeURIComponent(this.props.rssFeedUrl)}${apiKey}`);
        return await response.json();
      }, moment(now).add(this.props.cacheDuration, 'm').toDate());
      this.setState({ items: (json.items) ? json.items.splice(0, this.props.itemsCount) : [] });
    } catch (error) {
      throw error;
    }
  }

  public render(): React.ReactElement<IRssFeedProps> {
    return (
      <div className={styles.rssFeed}>
        <div className={styles.container}>
          <div className={styles.webpartHeader}>
            <span>{this.props.title}</span>
            <span className={styles.showAll}>
              <Text onClick={() => window.open(this.props.seeAllUrl, '_blank')} >{strings.SeeAllText}</Text>
            </span>
          </div>
          <ul className={styles.itemsList}>
            {(this.state.items) ? this.state.items.map(({ title, pubDate, link }, index) => (
              <Text className={styles.listItem} onClick={() => window.open(link.replace(/&amp;/g, '&'), '_blank')} key={`listItem_${index}`} >
                <Icon iconName={this.props.officeUIFabricIcon} className={styles.icon} />
                <div>
                  <div className={`${styles.listItemTitle}`}>
                    {title}
                  </div>
                  {pubDate ? <div className={`${styles.listItemPubDate} ms-font-xs`}>{strings.View_PublishLabel} {moment(pubDate).format("DD.MM.YYYY")}</div> : null}
                </div>
              </Text>
            )) : null}
          </ul>
        </div>
      </div>
    );
  }
}
