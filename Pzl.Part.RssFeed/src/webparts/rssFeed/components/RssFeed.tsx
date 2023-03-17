import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './RssFeed.module.scss';
import * as moment from 'moment';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PnPClientStorage } from "@pnp/common";
import * as strings from 'RssFeedWebPartStrings';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRssFeedEnclosure {
  link: string;
}
export interface IRssFeedItem {
  title: string;
  pubDate: string;
  link: string;
  description: string;
  enclosure: IRssFeedEnclosure
}

export interface IRssFeedProps {
  title: string;
  seeAllUrl: string;
  rssFeedUrl: string;
  apiKey: string;
  itemsCount: number;
  officeUIFabricIcon: string;
  showItemDescription: boolean;
  showItemPubDate: boolean;
  showItemImage: boolean;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateProperty: (value: string) => void;
  cacheDuration: number;
  instanceId: string;
}

export const RssFeed: React.FunctionComponent<IRssFeedProps> = (props) => {

  const [items, setItems] = useState<IRssFeedItem[]>([]);

  useEffect(() => {

    const fetchData = async () => {
      const storage = new PnPClientStorage();
      const now = new Date();
      const cacheKey = `rssfeed-${props.context.pageContext.web.serverRelativeUrl}-${props.instanceId}`;
      const apiKey = props.apiKey && props.apiKey !== undefined ? `&api_key=${props.apiKey}` : '';
      const rss2jsonUrl = `https://api.rss2json.com/v1/api.json?rss_url=${encodeURIComponent(props.rssFeedUrl)}${apiKey}`;

      const json = await storage.local.getOrPut(cacheKey, async () => {
        const response = await fetch(rss2jsonUrl);
        return await response.json();
      }, moment(now).add(props.cacheDuration, 'm').toDate());

      setItems((json.items) ? json.items.splice(0, props.itemsCount) : []);
    };

    fetchData();

  }, []);

  /**
   * http://stackoverflow.com/a/10997390/11236
   */
  const updateURLParameter = (url: string, param: string, paramVal: string) => {
    let newAdditionalURL = "";
    let tempArray = url.split("?");
    const baseURL = tempArray[0];
    const additionalURL = tempArray[1];
    let temp = "";
    if (additionalURL) {
      tempArray = additionalURL.split("&");
      for (let i = 0; i < tempArray.length; i++) {
        if (tempArray[i].split('=')[0] != param) {
          newAdditionalURL += temp + tempArray[i];
          temp = "&";
        }
      }
    }

    const rows_txt = temp + "" + param + "=" + paramVal;
    return baseURL + "?" + newAdditionalURL + rows_txt;
  }

  const getThumbnail = (item: IRssFeedItem) => {
    const url = item.enclosure.link.replace(/&amp;/g, '&');

    let updatedUrl = updateURLParameter(url, 'w', '200');
    updatedUrl = updateURLParameter(updatedUrl, 'h', '140');

    return updatedUrl;
  }

  return (
    <div className={styles.rssFeed}>
      <div className={styles.container}>
        {props.title || props.seeAllUrl ? <div className={styles.webpartHeader}>
          {props.title ? <span>{props.title}</span> : ''}
          <span className={styles.showAll}>
            {props.seeAllUrl ? <Text onClick={() => window.open(props.seeAllUrl, '_blank')} >{strings.SeeAllText}</Text> : ''}
          </span>
        </div> : ''}
        <ul className={styles.itemsList}>
          {(items) ? items.map((item, index) => (
            <Text className={styles.listItem} onClick={() => window.open(item.link.replace(/&amp;/g, '&'), '_blank')} key={`listItem_${index}`} title={item.link}>
              {props.officeUIFabricIcon ? <Icon iconName={props.officeUIFabricIcon} className={styles.icon} /> : ''}
              <div className={item.enclosure && item.enclosure.link && props.showItemImage ? styles.contentWithImage : styles.contentNoImage}>
                <div className={styles.containerText}>
                  <div className={`${styles.listItemTitle}`}>
                    {item.title}
                  </div>
                    {item.pubDate && props.showItemPubDate ? <div className={`${styles.listItemPubDate}`}>{moment(item.pubDate).format("DD.MM.YYYY")}</div> : ''}
                    {item.description && props.showItemDescription ? <div className={`${styles.listItemDescription}`} dangerouslySetInnerHTML={{ __html: item.description }} /> : ''}
                </div>
                  {item.enclosure && item.enclosure.link && props.showItemImage ? <div className={styles.image}><img src={getThumbnail(item)} alt={item.title} /></div> : ''}
              </div>
            </Text>
          )) : null}
        </ul>
      </div>
    </div>
  );
};
