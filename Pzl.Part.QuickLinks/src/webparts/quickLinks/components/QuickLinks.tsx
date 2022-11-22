import "@pnp/polyfill-ie11";
import * as React from 'react';
import * as strings from 'QuickLinksWebPartStrings';
import styles from './QuickLinks.module.scss';
import { sp } from "@pnp/sp";
import { IQuickLinksProps } from './IQuickLinksProps';
import { find, isEqual } from '@microsoft/sp-lodash-subset';
import { IQuickLinksState, ILink, ICategory } from './IQuickLinksState';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { stringIsNullOrEmpty } from "@pnp/common";

export default class QuickLinks extends React.Component<IQuickLinksProps, IQuickLinksState> {
  public constructor(props: IQuickLinksProps) {
    super(props);

    this.state = { linkStructure: [] };
  }

  // eslint-disable-next-line react/no-deprecated
  public componentWillMount() { // TODO remove (deprecated)
    this.fetchData();
  }

  public render(): React.ReactElement<IQuickLinksProps> {
    const theme: IReadonlyTheme = this.props.theme;
    const backgroundColor: string = theme?.semanticColors?.bodyBackground ?? "#ffffff";
    const links: JSX.Element[] = this.generateLinks(this.state.linkStructure);
    return (
      <div className={styles.quickLinks} style={{ backgroundColor }}>
        <div className={styles.webpartHeader}>
          <span>{this.props.title}</span>
          <span className={styles.showAll}>
            <Text onClick={() => window.open(this.props.allLinksUrl, '_blank')}>
              {strings.component_AllLinksLabel}
            </Text>
          </span>
        </div>
        <div className={styles.linkGrid}>
          {links}
        </div>
      </div>
    );
  }

  private generateLinks(categories: Array<ICategory>) {
    return (
      categories.map((cat: ICategory, catIndex: number): JSX.Element => {
        const linkItems: JSX.Element[] = cat.links.map((link: ILink, linkIndex): JSX.Element => {
          const linkIcon: JSX.Element = (
            <Icon 
              className={styles.icon} 
              style={{ opacity: this.props.iconOpacity / 100 }}
              iconName={(link.icon) ? link.icon : this.props.defaultIcon}
            />
          );
          const linkStyle = { width: this.props.maxLinkLength };
          const linkTarget = link.openInSameTab ? '_self' : '_blank';
          return (
            <div key={`link_${linkIndex}`} className={styles.linkGridColumn} style={{ lineHeight: `${this.props.lineHeight}px` }}>
              <Text className={styles.linkContainer} onClick={() => {
                  this.callWebHook(link.url, link.category);
                  window.open(link.url, linkTarget)
                }
              }>
                {linkIcon}
                <span style={linkStyle}>{link.displayText}</span>
              </Text>
            </div>
          );
        });
        if (this.props.groupByCategory) {
          return (
          <div className={styles.categorySection}><div className={styles.linkCategoryHeading}>{cat.displayText}</div>{linkItems}</div>
          );
        }
        return <div key={`category_${catIndex}`}>{linkItems}</div>;
      })
    );
  }

  private async callWebHook(uri: string, category: string): Promise<any> {
    
    if( stringIsNullOrEmpty(this.props.linkClickWebHook) ) {
      return;
    }
    
    const body = {
      uri: uri,
      category: category
    }

    const postRequest = {
      method: 'POST',
      body: JSON.stringify(body),
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-cache'
      }
    }

    fetch(this.props.linkClickWebHook, postRequest)

  }

  private async fetchData() {
    const searchString: string = `AuthorId eq '${this.props.userId}'`;
    
    const editorLinks = (
      await 
        sp.web
        .getList(this.props.webServerRelativeUrl + "/Lists/EditorLinks")
        .items
        .filter("(PzlLinkActive eq 1) and (PzlLinkMandatory eq 1)")
        .orderBy("PzlLinkPriority")
        .orderBy("Title").
        top(this.props.numberOfLinks)
        .get()
    );

    const newNonMandatoryLinks = (
      await 
        sp.web
        .getList(this.props.webServerRelativeUrl + "/Lists/EditorLinks")
        .items
        .filter("(PzlLinkActive eq 1) and (PzlLinkMandatory eq 0)")
        .orderBy("PzlLinkPriority")
        .orderBy("Title")
        .top(this.props.numberOfLinks)
        .get()
    );

    const newNonMandatoryLinksObject = newNonMandatoryLinks.map((link) => {
      return { 
        id: link.Id, 
        displayText: link.Title, 
        url: link.PzlUrl, 
        icon: link.PzlOfficeUIFabricIcon, 
        priority: link.PzlLinkPriority, 
        category: link.PzlLinkCategory || "Ingen kategori", 
        openInSameTab: link.PzlOpenInSameTab 
      };
    });

    const favouriteLinkStrings = (
      await 
        sp.web
        .getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks")
        .items
        .select("Id", "AuthorId", "PzlPersonalLinks")
        .filter(searchString)
        .get()
    );

    const favouriteLinksObject: ILink[] = (
      favouriteLinkStrings.length > 0 ? 
        JSON.parse(favouriteLinkStrings[0].PzlPersonalLinks) : 
        []
    );

    const displayLinks = editorLinks.map(link => {
      return { 
        displayText: link.Title, 
        url: link.PzlUrl, 
        icon: link.PzlOfficeUIFabricIcon || "Link", 
        priority: link.PzlLinkPriority || "0", 
        category: link.PzlLinkCategory || "Ingen kategori", 
        openInSameTab: link.PzlOpenInSameTab 
      };
    });

    if (favouriteLinkStrings.length > 0) {
      const updatedFavoriteLinksObject = (
        await this.checkForUpdatedLinks(favouriteLinksObject, newNonMandatoryLinksObject, favouriteLinkStrings[0].Id)
      );
      displayLinks.push(...updatedFavoriteLinksObject);
    }

    let categories: Array<ICategory> = [{ displayText: strings.component_NoCategoryLabel, links: displayLinks }];
    if (this.props.groupByCategory) {
      let categoryNames: string[] = displayLinks.map(lnk => { return lnk.category; }).sort();
      categoryNames = categoryNames.filter((item, index) => { return categoryNames.indexOf(item) == index; });
      categories = categoryNames.map(catName => {
        return { displayText: catName, links: displayLinks.filter(lnk => { return lnk.category === catName; }) };
      });
    }

    this.setState({ linkStructure: categories });
  }

  private async checkForUpdatedLinks(userFavoriteLinks: ILink[], allFavoriteLinks: ILink[], currentItemId: number) {
    const personalLinks: ILink[] = new Array<ILink>();
    let shouldUpdate: boolean = false;
    userFavoriteLinks.forEach((userLink: ILink): void => {
      const linkMatch: ILink = find(allFavoriteLinks, (favoriteLink => favoriteLink.id === userLink.id));
      if (linkMatch && (!isEqual(linkMatch.url, userLink.url) || !isEqual(linkMatch.displayText, userLink.displayText) || !isEqual(linkMatch.icon, userLink.icon))) {
        shouldUpdate = true;
        personalLinks.push(linkMatch);
      } else {
        personalLinks.push(userLink);
      }
    });
    if (shouldUpdate) {
      await this.updatePersonalLinks(personalLinks, currentItemId);
    }
    return personalLinks;
  }

  private async updatePersonalLinks(newFavoriteLinks, itemId: number) {
    try {
      await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks").items.getById(itemId).update({
        PzlPersonalLinks: JSON.stringify(newFavoriteLinks),
      });
    } catch (e) {
      console.log(e);
    }
  }
}
