import * as React from 'react';
import styles from './QuickLinks.module.scss';
import { IQuickLinksProps } from './IQuickLinksProps';
import { find, isEqual, findIndex } from '@microsoft/sp-lodash-subset';
import { IQuickLinksState, ILink, ICategory } from './IQuickLinksState';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import * as strings from 'QuickLinksWebPartStrings';

export default class QuickLinks extends React.Component<IQuickLinksProps, IQuickLinksState> {

    public constructor(props) {
        super(props);

        this.state = { linkStructure: [] };
    }

    public componentWillMount() {
        this.fetchData();
    }


    public render(): React.ReactElement<IQuickLinksProps> {
        let links = this.generateLinks(this.state.linkStructure);
        return (
            <div className={styles.quickLinks} >
                <div className={styles.webpartHeader}>
                    <span>{this.props.title}</span>
                    <span className={styles.showAll}><a href={this.props.allLinksUrl}>{strings.component_AllLinksLabel}</a></span>
                </div>
                <div className={styles.linkGrid}>
                    {links}
                </div>
            </div>
        );
    }

    private generateLinks(categories: Array<ICategory>) {
        return categories.map(cat => {
            let linkItems = cat.links.map(link => {
                let linkIcon = <Icon iconName={(link.icon) ? link.icon : this.props.defaultIcon} className={styles.icon} />;
                let linkStyle = { width: this.props.maxLinkLength };
                let linkTarget = link.openInSameTab ? '_self' : '_blank';
                return (
                    <div className={styles.linkGridColumn}>
                        <a className={styles.linkContainer} data-interception="off" href={link.url} title={link.displayText} target={linkTarget}>
                            {linkIcon}
                            <span style={linkStyle}>{link.displayText}</span>
                        </a>
                    </div>
                );
            });
            if (this.props.groupByCategory) {
                return <div className={styles.categorySection}><div className={styles.linkCategoryHeading}>{cat.displayText}</div>{linkItems}</div>;
            }
            return <div>{linkItems}</div>;
        });
    }

    private async fetchData() {
        let searchString = `AuthorId eq '${this.props.userId}'`;
        let editorLinks = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/EditorLinks").items.filter("(PzlLinkActive eq 1) and (PzlLinkMandatory eq 1)").orderBy("PzlLinkPriority").orderBy("Title").top(this.props.numberOfLinks).get();
        let newNonMandatoryLinks = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/EditorLinks").items.filter("(PzlLinkActive eq 1) and (PzlLinkMandatory eq 0)").orderBy("PzlLinkPriority").orderBy("Title").top(this.props.numberOfLinks).get();
        let newNonMandatoryLinksObject = newNonMandatoryLinks.map(link => {
            return {id: link.Id, displayText: link.Title, url: link.PzlUrl, icon: link.PzlOfficeUIFabricIcon, priority: link.PzlLinkPriority, category: link.PzlLinkCategory || "Ingen kategori", openInSameTab: link.PzlOpenInSameTab };
        });
        let favouriteLinkStrings = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks").items.select("Id", "AuthorId", "PzlPersonalLinks").filter(searchString).get();
        let favouriteLinksObject: Array<ILink> = favouriteLinkStrings.length > 0 ? JSON.parse(favouriteLinkStrings[0].PzlPersonalLinks) : [];
        
        let displayLinks = editorLinks.map(link => {
            return { displayText: link.Title, url: link.PzlUrl, icon: link.PzlOfficeUIFabricIcon || "Link", priority: link.PzlLinkPriority || "0", category: link.PzlLinkCategory || "Ingen kategori", openInSameTab: link.PzlOpenInSameTab };
        });
        
        if (favouriteLinkStrings.length > 0) {
            let updatedFavoriteLinksObject = await this.checkForUpdatedLinks(favouriteLinksObject, newNonMandatoryLinksObject, favouriteLinkStrings[0].Id);
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
        this.setState({
            linkStructure: categories
        });
    }

    private async checkForUpdatedLinks(userFavoriteLinks: ILink[], allFavoriteLinks: ILink[], currentItemId: number) {
        let personalLinks = new Array<ILink>(); 
        let shouldUpdate = false;
        userFavoriteLinks.forEach(userLink => {
            let linkMatch = find(allFavoriteLinks, (favoriteLink => favoriteLink.id === userLink.id));
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
            let result = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks").items.getById(itemId).update({
                PzlPersonalLinks:  JSON.stringify(newFavoriteLinks),
            });
        } catch(e) {
            console.log(e);
        }
    }
}
