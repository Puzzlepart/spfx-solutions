import * as React from 'react';
import styles from './AllLinks.module.scss';
import { IAllLinksProps } from './IAllLinksProps';
import { IAllLinksState, Link, User, LinkType, ILink, ICategory } from './IAllLinksState';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Button, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Spinner, SpinnerSize, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import * as strings from 'AllLinksWebPartStrings';
import "@pnp/polyfill-ie11";
import { sp } from "@pnp/sp";
import { find, isEqual, findIndex } from '@microsoft/sp-lodash-subset';
import { Text } from 'office-ui-fabric-react/lib/Text';
import { stringIsNullOrEmpty } from '@pnp/common';
import { add } from 'lodash';

enum CategoryOperation {add,remove}
export default class AllLinks extends React.Component<IAllLinksProps, IAllLinksState> {
  public constructor(props) {
    super(props);
    this.state = {
      mandatoryLinks: undefined,
      editorLinks: undefined,
      categoryLinks: undefined,
      favouriteLinks: [],
      showModal: false,
      saveButtonDisabled: true,
      isLoading: true
    };
  }


  public render(): React.ReactElement<IAllLinksProps> {
    if (this.state.isLoading) {
      return (
        <Spinner type={SpinnerType.large} />
      );
    } else {
      let mandatoryLinks = this.state.mandatoryLinks ? this.generateMandatoryLinkComponents(this.state.mandatoryLinks) : null;
      let editorLinks = this.state.editorLinks ? this.generateEditorLinkComponents(this.state.editorLinks) : null;
      let favouriteLinks = this.state.favouriteLinks ? this.generateFavouriteLinkComponents(this.state.favouriteLinks) : null;
      let newLinkModal = this.state.showModal ? this.generateNewLinkModal() : null;
      let categoryLinks = this.generateLinks(this.state.categoryLinks);
      let errorMessage = this.state.showErrorMessage ? <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ showErrorMessage: false })}>{strings.component_SaveErrorLabel}</MessageBar> : null;
      //let successMessage = this.state.showSuccessMessage ? <MessageBar messageBarType={MessageBarType.success} onDismiss={() => this.setState({ showSuccessMessage: false })} >{strings.component_SaveOkLabel}</MessageBar> : null;
      let loadingSpinner = this.state.showLoadingSpinner ? <Spinner style={{position:'absolute',right:10,top:-10}} className={styles.spinner} size={SpinnerSize.small} /> : null;
      let linkListing = this.props.listingByCategory ?
        <div className={styles.allLinks}>
          <div className={styles.webpartHeader}>
            <span>{this.props.listingByCategoryTitle}</span>
          </div>
          <div className={styles.linkGrid}>
            {categoryLinks}
          </div>
        </div>
        :
        <div>
          <div className={styles.webpartHeading} >{stringIsNullOrEmpty(this.props.mandatoryLinksTitle) ? strings.component_MandatoryLinksLabel : this.props.mandatoryLinksTitle}</div>
          <div className={styles.editorLinksContainer}>{mandatoryLinks}</div>
          <hr />
          <div className={styles.webpartHeading} >{stringIsNullOrEmpty(this.props.reccomendedLinksTitle) ? strings.component_PromotedLinksLabel : this.props.reccomendedLinksTitle}</div>
          <div className={styles.editorLinksContainer}>{editorLinks}</div>
        </div>;
      let myLinks = <div>
        <div className={styles.webpartHeading}>{stringIsNullOrEmpty(this.props.myLinksTitle) ? strings.component_YourLinksLabel : this.props.myLinksTitle}</div>
        <div className={styles.editorLinksContainer}>{favouriteLinks}</div>
        <div className={styles.buttonRow} >
          {/* <Button onClick={() => this.saveData()} text={strings.component_SaveYourLinksLabel} disabled={this.state.saveButtonDisabled} /> */}
          <Button onClick={() => this.openNewItemModal()} text={strings.component_NewLinkLabel} iconProps={{ iconName: 'Add' }} />

        </div>
      </div>;
      return (
        <div className={styles.allLinks}>
          {errorMessage}

          {loadingSpinner}
          {newLinkModal}
          {(this.props.mylinksOnTop) ?
            <div>
              {myLinks}
              <hr />
              {linkListing}
            </div>
            :
            <div>
              {linkListing}
              <hr />
              {myLinks}
            </div>
          }
        </div>
      );
    }
  }

  public componentWillMount() {
    this.fetchData();
  }

  private openNewItemModal() {
    let emptyLink: Link = {
      id: -1,
      displayText: "",
      url: "",
      icon: this.props.defaultIcon,
      priority: "1000",
      mandatory: 0,
      linkType: LinkType.favouriteLinks
    };

    this.setState({
      showModal: true,
      modalData: emptyLink
    });
  }

  private appendToFavourites(link: Link) {
    let newFavourites = this.state.favouriteLinks.slice();
    newFavourites.push(link);
    let newEditorLinks = this.state.editorLinks.slice();
    newEditorLinks.splice(newEditorLinks.indexOf(link), 1);
    let categoryLinks = this.state.categoryLinks;

    categoryLinks = this.updateCategoryLinks(CategoryOperation.remove, link as ILink, categoryLinks);

    this.setState({
      favouriteLinks: newFavourites,
      editorLinks: newEditorLinks,
      saveButtonDisabled: false
    },()=>this.saveData());

  }


  private removeFromFavourites(link: Link) {
    let newEditorLinks = this.state.editorLinks.slice();
    newEditorLinks.push(link);
    let newFavourites = this.state.favouriteLinks.slice();
    newFavourites.splice(newFavourites.indexOf(link), 1);

    let categoryLinks = this.state.categoryLinks;

    categoryLinks = this.updateCategoryLinks(CategoryOperation.add, link as ILink, categoryLinks);

    this.setState({
      favouriteLinks: newFavourites,
      editorLinks: newEditorLinks,
      categoryLinks: categoryLinks,
      saveButtonDisabled: false
    },()=>this.saveData());
  }

  private updateCategoryLinks(operation: CategoryOperation, link: ILink, categoryLinks: ICategory[]) {
    if (this.props.listingByCategory) {

      categoryLinks = categoryLinks.map(category => {
        if (category.displayText === link['category']) {
          if(operation === CategoryOperation.remove ){
            category.links = category.links.filter(clink => link.url !== clink.url);
          } else {
            category.links.push(link);
          }
        }
        return category;
      });
    }
    return categoryLinks;
  }

  private removeCustomFromFavourites(link: Link) {
    let newFavourites = this.state.favouriteLinks.slice();
    newFavourites.splice(newFavourites.indexOf(link), 1);
    this.setState({
      favouriteLinks: newFavourites,
      saveButtonDisabled: false
    });
  }

  private generateEditorLinkComponents(links: Array<Link>) {
    return links.map(link => {
      return (
        <div className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, "_blank")}>
            <Icon iconName={(link.icon) ? link.icon : this.props.defaultIcon} className={styles.icon} />
            <span title={link.displayText}>{link.displayText}</span>
          </Text>
          <Icon className={styles.actionIcon} iconName='CirclePlus' onClick={() => this.appendToFavourites(link)} />
        </div>
      );
    });
  }

  private generateMandatoryLinkComponents(links: Array<Link>) {
    return links.map(link => {
      return (
        <div className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, "_blank")}>
            <Icon iconName={(link.icon) ? link.icon : this.props.defaultIcon} className={styles.icon} />
            <span>{link.displayText}</span>
          </Text>
        </div>
      );
    });
  }

  private generateFavouriteLinkComponents(links: Array<Link>) {
    return links.map(link => {
      let linkIcon = <Icon iconName={(link.icon) ? link.icon : this.props.defaultIcon} className={styles.icon} />;
      let removeLinkButton = link.linkType === LinkType.editorLink ?
        <Icon className={styles.actionIcon} iconName='SkypeCircleMinus' onClick={() => this.removeFromFavourites(link)} /> :
        <Icon className={styles.actionIcon} iconName='SkypeCircleMinus' onClick={() => this.removeCustomFromFavourites(link)} />;
      return (
        <div className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, "_blank")}>
            {linkIcon}
            <div>{link.displayText}</div>
          </Text>
          {removeLinkButton}
        </div>
      );
    });
  }

  private generateNewLinkModal() {
    return (
      <Modal isOpen={this.state.showModal} isBlocking={false} containerClassName={styles.newLinkModal} onDismiss={() => this.setState({ modalData: null, showModal: false })}>
        <div className={styles.modalHeader}>{strings.component_NewLinkLabel}</div>
        <div className={styles.modalBody}>
          <TextField label='Url' onChanged={(newVal) => this.onModalValueChanged('url', newVal)} value={this.state.modalData['url']} onGetErrorMessage={(value) => this.getUrlErrorMessage(value)} />
          <TextField label={strings.component_TitleLabel} onChanged={(newVal) => this.onModalValueChanged('displayText', newVal)} value={this.state.modalData['displayText']} />
          <DefaultButton text={strings.component_AddLabel} onClick={() => this.addNewLink()} />
          <Button text={strings.component_CancelLabel} onClick={() => { this.setState({ modalData: null, showModal: false }); }} />
        </div>
      </Modal>
    );
  }

  private addNewLink() {
    let newFavourites = this.state.favouriteLinks.slice();
    newFavourites.push(this.state.modalData);
    this.setState({
      favouriteLinks: newFavourites,
      modalData: null,
      showModal: false,
      saveButtonDisabled: false
    },()=>this.saveData());
  }

  private onModalValueChanged(field, newVal) {
    let newModalData: Link = { ...this.state.modalData };
    newModalData[field] = newVal;
    this.setState({
      modalData: newModalData
    });
  }

  private getUrlErrorMessage(value: any): string {
    if (value.length > 0) {
      let urlRegex = /(http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-]*[\w@?^=%&/~+#-])?/;
      return value.match(urlRegex) === null ? strings.component_UrlValidationLabel : '';
    } else {
      return '';
    }
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
            {(link.mandatory) ?

              <Icon className={styles.icon} iconName='Lock' title={strings.component_action_removeMandatory} />
              :
              <Icon className={styles.actionIcon} iconName='CirclePlus' onClick={() => this.appendToFavourites(link)} />
            }
          </div>
        );
      });
      if (this.props.listingByCategory) {
        return <div className={styles.categorySection}><div className={styles.linkCategoryHeading}>{cat.displayText}</div>{linkItems}</div>;
      }
      return <div>{linkItems}</div>;
    });
  }

  private async fetchData() {
    try {
      let searchString = `AuthorId eq '${this.props.currentUserId}'`;
      let favouriteLinkListItem = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks").items.select("Id", "AuthorId", "PzlPersonalLinks").filter(searchString).get();
      let favouriteItemsIds: Array<number>;
      let favouriteItems: Array<Link> = [];
      if (favouriteLinkListItem.length > 0 && favouriteLinkListItem[0]["PzlPersonalLinks"] !== null) {
        favouriteItems = JSON.parse(favouriteLinkListItem[0]["PzlPersonalLinks"]);
        favouriteItemsIds = favouriteItems.map(link => link.id);
        this.setState({
          favouriteLinks: favouriteItems
        });
      } else {
        favouriteItemsIds = [];
        this.setState({
          isFirstUpdate: true,
        });
      }
      let editorLinks = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/EditorLinks").items.filter("PzlLinkActive eq 1").get();
      let mappedLinks: Array<Link> = editorLinks.map(link => {
        return { id: link.Id, displayText: link.Title, url: link.PzlUrl, icon: link.PzlOfficeUIFabricIcon, priority: link.PzlPriority, mandatory: link.PzlLinkMandatory, linkType: LinkType.editorLink };
      });
      let mandatorymappedLinks = mappedLinks.filter(mandatory => mandatory.mandatory == 1);
      let promotedmappedLinks = mappedLinks.filter(mandatory => mandatory.mandatory == 0);
      let prunedLinks: Array<Link> = promotedmappedLinks.filter(link => {
        return favouriteItemsIds.indexOf(link.id) === -1;
      });
      if (favouriteLinkListItem.length > 0 && favouriteItems !== null && favouriteItems.length > 0) {
        let favoriteLinks = await this.checkForUpdatedLinks(favouriteItems, promotedmappedLinks);
        favouriteItemsIds = favoriteLinks.map(item => item.id);
      }
      let linkFieldId = favouriteLinkListItem.length > 0 ? favouriteLinkListItem[0].Id : null;
      let currentUser: User = {
        id: this.props.currentUserId,
        linkFieldId: linkFieldId
      };

      let displayLinks = editorLinks.map(link => {
        return { displayText: link.Title, url: link.PzlUrl, icon: link.PzlOfficeUIFabricIcon || "Link", priority: link.PzlLinkPriority || "0", category: link.PzlLinkCategory || "Ingen kategori", mandatory: link.PzlLinkMandatory, linkType: LinkType.editorLink };
      });

      let categories: Array<ICategory> = [{ displayText: strings.component_NoCategoryLabel, links: displayLinks }];
      if (this.props.listingByCategory) {
        let categoryNames: string[] = displayLinks.map(lnk => { return lnk.category; }).sort();
        categoryNames = categoryNames.filter((item, index) => { return categoryNames.indexOf(item) == index; });
        categories = categoryNames.map(catName => {
          return { displayText: catName, links: displayLinks.filter(lnk => { return lnk.category === catName; }) };
        });
      }

      this.setState({
        currentUser: currentUser,
        editorLinks: prunedLinks,
        mandatoryLinks: mandatorymappedLinks,
        categoryLinks: categories,
        isLoading: false,
      });
    } catch (err) {
      console.log(err);
      this.setState({
        isLoading: false,
      });
    }
  }

  private async saveData() {
    this.setState({
      showLoadingSpinner: true,
      saveButtonDisabled: true,
    });
    try {
      let linksAsString = JSON.stringify(this.state.favouriteLinks);
      if (this.state.isFirstUpdate) {
        let result = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks").items.add({
          'PzlPersonalLinks': linksAsString,
          'Title': this.props.currentUserName,
        });
        let currentUser: User = {
          id: this.state.currentUser.id,
          linkFieldId: result.data.Id,
        };
        this.setState({
          isFirstUpdate: false,
          saveButtonDisabled: true,
          currentUser: currentUser,
          showSuccessMessage: true,
          showLoadingSpinner: false,
        });
        setTimeout(() => this.setState({ showSuccessMessage: false }), 5000);
      } else {
        let result = await sp.web.getList(this.props.webServerRelativeUrl + "/Lists/FavouriteLinks").items.getById(+this.state.currentUser.linkFieldId).update({
          'PzlPersonalLinks': linksAsString,
        });
        this.setState({
          saveButtonDisabled: true,
          showSuccessMessage: true,
          showLoadingSpinner: false,
        });
        setTimeout(() => this.setState({ showSuccessMessage: false }), 5000);
      }
    } catch (err) {
      this.setState({
        showErrorMessage: true,
        showLoadingSpinner: false,
        saveButtonDisabled: false,
      });
      setTimeout(() => this.setState({ showErrorMessage: false }), 5000);
    }
  }


  private async checkForUpdatedLinks(userFavoriteLinks: any[], allFavoriteLinks: any[]) {
    let personalLinks = new Array<Link>();
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
      this.setState({
        favouriteLinks: personalLinks
      });
    }
    return personalLinks;
  }
}
