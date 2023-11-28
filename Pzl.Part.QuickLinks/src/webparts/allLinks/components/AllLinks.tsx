import * as React from 'react'
import styles from './AllLinks.module.scss'
import { IAllLinksProps } from './IAllLinksProps'
import { IAllLinksState, Link, User, LinkType, ILink, ICategory } from './IAllLinksState'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import { Button, DefaultButton } from 'office-ui-fabric-react/lib/Button'
import { TextField } from 'office-ui-fabric-react/lib/TextField'
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar'
import { Spinner, SpinnerSize, SpinnerType } from 'office-ui-fabric-react/lib/Spinner'
import { Modal } from 'office-ui-fabric-react/lib/Modal'
import * as strings from 'AllLinksWebPartStrings'
import '@pnp/polyfill-ie11'
import { ItemAddResult, sp } from '@pnp/sp'
import { isEqual } from '@microsoft/sp-lodash-subset'
import { Text } from 'office-ui-fabric-react/lib/Text'
import { stringIsNullOrEmpty } from '@pnp/common'
import { IReadonlyTheme } from '@microsoft/sp-component-base'

enum CategoryOperation {
  add,
  remove
}

export default class AllLinks extends React.Component<IAllLinksProps, IAllLinksState> {
  public constructor(props: IAllLinksProps) {
    super(props)
    this.state = {
      mandatoryLinks: undefined,
      editorLinks: undefined,
      categoryLinks: undefined,
      favouriteLinks: [],
      showModal: false,
      saveButtonDisabled: true,
      isLoading: true
    }
  }

  public render(): React.ReactElement<IAllLinksProps> {
    const theme: IReadonlyTheme = this.props.theme
    const backgroundColor: string = theme?.semanticColors?.bodyBackground ?? '#ffffff'

    console.log('LOG BG: ', backgroundColor, theme)

    if (this.state.isLoading) {
      return <Spinner styles={{ root: { backgroundColor } }} type={SpinnerType.large} />
    }
    const mandatoryLinks: JSX.Element[] = this.state.mandatoryLinks
      ? this.generateMandatoryLinkComponents(this.state.mandatoryLinks)
      : null
    const editorLinks: JSX.Element[] = this.state.editorLinks
      ? this.generateEditorLinkComponents(this.state.editorLinks)
      : null
    const favouriteLinks: JSX.Element[] = this.state.favouriteLinks
      ? this.generateFavouriteLinkComponents(this.state.favouriteLinks)
      : null
    const newLinkModal: JSX.Element = this.state.showModal ? this.generateNewLinkModal() : null
    const categoryLinks: JSX.Element[] = this.generateLinks(this.state.categoryLinks)
    const errorMessage: JSX.Element = this.state.showErrorMessage ? (
      <MessageBar
        messageBarType={MessageBarType.error}
        onDismiss={() => this.setState({ showErrorMessage: false })}
      >
        {strings.SaveErrorLabel}
      </MessageBar>
    ) : null
    const loadingSpinner: JSX.Element = this.state.showLoadingSpinner ? (
      <Spinner
        style={{ position: 'absolute', right: 10, top: -10 }}
        className={styles.spinner}
        size={SpinnerSize.small}
      />
    ) : null
    const linkListing: JSX.Element = this.props.listingByCategory ? (
      <div className={styles.allLinks}>
        <div className={styles.webpartHeader}>
          <span>{this.props.listingByCategoryTitle}</span>
        </div>
        <div className={styles.linkGrid}>{categoryLinks}</div>
      </div>
    ) : (
      <div>
        <div className={styles.webpartHeading}>
          {stringIsNullOrEmpty(this.props.mandatoryLinksTitle)
            ? strings.MandatoryLinksLabel
            : this.props.mandatoryLinksTitle}
        </div>
        <div className={styles.editorLinksContainer}>{mandatoryLinks}</div>
        <hr />
        <div className={styles.webpartHeading}>
          {stringIsNullOrEmpty(this.props.recommendedLinksTitle)
            ? strings.RecommendedLinksLabel
            : this.props.recommendedLinksTitle}
        </div>
        <div className={styles.editorLinksContainer}>{editorLinks}</div>
      </div>
    )
    const myLinks: JSX.Element = (
      <div>
        <div className={styles.webpartHeading}>
          {stringIsNullOrEmpty(this.props.myLinksTitle)
            ? strings.YourLinksLabel
            : this.props.myLinksTitle}
        </div>
        <div className={styles.editorLinksContainer}>{favouriteLinks}</div>
        <div className={styles.buttonRow}>
          <Button
            onClick={() => this.openNewItemModal()}
            text={strings.NewLinkLabel}
            iconProps={{ iconName: 'Add' }}
          />
        </div>
      </div>
    )
    return (
      <div className={styles.allLinks} style={{ backgroundColor }}>
        {errorMessage}
        {loadingSpinner}
        {newLinkModal}
        {this.props.mylinksOnTop ? (
          <div>
            {myLinks}
            <hr />
            {linkListing}
          </div>
        ) : (
          <div>
            {linkListing}
            <hr />
            {myLinks}
          </div>
        )}
      </div>
    )
  }

  // eslint-disable-next-line react/no-deprecated
  public componentWillMount(): void {
    // TODO remove (deprecated)
    this.fetchData()
  }

  private openNewItemModal(): void {
    const emptyLink: Link = {
      id: -1,
      displayText: '',
      url: '',
      icon: this.props.defaultIcon,
      priority: '1000',
      mandatory: 0,
      linkType: LinkType.favouriteLinks
    }

    this.setState({
      showModal: true,
      modalData: emptyLink
    })
  }

  private appendToFavourites(link: Link): void {
    const newFavourites: Link[] = this.state.favouriteLinks.slice()
    newFavourites.push(link)

    const newEditorLinks: Link[] = this.state.editorLinks.slice()
    newEditorLinks.splice(newEditorLinks.indexOf(link), 1)

    this.updateCategoryLinks(CategoryOperation.remove, link as ILink, this.state.categoryLinks)

    this.setState(
      {
        favouriteLinks: newFavourites,
        editorLinks: newEditorLinks,
        saveButtonDisabled: false
      },
      (): Promise<unknown> => this.saveData()
    )
  }

  private removeFromFavourites(link: Link): void {
    const newEditorLinks: Link[] = this.state.editorLinks.slice()
    newEditorLinks.push(link)

    const newFavourites: Link[] = this.state.favouriteLinks.slice()
    newFavourites.splice(newFavourites.indexOf(link), 1)

    const categoryLinks: ICategory[] = this.updateCategoryLinks(
      CategoryOperation.add,
      link as ILink,
      this.state.categoryLinks
    )

    this.setState(
      {
        favouriteLinks: newFavourites,
        editorLinks: newEditorLinks,
        categoryLinks: categoryLinks,
        saveButtonDisabled: false
      },
      (): Promise<void> => this.saveData()
    )
  }

  private updateCategoryLinks(
    operation: CategoryOperation,
    link: ILink,
    categoryLinks: ICategory[]
  ): ICategory[] {
    if (this.props.listingByCategory) {
      categoryLinks = categoryLinks.map((category) => {
        if (category.displayText === link['category']) {
          if (operation === CategoryOperation.remove) {
            category.links = category.links.filter((clink) => link.url !== clink.url)
          } else {
            category.links.push(link)
          }
        }
        return category
      })
    }
    return categoryLinks
  }

  private removeCustomFromFavourites(link: Link): void {
    const newFavourites: Link[] = this.state.favouriteLinks.slice()
    newFavourites.splice(newFavourites.indexOf(link), 1)
    this.setState(
      {
        favouriteLinks: newFavourites,
        saveButtonDisabled: false
      },
      (): Promise<void> => this.saveData()
    )
  }

  private generateEditorLinkComponents(links: Array<Link>): JSX.Element[] {
    return links.map((link: Link, index: number): JSX.Element => {
      return (
        <div key={`editor_link_${index}`} className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, '_blank')}>
            <Icon
              iconName={link.icon ? link.icon : this.props.defaultIcon}
              className={styles.icon}
            />
            <span title={link.displayText}>{link.displayText}</span>
          </Text>
          <Icon
            className={styles.actionIcon}
            iconName='CirclePlus'
            onClick={() => this.appendToFavourites(link)}
          />
        </div>
      )
    })
  }

  private generateMandatoryLinkComponents(links: Array<Link>): JSX.Element[] {
    return links.map((link: Link, index: number): JSX.Element => {
      return (
        <div key={`required_link_${index}`} className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, '_blank')}>
            <Icon
              iconName={link.icon ? link.icon : this.props.defaultIcon}
              className={styles.icon}
            />
            <span>{link.displayText}</span>
          </Text>
        </div>
      )
    })
  }

  private generateFavouriteLinkComponents(links: Array<Link>): JSX.Element[] {
    return links.map((link: Link, index: number): JSX.Element => {
      const linkIcon: JSX.Element = (
        <Icon iconName={link.icon ? link.icon : this.props.defaultIcon} className={styles.icon} />
      )
      const removeLinkButton: JSX.Element =
        link.linkType === LinkType.editorLink ? (
          <Icon
            className={styles.actionIcon}
            iconName='SkypeCircleMinus'
            onClick={() => this.removeFromFavourites(link)}
          />
        ) : (
          <Icon
            className={styles.actionIcon}
            iconName='SkypeCircleMinus'
            onClick={() => this.removeCustomFromFavourites(link)}
          />
        )
      return (
        <div key={`favourite_link_${index}`} className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, '_blank')}>
            {linkIcon}
            <div>{link.displayText}</div>
          </Text>
          {removeLinkButton}
        </div>
      )
    })
  }

  private generateNewLinkModal(): JSX.Element {
    return (
      <Modal
        isOpen={this.state.showModal}
        isBlocking={false}
        containerClassName={styles.newLinkModal}
        onDismiss={() => this.setState({ modalData: null, showModal: false })}
      >
        <div className={styles.modalHeader}>{strings.NewLinkLabel}</div>
        <div className={styles.modalBody}>
          <TextField
            label='Url'
            onChange={(_, newVal: any): void => this.onModalValueChanged('url', newVal)}
            value={this.state.modalData['url']}
            onGetErrorMessage={(value) => this.getUrlErrorMessage(value)}
          />
          <TextField
            label={strings.TitleLabel}
            onChange={(_, newVal: any) => this.onModalValueChanged('displayText', newVal)}
            value={this.state.modalData['displayText']}
          />
          <DefaultButton text={strings.AddLabel} onClick={() => this.addNewLink()} />
          <Button
            text={strings.CancelLabel}
            onClick={() => {
              this.setState({ modalData: null, showModal: false })
            }}
          />
        </div>
      </Modal>
    )
  }

  private addNewLink(): void {
    const newFavourites: Link[] = this.state.favouriteLinks.slice()
    newFavourites.push(this.state.modalData)
    this.setState(
      {
        favouriteLinks: newFavourites,
        modalData: null,
        showModal: false,
        saveButtonDisabled: false
      },
      () => this.saveData()
    )
  }

  private onModalValueChanged(field: string, newVal: any): void {
    const newModalData: Link = { ...this.state.modalData }
    newModalData[field] = newVal
    this.setState({ modalData: newModalData })
  }

  private getUrlErrorMessage(value: any): string {
    if (value.length > 0) {
      const urlRegex: RegExp =
        /(http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-]*[\w@?^=%&/~+#-])?/
      return value.match(urlRegex) === null ? strings.UrlValidationLabel : ''
    } else {
      return ''
    }
  }

  private generateLinks(categories: Array<ICategory>): JSX.Element[] {
    return categories.map((cat: ICategory, index: number): JSX.Element => {
      const linkItems: JSX.Element[] = cat.links.map(
        (link: ILink, subIndex: number): JSX.Element => {
          const linkIcon: JSX.Element = (
            <Icon
              iconName={link.icon ? link.icon : this.props.defaultIcon}
              className={styles.icon}
            />
          )
          const linkStyle = { width: this.props.maxLinkLength }
          const linkTarget: string = link.openInSameTab ? '_self' : '_blank'
          return (
            <div key={`link_cat_sub_${subIndex}`} className={styles.linkGridColumn}>
              <a
                className={styles.linkContainer}
                data-interception='off'
                href={link.url}
                title={link.displayText}
                target={linkTarget}
              >
                {linkIcon}
                <span style={linkStyle}>{link.displayText}</span>
              </a>
              {link.mandatory ? (
                <Icon
                  className={styles.icon}
                  iconName='Lock'
                  title={strings.ActionRemoveMandatory}
                />
              ) : (
                <Icon
                  className={styles.actionIcon}
                  iconName='CirclePlus'
                  onClick={() => this.appendToFavourites(link)}
                />
              )}
            </div>
          )
        }
      )
      if (this.props.listingByCategory) {
        return (
          <div key={`link_cat_${index}`} className={styles.categorySection}>
            <div className={styles.linkCategoryHeading}>{cat.displayText}</div>
            {linkItems}
          </div>
        )
      }
      return <div key={`link_no_cat_${index}`}>{linkItems}</div>
    })
  }

  private async fetchData(): Promise<void> {
    try {
      const searchString: string = `AuthorId eq '${this.props.currentUserId}'`
      const favouriteLinkListItem = await sp.web
        .getList(this.props.webServerRelativeUrl + '/Lists/FavouriteLinks')
        .items.select('Id', 'AuthorId', 'PzlPersonalLinks')
        .filter(searchString)
        .get()
      let favouriteItemsIds: number[]
      let favouriteItems: Link[] = []
      if (
        favouriteLinkListItem.length > 0 &&
        favouriteLinkListItem[0]['PzlPersonalLinks'] !== null
      ) {
        favouriteItems = JSON.parse(favouriteLinkListItem[0]['PzlPersonalLinks'])
        favouriteItemsIds = favouriteItems.map((link) => link.id)
        this.setState({
          favouriteLinks: favouriteItems
        })
      } else {
        favouriteItemsIds = []
        this.setState({
          isFirstUpdate: true
        })
      }

      console.log(this.props)

      const editorLinks = await sp.web
        .getList(this.props.webServerRelativeUrl + '/Lists/EditorLinks')
        .items.filter('PzlLinkActive eq 1')
        .get()

      const mappedLinks: Link[] = editorLinks.map((link) => {
        return {
          id: link.Id,
          displayText: link.Title,
          url: link.PzlUrl,
          icon: link.PzlOfficeUIFabricIcon,
          priority: link.PzlPriority,
          mandatory: link.PzlLinkMandatory,
          linkType: LinkType.editorLink
        }
      })
      const mandatorymappedLinks: Link[] = mappedLinks.filter((link) => link.mandatory)

      const recommendedmappedLinks: Link[] = mappedLinks.filter((link) => !link.mandatory)

      const prunedLinks: Link[] = recommendedmappedLinks.filter(
        (link) => !favouriteItemsIds.includes(link.id)
      )
      if (
        favouriteLinkListItem.length > 0 &&
        favouriteItems !== null &&
        favouriteItems.length > 0
      ) {
        const favoriteLinks: Link[] = await this.checkForUpdatedLinks(
          favouriteItems,
          recommendedmappedLinks
        )
        favouriteItemsIds = favoriteLinks.map((item: Link): number => item.id)
      }
      const linkFieldId = favouriteLinkListItem.length > 0 ? favouriteLinkListItem[0].Id : null
      const currentUser: User = {
        id: this.props.currentUserId,
        linkFieldId: linkFieldId
      }

      const displayLinks = editorLinks.map((link) => {
        return {
          displayText: link.Title,
          url: link.PzlUrl,
          icon: link.PzlOfficeUIFabricIcon || 'Link',
          priority: link.PzlLinkPriority || '0',
          category: link.PzlLinkCategory || 'Ingen kategori',
          mandatory: link.PzlLinkMandatory,
          linkType: LinkType.editorLink
        }
      })

      let categories: Array<ICategory> = [
        { displayText: strings.NoCategoryLabel, links: displayLinks }
      ]
      if (this.props.listingByCategory) {
        let categoryNames: string[] = displayLinks
          .map((lnk) => {
            return lnk.category
          })
          .sort()
        categoryNames = categoryNames.filter((item, index) => {
          return categoryNames.indexOf(item) == index
        })
        categories = categoryNames.map((catName) => {
          return {
            displayText: catName,
            links: displayLinks.filter((lnk) => {
              return lnk.category === catName
            })
          }
        })
      }

      this.setState({
        currentUser: currentUser,
        editorLinks: prunedLinks,
        mandatoryLinks: mandatorymappedLinks,
        categoryLinks: categories,
        isLoading: false
      })
    } catch (err) {
      console.log(err)
      this.setState({
        isLoading: false
      })
    }
  }

  private async saveData() {
    this.setState({
      showLoadingSpinner: true,
      saveButtonDisabled: true
    })
    try {
      const linksAsString: string = JSON.stringify(this.state.favouriteLinks)
      if (this.state.isFirstUpdate) {
        const result: ItemAddResult = await sp.web
          .getList(this.props.webServerRelativeUrl + '/Lists/FavouriteLinks')
          .items.add({
            PzlPersonalLinks: linksAsString,
            Title: this.props.currentUserName
          })
        const currentUser: User = {
          id: this.state.currentUser.id,
          linkFieldId: result.data.Id
        }
        this.setState({
          isFirstUpdate: false,
          saveButtonDisabled: true,
          currentUser: currentUser,
          showSuccessMessage: true,
          showLoadingSpinner: false
        })
        setTimeout((): void => this.setState({ showSuccessMessage: false }), 5000)
      } else {
        await sp.web
          .getList(this.props.webServerRelativeUrl + '/Lists/FavouriteLinks')
          .items.getById(+this.state.currentUser.linkFieldId)
          .update({
            PzlPersonalLinks: linksAsString
          })
        this.setState({
          saveButtonDisabled: true,
          showSuccessMessage: true,
          showLoadingSpinner: false
        })
        setTimeout((): void => this.setState({ showSuccessMessage: false }), 5000)
      }
    } catch (err) {
      this.setState({
        showErrorMessage: true,
        showLoadingSpinner: false,
        saveButtonDisabled: false
      })
      setTimeout((): void => this.setState({ showErrorMessage: false }), 5000)
    }
  }

  private async checkForUpdatedLinks(userFavoriteLinks: any[], allFavoriteLinks: any[]) {
    const personalLinks: Link[] = new Array<Link>()
    let shouldUpdate: boolean = false
    userFavoriteLinks.forEach((userLink): void => {
      const linkMatch = allFavoriteLinks.find((favoriteLink) => favoriteLink.id === userLink.id)
      if (
        linkMatch &&
        (!isEqual(linkMatch.url, userLink.url) ||
          !isEqual(linkMatch.displayText, userLink.displayText) ||
          !isEqual(linkMatch.icon, userLink.icon))
      ) {
        shouldUpdate = true
        personalLinks.push(linkMatch)
      } else {
        personalLinks.push(userLink)
      }
    })

    if (shouldUpdate) {
      this.setState({
        favouriteLinks: personalLinks
      })
    }
    return personalLinks
  }
}
