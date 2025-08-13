import { getSP } from '../../pnpjsConfig'
import { useAllLinksState } from './useAllLinksState'
import { CategoryOperation, IAllLinksProps, ICategory, ILink, LinkType, User } from './types'
import { useEffect } from 'react'
import strings from 'AllLinksWebPartStrings'
import { isEqual } from '@microsoft/sp-lodash-subset'
import tinycolor from 'tinycolor2'
import { customDarkTheme, customLightTheme } from '../../../util/theme'
import { WebPartContext } from '@microsoft/sp-webpart-base'

/**
 * Component logic hook for `allLinks`. This hook is responsible for
 * fetching all the links
 *
 * @param props Props for `AllLinks` component
 */
export const useAllLinks = (props: IAllLinksProps) => {
  const { state, setState } = useAllLinksState()
  const sp = getSP(props.context)

  const backgroundColor: string = props.theme?.semanticColors?.bodyBackground ?? '#ffffff'
  const theme = tinycolor(backgroundColor).isDark() ? customDarkTheme : customLightTheme

  useEffect(() => {
    if (!sp) return
    fetchData()
  }, [sp])

  const openNewLinkDialog = (): void => {
    const emptyLink: ILink = {
      id: -1,
      displayText: '',
      url: '',
      icon: props.defaultIcon,
      priority: '1000',
      mandatory: false,
      linkType: LinkType.favouriteLinks
    }

    setState({
      showDialog: true,
      dialogData: emptyLink
    })
  }

  const appendToFavourites = (link: ILink): void => {
    const newFavourites: ILink[] = state.favouriteLinks.slice()
    newFavourites.push(link)

    const newEditorLinks: ILink[] = state.editorLinks.slice()
    newEditorLinks.splice(newEditorLinks.indexOf(link), 1)

    updateCategoryLinks(CategoryOperation.remove, link as ILink, state.categoryLinks)
    saveData(newFavourites)
    setState({
      favouriteLinks: newFavourites,
      editorLinks: newEditorLinks,
      saveButtonDisabled: false
    })
  }

  const removeFromFavourites = (link: ILink): void => {
    const newEditorLinks: ILink[] = state.editorLinks.slice()
    newEditorLinks.push(link)

    const newFavourites: ILink[] = state.favouriteLinks.slice()
    newFavourites.splice(newFavourites.indexOf(link), 1)

    const categoryLinks: ICategory[] = updateCategoryLinks(
      CategoryOperation.add,
      link as ILink,
      state.categoryLinks
    )
    saveData(newFavourites)
    setState({
      favouriteLinks: newFavourites,
      editorLinks: newEditorLinks,
      categoryLinks: categoryLinks,
      saveButtonDisabled: false
    })
  }

  const updateCategoryLinks = (
    operation: CategoryOperation,
    link: ILink,
    categoryLinks: ICategory[]
  ): ICategory[] => {
    if (props.groupByCategory) {
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

  const removeCustomFromFavourites = (link: ILink): void => {
    const newFavourites: ILink[] = state.favouriteLinks.slice()
    newFavourites.splice(newFavourites.indexOf(link), 1)
    saveData(newFavourites)
    setState({
      favouriteLinks: newFavourites,
      saveButtonDisabled: false
    })
  }

  const addNewLink = (): void => {
    const newFavourites: ILink[] = state.favouriteLinks.slice()
    newFavourites.push(state.dialogData)
    saveData(newFavourites)
    setState({
      favouriteLinks: newFavourites,
      dialogData: null,
      showDialog: false,
      saveButtonDisabled: false
    })
  }

  const onDialogValueChanged = (field: string, newVal: any): void => {
    const newDialogData: ILink = { ...state.dialogData }
    newDialogData[field] = newVal
    setState({ dialogData: newDialogData })
  }

  const validateUrl = (value: any) => {
    if (value.length > 0) {
      const urlRegex: RegExp =
        /(http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-]*[\w@?^=%&/~+#-])?/
      value.match(urlRegex) === null
        ? setState({ validationError: true })
        : setState({ validationError: false })
    } else {
      setState({ validationError: false })
    }
  }

  const fetchData = async (): Promise<void> => {
    try {
      const searchString: string = `AuthorId eq '${props.currentUserId}'`
      const favouriteLinkListItem = await sp.web
        .getList(props.webServerRelativeUrl + '/Lists/FavouriteLinks')
        .items.select('Id', 'AuthorId', 'PzlPersonalLinks')
        .filter(searchString)()
      let favouriteItemsIds: number[]
      let favouriteItems: ILink[] = []
      if (
        favouriteLinkListItem.length > 0 &&
        favouriteLinkListItem[0]['PzlPersonalLinks'] !== null
      ) {
        favouriteItems = JSON.parse(favouriteLinkListItem[0]['PzlPersonalLinks'])
        favouriteItemsIds = favouriteItems.map((link) => link.id)
        setState({
          favouriteLinks: favouriteItems
        })
      } else {
        favouriteItemsIds = []
        setState({
          isFirstUpdate: true
        })
      }

      const editorLinks = await sp.web
        .getList(props.webServerRelativeUrl + '/Lists/EditorLinks')
        .items.filter('PzlLinkActive eq 1')()

      const mappedLinks: ILink[] = editorLinks.map((link) => {
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
      const mandatorymappedLinks: ILink[] = mappedLinks.filter((link) => link.mandatory)

      const recommendedmappedLinks: ILink[] = mappedLinks.filter((link) => !link.mandatory)

      const prunedLinks: ILink[] = recommendedmappedLinks.filter(
        (link) => !favouriteItemsIds.includes(link.id)
      )
      if (
        favouriteLinkListItem.length > 0 &&
        favouriteItems !== null &&
        favouriteItems.length > 0
      ) {
        const favouriteLinks: ILink[] = await checkForUpdatedLinks(
          favouriteItems,
          recommendedmappedLinks
        )
        favouriteItemsIds = favouriteLinks.map((item: ILink): number => item.id)
      }
      const linkFieldId = favouriteLinkListItem.length > 0 ? favouriteLinkListItem[0].Id : null
      const currentUser: User = {
        id: props.currentUserId,
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

      if (props.groupByCategory) {
        let categoryNames: string[] = displayLinks
          .map((lnk) => {
            return lnk.category
          })
          .sort()
        categoryNames = categoryNames.filter((item, index) => {
          return categoryNames.indexOf(item) === index
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

      setState({
        currentUser: currentUser,
        editorLinks: prunedLinks,
        mandatoryLinks: mandatorymappedLinks,
        categoryLinks: categories,
        loading: false
      })
    } catch (err) {
      console.log(err)
      setState({
        loading: false
      })
    }
  }

  const saveData = async (favouriteLinks?: Array<ILink>) => {
    setState({
      saveButtonDisabled: true
    })
    try {
      const linksAsString: string = JSON.stringify(favouriteLinks)
      if (state.isFirstUpdate) {
        const result = await sp.web
          .getList(props.webServerRelativeUrl + '/Lists/FavouriteLinks')
          .items.add({
            PzlPersonalLinks: linksAsString,
            Title: props.currentUserName
          })

        const currentUser: User = {
          id: state.currentUser.id,
          linkFieldId: result.data.Id
        }

        setState({
          isFirstUpdate: false,
          saveButtonDisabled: true,
          currentUser: currentUser,
          showSuccessMessage: true,
          loading: false
        })

        setTimeout((): void => setState({ showSuccessMessage: false }), 5000)
      } else {
        await sp.web
          .getList(props.webServerRelativeUrl + '/Lists/FavouriteLinks')
          .items.getById(+state.currentUser.linkFieldId)
          .update({
            PzlPersonalLinks: linksAsString
          })

        setState({
          saveButtonDisabled: true,
          showSuccessMessage: true,
          loading: false
        })

        setTimeout((): void => setState({ showSuccessMessage: false }), 5000)
      }
    } catch (err) {
      setState({
        error: true,
        loading: false,
        saveButtonDisabled: false
      })

      setTimeout((): void => setState({ error: false }), 5000)
    }
  }

  const checkForUpdatedLinks = (userFavouriteLinks: any[], allFavouriteLinks: any[]) => {
    const personalLinks: ILink[] = new Array<ILink>()
    let shouldUpdate: boolean = false
    userFavouriteLinks.forEach((userLink): void => {
      const linkMatch = allFavouriteLinks.find((favouriteLink) => favouriteLink.id === userLink.id)
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
      setState({
        favouriteLinks: personalLinks
      })
    }
    return personalLinks
  }

  return {
    state,
    setState,
    backgroundColor,
    openNewLinkDialog,
    appendToFavourites,
    removeFromFavourites,
    removeCustomFromFavourites,
    addNewLink,
    onDialogValueChanged,
    validateUrl,
    theme
  }
}
