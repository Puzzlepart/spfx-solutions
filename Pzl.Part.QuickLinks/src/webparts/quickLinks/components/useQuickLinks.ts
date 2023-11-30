import { sp } from '@pnp/sp'
import { IReadonlyTheme } from '@microsoft/sp-component-base'
import { isEqual } from 'lodash'
import { stringIsNullOrEmpty } from '@pnp/common'
import { useEffect } from 'react'
import { useQuickLinksState } from './useQuickLinksState'
import { ICategory, ILink, IQuickLinksProps } from './types'
import strings from 'QuickLinksWebPartStrings'

/**
 * Component logic hook for `quickLinks`. This hook is responsible for
 * fetching quickLinks
 *
 * @param props Props for `QuickLinks` component
 */
export const useQuickLinks = (props: IQuickLinksProps) => {
  const { state, setState } = useQuickLinksState()

  const theme: IReadonlyTheme = props.theme
  const backgroundColor: string = theme?.semanticColors?.bodyBackground ?? '#ffffff'

  useEffect(() => {
    fetchData()
  }, [])

  const fetchData = async () => {
    const searchString: string = `AuthorId eq '${props.userId}'`

    const editorLinks = await sp.web
      .getList(props.webServerRelativeUrl + '/Lists/EditorLinks')
      .items.filter('(PzlLinkActive eq 1) and (PzlLinkMandatory eq 1)')
      .orderBy('PzlLinkPriority')
      .orderBy('Title')
      .get()

    const newNonMandatoryLinks = await sp.web
      .getList(props.webServerRelativeUrl + '/Lists/EditorLinks')
      .items.filter('(PzlLinkActive eq 1) and (PzlLinkMandatory eq 0)')
      .orderBy('PzlLinkPriority')
      .orderBy('Title')
      .get()

    const newNonMandatoryLinksObject = newNonMandatoryLinks.map((link) => {
      return {
        id: link.Id,
        displayText: link.Title,
        url: link.PzlUrl,
        icon: link.PzlOfficeUIFabricIcon,
        priority: link.PzlLinkPriority,
        category: link.PzlLinkCategory || 'Ingen kategori',
        openInSameTab: link.PzlOpenInSameTab
      }
    })

    const favouriteLinkStrings = await sp.web
      .getList(props.webServerRelativeUrl + '/Lists/FavouriteLinks')
      .items.select('Id', 'AuthorId', 'PzlPersonalLinks')
      .filter(searchString)
      .get()

    const favouriteLinksObject: ILink[] =
      favouriteLinkStrings.length > 0 ? JSON.parse(favouriteLinkStrings[0].PzlPersonalLinks) : []

    const displayLinks = editorLinks.map((link) => {
      return {
        displayText: link.Title,
        url: link.PzlUrl,
        icon: link.PzlOfficeUIFabricIcon || 'Link',
        priority: link.PzlLinkPriority || '0',
        category: link.PzlLinkCategory || 'Ingen kategori',
        openInSameTab: link.PzlOpenInSameTab
      }
    })

    if (favouriteLinkStrings.length > 0) {
      const updatedFavoriteLinksObject = await checkForUpdatedLinks(
        favouriteLinksObject,
        newNonMandatoryLinksObject,
        favouriteLinkStrings[0].Id
      )
      displayLinks.push(...updatedFavoriteLinksObject)
    }

    let categories: Array<ICategory> = [
      { displayText: strings.NoCategoryLabel, links: displayLinks }
    ]
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

    setState({ linkStructure: categories })
  }

  const checkForUpdatedLinks = async (
    userFavoriteLinks: ILink[],
    allFavoriteLinks: ILink[],
    currentItemId: number
  ) => {
    const personalLinks: ILink[] = new Array<ILink>()
    let shouldUpdate: boolean = false
    userFavoriteLinks.forEach((userLink: ILink): void => {
      const linkMatch: ILink = allFavoriteLinks.find(
        (favoriteLink) => favoriteLink.id === userLink.id
      )
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
      await updatePersonalLinks(personalLinks, currentItemId)
    }
    return personalLinks
  }

  const updatePersonalLinks = async (newFavoriteLinks, itemId: number) => {
    try {
      await sp.web
        .getList(props.webServerRelativeUrl + '/Lists/FavouriteLinks')
        .items.getById(itemId)
        .update({
          PzlPersonalLinks: JSON.stringify(newFavoriteLinks)
        })
    } catch (e) {
      console.log(e)
    }
  }

  const callWebHook = (uri: string, category: string): Promise<any> => {
    if (stringIsNullOrEmpty(props.linkClickWebHook)) {
      return
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

    fetch(props.linkClickWebHook, postRequest)
  }

  return {
    state,
    callWebHook,
    backgroundColor
  }
}