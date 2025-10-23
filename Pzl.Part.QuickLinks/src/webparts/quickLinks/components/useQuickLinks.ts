import { getSP } from '../../../util/spContext'
import { isEqual } from 'lodash'
import { stringIsNullOrEmpty } from '@pnp/common'
import { useEffect } from 'react'
import { useQuickLinksState } from './useQuickLinksState'
import { ICategory, ILink, IQuickLinksProps } from './types'
import strings from 'QuickLinksWebPartStrings'
import tinycolor from 'tinycolor2'
import { customDarkTheme, customLightTheme } from '../../../util/theme'

/**
 * Component logic hook for `quickLinks`. This hook is responsible for
 * fetching quickLinks
 *
 * @param props Props for `QuickLinks` component
 */
export const useQuickLinks = (props: IQuickLinksProps) => {
  const { state, setState } = useQuickLinksState()
  const sp = getSP(props.context, props.globalConfigurationUrl)

  const backgroundColor: string = props.theme?.semanticColors?.bodyBackground ?? '#ffffff'
  const theme = tinycolor(backgroundColor).isDark() ? customDarkTheme : customLightTheme

  useEffect(() => {
    if (!sp) return
    fetchData()
  }, [sp])

  const fetchData = async (): Promise<void> => {
    let webServerRelativeUrl: string = props.webServerRelativeUrl
    let searchString: string = `AuthorId eq '${props.userId}'`

    if (props.globalConfigurationUrl) {
      const web = await sp.web()
      const user = await sp.web.currentUser()
      const userId = user.Id
      const webUrl = web.Url
      const sitesIndex = webUrl.indexOf('/sites/')
      searchString = `AuthorId eq '${userId}'`

      if (sitesIndex > -1) {
        webServerRelativeUrl = webUrl.substring(sitesIndex)
      } else {
        webServerRelativeUrl = ''
      }
    }

    const editorLinks = await sp.web
      .getList(`${webServerRelativeUrl}/Lists/EditorLinks`)
      .items.filter('(PzlLinkActive eq 1) and (PzlLinkMandatory eq 1)')
      .orderBy('PzlLinkPriority')
      .orderBy('Title')()
    const newNonMandatoryLinks = await sp.web
      .getList(`${webServerRelativeUrl}/Lists/EditorLinks`)
      .items.filter('(PzlLinkActive eq 1) and (PzlLinkMandatory eq 0)')
      .orderBy('PzlLinkPriority')
      .orderBy('Title')()

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
      .getList(`${webServerRelativeUrl}/Lists/FavouriteLinks`)
      .items.select('Id', 'AuthorId', 'PzlPersonalLinks')
      .filter(searchString)()

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
      const updatedFavouriteLinksObject = await checkForUpdatedLinks(
        favouriteLinksObject,
        newNonMandatoryLinksObject,
        favouriteLinkStrings[0].Id
      )
      displayLinks.push(...updatedFavouriteLinksObject)
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
    userFavouriteLinks: ILink[],
    allFavouriteLinks: ILink[],
    currentItemId: number
  ) => {
    const personalLinks: ILink[] = new Array<ILink>()
    let shouldUpdate: boolean = false
    userFavouriteLinks.forEach((userLink: ILink): void => {
      const linkMatch: ILink = allFavouriteLinks.find(
        (favouriteLink) => favouriteLink.id === userLink.id
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

  const updatePersonalLinks = async (newFavouriteLinks, itemId: number) => {
    try {
      await sp.web
        .getList(props.webServerRelativeUrl + '/Lists/FavouriteLinks')
        .items.getById(itemId)
        .update({
          PzlPersonalLinks: JSON.stringify(newFavouriteLinks)
        })
    } catch (e) {
      console.log(e)
    }
  }

  const callWebHook = (id: number, uri: string, category: string): Promise<any> => {
    if (stringIsNullOrEmpty(props.linkClickWebHook)) {
      return
    }

    const body = {
      id: id,
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
    backgroundColor,
    theme
  }
}
