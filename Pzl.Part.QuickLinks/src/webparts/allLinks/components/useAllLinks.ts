import { getSP } from '../../../util/spContext'
import { useAllLinksState } from './useAllLinksState'
import { CategoryOperation, IAllLinksProps, ICategory, ILink, LinkType, User } from './types'
import { useEffect } from 'react'
import strings from 'AllLinksWebPartStrings'
import { isEqual } from '@microsoft/sp-lodash-subset'
import tinycolor from 'tinycolor2'
import { customDarkTheme, customLightTheme } from '../../../util/theme'
import { PermissionKind } from '@pnp/sp/security'
import '@pnp/sp/security'

/**
 * Component logic hook for `allLinks`. This hook is responsible for
 * fetching all the links
 *
 * @param props Props for `AllLinks` component
 */
export const useAllLinks = (props: IAllLinksProps) => {
  const { state, setState } = useAllLinksState()
  const sp = getSP(props.context)

  const getErrorMessage = (error: unknown): string => {
    if (error instanceof Error && error.message) {
      return error.message
    }

    if (typeof error === 'object' && error !== null) {
      const message = Reflect.get(error, 'message')
      if (typeof message === 'string' && message) {
        return message
      }

      const data = Reflect.get(error, 'data')
      if (typeof data === 'object' && data !== null) {
        const dataMessage = Reflect.get(data, 'message')
        if (typeof dataMessage === 'string' && dataMessage) {
          return dataMessage
        }
      }
    }

    return strings.SaveErrorLabel
  }

  const getAddedItemId = async (result: {
    Id?: number
    ID?: number
    data?: { Id?: number; ID?: number }
    item?: { select?: (fields: string) => () => Promise<{ Id?: number; ID?: number }>; (): Promise<{ Id?: number; ID?: number }> }
  }): Promise<number> => {
    const idFromResult = result?.Id ?? result?.ID
    if (typeof idFromResult === 'number') {
      return idFromResult
    }

    const idFromData = result?.data?.Id ?? result?.data?.ID
    if (typeof idFromData === 'number') {
      return idFromData
    }

    if (result?.item) {
      const addedItem = typeof result.item.select === 'function' ? await result.item.select('Id')() : await result.item()
      const idFromItem = addedItem?.Id ?? addedItem?.ID
      if (typeof idFromItem === 'number') {
        return idFromItem
      }
    }

    throw new Error('Could not resolve created list item ID.')
  }

  const backgroundColor: string = props.theme?.semanticColors?.bodyBackground ?? '#ffffff'
  const theme = tinycolor(backgroundColor).isDark() ? customDarkTheme : customLightTheme

  useEffect(() => {
    if (!sp) return
    fetchData()
  }, [sp])

  const openNewLinkDialog = (): void => {
    const emptyLink: ILink = {
      localKey: `personal-${Date.now()}`,
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

  const addNewLink = (): ILink => {
    const createdLink: ILink = {
      ...state.dialogData,
      localKey: state.dialogData?.localKey || `personal-${Date.now()}`,
      linkType: LinkType.favouriteLinks
    }
    const newFavourites: ILink[] = state.favouriteLinks.slice()
    newFavourites.push(createdLink)
    saveData(newFavourites)
    setState({
      favouriteLinks: newFavourites,
      dialogData: null,
      showDialog: false,
      saveButtonDisabled: false
    })

    return createdLink
  }

  const addEditorLink = async (link: ILink): Promise<ILink | null> => {
    try {
      const normalizedCategory = link.category?.trim() || ''
      const normalizedIcon = link.icon || props.defaultIcon
      const normalizedPriority = link.priority?.trim() || '1000'
      const normalizedLink: ILink = {
        ...link,
        category: normalizedCategory,
        icon: normalizedIcon,
        priority: normalizedPriority,
        active: link.active ?? true,
        mandatory: !!link.mandatory,
        linkType: LinkType.editorLink
      }

      const result = await sp.web
        .getList(props.webServerRelativeUrl + '/Lists/EditorLinks')
        .items.add({
          Title: normalizedLink.displayText,
          PzlUrl: normalizedLink.url,
          PzlOfficeUIFabricIcon: normalizedLink.icon,
          PzlLinkCategory: normalizedLink.category,
          PzlLinkPriority: Number(normalizedLink.priority),
          PzlLinkMandatory: normalizedLink.mandatory,
          PzlLinkActive: normalizedLink.active
        })
      const createdItemId = await getAddedItemId(result)

      const createdLink: ILink = {
        ...normalizedLink,
        id: createdItemId
      }

      const nextEditorLinks = createdLink.mandatory
        ? state.editorLinks ?? []
        : [...(state.editorLinks ?? []), createdLink]

      const nextMandatoryLinks = createdLink.mandatory
        ? [...(state.mandatoryLinks ?? []), createdLink]
        : state.mandatoryLinks ?? []

      const nextCategoryLinks = props.groupByCategory
        ? (() => {
            const categoryName = createdLink.category?.trim() || strings.NoCategoryLabel
            const existingCategory = (state.categoryLinks ?? []).find(
              (category) => category.displayText === categoryName
            )

            if (existingCategory) {
              return (state.categoryLinks ?? []).map((category) => {
                if (category.displayText !== categoryName) {
                  return category
                }

                return {
                  ...category,
                  links: [...category.links, createdLink]
                }
              })
            }

            return [
              ...(state.categoryLinks ?? []),
              {
                displayText: categoryName,
                links: [createdLink]
              }
            ]
          })()
        : state.categoryLinks

      const nextCategoryOptions = createdLink.category?.trim()
        ? Array.from(new Set([...(state.categoryOptions ?? []), createdLink.category.trim()])).sort(
            (left, right) => left.localeCompare(right)
          )
        : state.categoryOptions ?? []

      setState({
        editorLinks: nextEditorLinks,
        mandatoryLinks: nextMandatoryLinks,
        categoryLinks: nextCategoryLinks,
        categoryOptions: nextCategoryOptions,
        error: false,
        errorMessage: ''
      })

      return createdLink
    } catch (error) {
      const errorMessage = getErrorMessage(error)
      console.error('AllLinks addEditorLink failed', error)

      setState({
        error: true,
        errorMessage
      })

      setTimeout((): void => setState({ error: false, errorMessage: '' }), 10000)
      return null
    }
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
      const editorLinksList = sp.web.getList(props.webServerRelativeUrl + '/Lists/EditorLinks')
      const canManageEditorLinks = await editorLinksList.currentUserHasPermissions(
        PermissionKind.AddListItems
      )

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

      const editorLinks = await editorLinksList.items.filter('PzlLinkActive eq 1')()
      const allEditorLinkCategories = await editorLinksList.items.select('PzlLinkCategory').top(5000)()

      const categoryOptions = Array.from(
        new Set(
          allEditorLinkCategories
            .map((link) => String(link.PzlLinkCategory || '').trim())
            .filter((category) => !!category)
        )
      ).sort((left, right) => left.localeCompare(right))

      const mappedLinks: ILink[] = editorLinks.map((link) => {
        return {
          id: link.Id,
          displayText: link.Title,
          url: link.PzlUrl,
          icon: link.PzlOfficeUIFabricIcon,
          priority: link.PzlPriority,
          mandatory: link.PzlLinkMandatory,
          active: link.PzlLinkActive,
          category: link.PzlLinkCategory || '',
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
          id: link.Id,
          displayText: link.Title,
          url: link.PzlUrl,
          icon: link.PzlOfficeUIFabricIcon || 'Link',
          priority: link.PzlLinkPriority || '0',
          category: link.PzlLinkCategory || strings.NoCategoryLabel,
          mandatory: link.PzlLinkMandatory,
          active: link.PzlLinkActive,
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
        canManageEditorLinks,
        categoryOptions,
        editorLinks: prunedLinks,
        mandatoryLinks: mandatorymappedLinks,
        categoryLinks: categories,
        errorMessage: '',
        loading: false
      })
    } catch (error) {
      console.error('AllLinks fetchData failed', error)
      setState({
        error: true,
        errorMessage: getErrorMessage(error),
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
        const createdItemId = await getAddedItemId(result)

        const currentUser: User = {
          id: state.currentUser.id,
          linkFieldId: String(createdItemId)
        }

        setState({
          isFirstUpdate: false,
          saveButtonDisabled: true,
          currentUser: currentUser,
          error: false,
          errorMessage: '',
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
          error: false,
          errorMessage: '',
          showSuccessMessage: true,
          loading: false
        })

        setTimeout((): void => setState({ showSuccessMessage: false }), 5000)
      }
    } catch (error) {
      const errorMessage = getErrorMessage(error)
      console.error('AllLinks saveData failed', error)

      setState({
        error: true,
        errorMessage,
        loading: false,
        saveButtonDisabled: false
      })

      setTimeout((): void => setState({ error: false, errorMessage: '' }), 10000)
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
    addEditorLink,
    onDialogValueChanged,
    validateUrl,
    theme
  }
}
