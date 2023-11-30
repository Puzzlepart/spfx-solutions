import { IReadonlyTheme } from '@microsoft/sp-component-base'

export interface IAllLinksProps {
  theme: IReadonlyTheme
  currentUserId: number
  currentUserName: string
  defaultIcon: string
  webServerRelativeUrl: string
  yourLinksOnTop: boolean
  listingByCategory: boolean
  listingByCategoryTitle: string
  mandatoryLinksTitle: string
  recommendedLinksTitle: string
  yourLinksTitle: string
}

export interface IAllLinksState {
  editorLinks?: Array<Link>
  favouriteLinks?: Array<Link>
  mandatoryLinks?: Array<Link>
  categoryLinks?: Array<ICategory>
  loading?: boolean
  error?: boolean
  validationError?: boolean
  showSuccessMessage?: boolean
  currentUser?: User
  showDialog?: boolean
  dialogData?: Link
  isFirstUpdate?: boolean
  saveButtonDisabled?: boolean
}

export interface Link {
  id?: number
  displayText: string
  url: string
  icon?: string
  priority?: string
  mandatory?: number
  category?: string
  linkType: LinkType
}

export enum LinkType {
  editorLink = 'EditorLink',
  favouriteLinks = 'FavouriteLink',
  mandatoryLinks = 'MandatoryLinks'
}

export interface User {
  id: number
  linkFieldId?: string
}

export interface ILink {
  id?: number
  displayText: string
  url: string
  icon: string
  category: string
  priority: string
  mandatory?: number
  linkType: LinkType
  openInSameTab?: boolean
}
export interface ICategory {
  links: Array<ILink>
  displayText: string
}

export enum CategoryOperation {
  add,
  remove
}
