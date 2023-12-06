import { IReadonlyTheme } from '@microsoft/sp-component-base'

export interface IAllLinksProps {
  theme: IReadonlyTheme
  currentUserId: number
  currentUserName: string
  defaultIcon: string
  webServerRelativeUrl: string
  groupByCategory: boolean
  mandatoryLinksTitle: string
  mandatoryLinksDescription: string
  recommendedLinksTitle: string
  recommendedLinksDescription: string
  yourLinksTitle: string
  yourLinksDescription: string
}

export interface IAllLinksState {
  editorLinks?: Array<ILink>
  favouriteLinks?: Array<ILink>
  mandatoryLinks?: Array<ILink>
  categoryLinks?: Array<ICategory>
  loading?: boolean
  error?: boolean
  validationError?: boolean
  showSuccessMessage?: boolean
  currentUser?: User
  showDialog?: boolean
  dialogData?: ILink
  isFirstUpdate?: boolean
  saveButtonDisabled?: boolean
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
  category?: string
  priority?: string
  mandatory?: boolean
  linkType?: LinkType
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
