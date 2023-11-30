import { IReadonlyTheme } from '@microsoft/sp-component-base'

export interface IAllLinksProps {
  theme: IReadonlyTheme
  currentUserId: number
  currentUserName: string
  defaultIcon: string
  webServerRelativeUrl: string
  mylinksOnTop: boolean
  listingByCategory: boolean
  listingByCategoryTitle: string
  mandatoryLinksTitle: string
  recommendedLinksTitle: string
  myLinksTitle: string
}
