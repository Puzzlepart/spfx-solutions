import { IReadonlyTheme } from '@microsoft/sp-component-base'

export interface IQuickLinksProps {
  theme: IReadonlyTheme
  title: string
  description: string
  userId: number
  allLinksUrl: string
  defaultIcon: string
  groupByCategory: boolean
  lineHeight: number
  gapSize: number
  iconsOnly: boolean
  iconSize: number
  renderShadow: boolean
  responsiveButtons: boolean
  webServerRelativeUrl: string
  linkClickWebHook: string
  hideHeader: boolean
  hideTitle: boolean
  hideShowAll: boolean
  allLinksText: string
  buttonAppearance: 'secondary' | 'primary' | 'outline' | 'subtle' | 'transparent'
  globalConfigurationUrl: string
  context: any
}

export interface IQuickLinksState {
  linkStructure: Array<ICategory>
}

export interface ILink {
  id?: number
  displayText: string
  url: string
  icon: string
  category: string
  priority: string
  openInSameTab: boolean
}
export interface ICategory {
  links: Array<ILink>
  displayText: string
}
