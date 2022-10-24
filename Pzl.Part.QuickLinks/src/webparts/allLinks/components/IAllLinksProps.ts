import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IAllLinksProps {
  theme: IReadonlyTheme,
  currentUserId: number;
  currentUserName: string;
  defaultIcon: string;
  webServerRelativeUrl: string;
  mylinksOnTop: boolean;
  listingByCategory: boolean;
  listingByCategoryTitle: string;
  maxLinkLength: number;
  iconOpacity: number;
  mandatoryLinksTitle: string;
  reccomendedLinksTitle: string;
  myLinksTitle: string;
}
