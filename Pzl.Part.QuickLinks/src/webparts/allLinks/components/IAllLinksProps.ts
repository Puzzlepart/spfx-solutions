import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IAllLinksProps {
  currentUserId: number;
  currentUserName: string;
  defaultIcon: string;
  webServerRelativeUrl: string;
  mylinksOnTop: boolean;
  listingByCategory: boolean;
  listingByCategoryTitle: string;
  maxLinkLength: number;
}


