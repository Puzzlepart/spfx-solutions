import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IQuickLinksProps {
  theme: IReadonlyTheme;
  title: string;
  userId: number;
  numberOfLinks: number;
  allLinksUrl: string;
  defaultIcon: string;
  groupByCategory: boolean;
  maxLinkLength: number;
  lineHeight: number;
  iconOpacity: number;
  webServerRelativeUrl: string;
  linkClickWebHook: string;
}
