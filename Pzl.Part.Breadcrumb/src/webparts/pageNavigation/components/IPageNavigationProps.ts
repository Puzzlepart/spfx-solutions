import { SPListItem } from "@microsoft/sp-page-context";

export interface IPageNavigationProps {
  listServerRelativeUrl: string;
  lookupField: string;
  topLevelPage: number;
  serverRequestPath: string;
  currentPage: SPListItem;
}
