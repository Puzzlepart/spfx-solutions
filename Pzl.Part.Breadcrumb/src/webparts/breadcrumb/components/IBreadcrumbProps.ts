import { SPListItem } from "@microsoft/sp-page-context";

export interface IBreadcrumbProps {
  description: string;
  listServerRelativeUrl: string;
  lookupField: string;
  currentPage: SPListItem;
}
