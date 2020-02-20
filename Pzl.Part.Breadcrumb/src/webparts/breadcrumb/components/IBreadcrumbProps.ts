import { SPListItem } from "@microsoft/sp-page-context";

/**
 * Breadcrumb component properties
 */
export interface IBreadcrumbProps {
  description: string;
  listServerRelativeUrl: string;
  lookupField: string;
  currentPage: SPListItem;
}
