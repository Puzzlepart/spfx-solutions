import { INavLink } from "office-ui-fabric-react/lib/Nav";

export interface IPageNavigationState {
  rootNode: INavLink;
  isLoading?: boolean;
  pages: Array<any>;
}
  