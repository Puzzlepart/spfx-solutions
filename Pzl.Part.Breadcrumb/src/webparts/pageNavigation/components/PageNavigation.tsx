import * as React from 'react';
import styles from './PageNavigation.module.scss';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { IPageNavigationProps } from './IPageNavigationProps';
import { IPageNavigationState } from './IPageNavigationState';
import { sp } from '@pnp/sp';
import { IODataListItem } from '@microsoft/sp-odata-types';
import { Nav, INavLink, INavStyles, INavStyleProps } from 'office-ui-fabric-react/lib/Nav';

export interface IPage extends IODataListItem {
  FileRef: string;
}

/**
 * 
 */
export default class PageNavigation extends React.Component<IPageNavigationProps, IPageNavigationState> {
  constructor(props) {
    super(props);
    this.state = ({
      rootNode: null,
      isLoading: true,
      pages: []
    });
  }

  /**
   * 
   * @returns {React.ReactElement<IPageNavigationProps>}
   */
  public render(): React.ReactElement<IPageNavigationProps> {
    let { rootNode } = this.state;
    if (!rootNode) {
      return null;
    }
    if (this.state.isLoading) {
      return <Spinner size={SpinnerSize.large} />;
    }
    rootNode.isExpanded = true;
    return (
      <div>
        <Nav
          styles={() => {
            return {
              chevronButton: {
                selectors: {
                  '&:after': {
                    borderLeft: "none",
                    content: '""',
                  }
                }
              }
            };
          }}
          groups={[{ links: [ rootNode ] }]}
        />
      </div>
    );
  }

  /**
   * 
   * @returns {Promise<void>}
   */
  public async componentDidMount(): Promise<void> {
    await this.fetchListItems();
    await this.buildPageNavigation();
  }

  /**
   * 
   * @param {*} prevprops
   * @returns {Promise<void>}
   */
  public async componentDidUpdate(prevprops): Promise<void> {
    if (this.props.topLevelPage !== prevprops.topLevelPage) {
      await this.buildPageNavigation();
    }
  }

  /**
   * 
   */
  private async buildPageNavigation() {
    try {
      let currentPage = await sp.web.getList(this.props.listServerRelativeUrl).items.getById(this.props.topLevelPage).select("Id", "Title", "FileRef").get();
      let currentPageNode: INavLink = { key: currentPage.Id, name: currentPage.Title, url: currentPage.FileRef };
      let navigationTree = this.setPageNavigation(currentPageNode);
      this.setState({ rootNode: navigationTree, isLoading: false });
    }
    catch (error) {
      throw error;
    }
  }

  /**
   * 
   * @param {INavLink} [parentNode]
   * @returns {INavLink}
   */
  private setPageNavigation(parentNode?: INavLink): INavLink {
      let subPages = this.state.pages.filter((item: INavLink) => {
        return (item[`${this.props.lookupField}Id`] === parentNode.key);
      });
      parentNode.links = new Array<INavLink>();
      subPages.map((item) => {
        var subNode: INavLink = { name: item.Title, key: item.Id, url: item.FileRef };
        parentNode.isExpanded = (parentNode.isExpanded || item.FileRef === this.props.serverRequestPath);
        parentNode.links.push(subNode);
        return this.setPageNavigation(subNode);
      });
      return parentNode;
  }

  /**
   * 
   * @returns {Promise<void>}
   */
  private async fetchListItems(): Promise<void> {
    try {
      let pages = await sp.web.getList(this.props.listServerRelativeUrl).items.select(this.props.lookupField + "Id", "Id", "Title", "FileRef").orderBy("Title").get();
      this.setState({ pages: pages });
    } catch (error) {
      throw error;
    }
  }
}
