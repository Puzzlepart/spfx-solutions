import * as React from 'react';
import styles from './Breadcrumb.module.scss';
import { Breadcrumb, IBreadcrumbItem, IDividerAsProps } from 'office-ui-fabric-react/lib/Breadcrumb';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { IBreadcrumbProps } from './IBreadcrumbProps';
import { IBreadcrumbState } from './IBreadcrumbState';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { sp } from "sp-pnp-js";
import { IODataListItem } from '@microsoft/sp-odata-types';

export interface IPage extends IODataListItem {
  FileRef: string;
}
/**
 *
 *
 * @export
 * @class Breadcrumb
 * @extends {React.Component<IBreadcrumbProps, IBreadcrumbState>}
 */
export default class SitePagesBreadcrumb extends React.Component<IBreadcrumbProps, IBreadcrumbState> {
  private breadcrumb: Array<IBreadcrumbItem>;
  constructor(props) {
    super(props);
    this.state = ({
      pages: [],
      items: [],
      isLoading: true
    });
    this.breadcrumb = [];
  }
  /**
   *
   *
   * @returns {React.ReactElement<IBreadcrumbProps>}
   * @memberof Breadcrumb
   */
  public render(): React.ReactElement<IBreadcrumbProps> {
    if (this.state.items.length <= 0) {
      return null;
    }
    if (this.state.isLoading) {
      return <Spinner  size={SpinnerSize.large}/>;
    }
    let elements: Array<IBreadcrumbItem> = this.state.items.map((item: IPage, index, { length }) => {
      return {text: item.Title, key: item.FileRef, onClick: this._onBreadcrumbItemClicked.bind(this), isCurrentItem: (length - 1 === index)};
    });
    return (
      <Breadcrumb
          className= {styles.breadcrumb}
          items={elements}
          getStyles={() => {
            return {
              itemLink: {
                maxWidth: "100%"
              },
            };
          }
        }
        />
    );
  }
  /**
   *
   *
   * @returns {Promise<void>}
   * @memberof Breadcrumb
   */
  public async componentDidMount(): Promise<void> {
    await this.fetchListItems();
    await this.buildBreadcrumb();
  }
  /**
   *
   *
   * @param {*} prevProps
   * @returns {Promise<void>}
   * @memberof Breadcrumb
   */
  public componentDidUpdate(prevProps): void {
    if (this.props.lookupField !== prevProps.lookupField) {
      this.buildBreadcrumb();
    }
  }
  /**
   *
   *
   * @private
   * @memberof Breadcrumb
   */
  private buildBreadcrumb() {
    this.breadcrumb = [];
    try {
      this.setBreadcrumb(this.props.currentPage.id, 0);

      this.setState({
        items: (this.breadcrumb) ? this.breadcrumb.reverse(): [],
        isLoading: false
      });
    }
    catch (error) {
      throw error;
    }
  }

  /**
   *
   *
   * @private
   * @param {*} pageId
   * @param {*} depth
   * @returns {Promise<void>}
   * @memberof Breadcrumb
   */
  private setBreadcrumb(pageId, depth): void {
    try {
      let currentPage = this.state.pages.filter((item)=> (item.Id === pageId))[0];
      this.breadcrumb.push(currentPage);
      depth ++;
      if(depth < 1000) { 
        if (currentPage[this.props.lookupField + "Id"] && currentPage[this.props.lookupField + "Id"] !== currentPage.Id ) {
          return this.setBreadcrumb(currentPage[this.props.lookupField + "Id"], depth);
        } else {
          return null;
        }
      }
      else {
        this.breadcrumb = null;
        return null;
      }
     
    } catch (error) {
      throw error;
    }
  }
  
  /**
   *
   *
   * @private
   * @returns {Promise<void>}
   * @memberof Breadcrumb
   */
  private async fetchListItems(): Promise<void> {
    try {
      let pages = await sp.web.getList(this.props.listServerRelativeUrl).items.select(this.props.lookupField + "Id", "Id", "Title", "FileRef").get();
      this.setState({pages: pages});
    } catch (error) {
      throw error;
    }
  }
  private _onBreadcrumbItemClicked = (ev: React.MouseEvent<HTMLElement>, item: IBreadcrumbItem): void => {
    window.location.href = item.key;
  }
}
