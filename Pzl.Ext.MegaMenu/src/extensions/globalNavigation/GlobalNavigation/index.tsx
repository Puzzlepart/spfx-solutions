import * as React from 'react';
import * as ClickOutHandler from 'react-onclickout';
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { SearchBox } from "office-ui-fabric-react/lib/SearchBox";
import IGlobalNavigationProps from './IGlobalNavigationProps';
import IGlobalNavigationState from './IGlobalNavigationState';
import NavigationElement, { INavigationElementProps } from './NavigationElement';
import styles from './GlobalNavigation.module.scss';
import * as Breakpoints from "./Breakpoints";

export default class GlobalNavigation extends React.PureComponent<IGlobalNavigationProps, IGlobalNavigationState> {
  constructor(props: IGlobalNavigationProps) {
    super(props);
    this.state = { isExpanded: false, items: [], isXtraLarge: this.isLargerThanLarge(), isLoading: true };
    this._onToggle = this._onToggle.bind(this);
    this._onWindowResize = this._onWindowResize.bind(this);
  }

  public componentWillUnmount(): void {
    window.removeEventListener("resize", _ => this._onWindowResize());
    this._onWindowResize();
  }

  public async componentDidMount(): Promise<void> {
    try {
      const data = await this.getData();
      this.setState(data);
      window.addEventListener("resize", _ => this._onWindowResize());
    } catch (error) {
      this.setState({ error: error });
    }
  }

  public render(): JSX.Element {
    if (this.state.error) {
      return (
        <div className={styles.globalNavigation}>
          <MessageBar messageBarType={MessageBarType.severeWarning}>{this.props.errorText}</MessageBar>
        </div>
      );
    } else if (!this.state.isLoading) {
      const HomeButton = this.renderHomeButton;
      const FocusButton = this.renderFocusButton;
      const HelpButton = this.renderHelpButton;
      const NavSearchBox = this.renderSearchBox;
      const settings = this.props.settings ? this.props.settings : {};
      const textColor = settings.navToggleTextColor ? settings.navToggleTextColor : '';
      const backgroundColor = settings.navToggleBackgroundColor ? settings.navToggleBackgroundColor : '';
      const linkTextColor = settings.linkTextColor ? settings.linkTextColor : '';
      const navContentBackgroundColor = settings.navContentBackgroundColor ? settings.navContentBackgroundColor : '';
      return (
        <ClickOutHandler onClickOut={e => this._onToggle(e, false)}>
          <div className={`${styles.globalNavigation} ${this.state.isFocusToggled ? ' isFocusToggled' : ''}`}>
            <div className={`${styles.toggleNavRow}`} style={{ color: textColor, backgroundColor: backgroundColor }} >
              {settings.homeButtonFloatLeft ? <HomeButton /> : ''}
              <div className={`${styles.toggleNavColumn}`} onClick={this._onToggle}>
                <div className={styles.toggleNav}>
                  <span className={styles.toggleNavText}>{settings.navToggleText ? settings.navToggleText : 'Menu'}</span>
                  <Icon className={styles.toggleNavIcon} iconName={this.state.isExpanded ? "ChevronUp" : "ChevronDown"} />
                </div>
              </div>
              <NavSearchBox />
              <FocusButton />
              {settings.homeButtonFloatLeft ? '' : <HomeButton />}
              <HelpButton />
            </div>
            <div className={styles.navRowsContainer} style={{ color: linkTextColor, backgroundColor: navContentBackgroundColor }}>
              {this.state.isExpanded ? this.renderRows() : null}
            </div>
          </div>
        </ClickOutHandler>
      );
    } else {
      return null;
    }
  }
  /**
   *
   *
   * @memberof GlobalNavigation
   */
  public closeDialog(): void {
    setTimeout(() => {
      this.setState({ isExpanded: false });
    }, 200);
  }

  private async getData(): Promise<Partial<IGlobalNavigationState>> {
    let navigationElements;
    let errorCount = 2;
    while (errorCount > 0) {
      try {
        navigationElements = await this.props.dataFetch.fetch();
        break;
      } catch (error) {
        errorCount--;
        if (errorCount == 0) throw error;
      }
    }

    navigationElements = navigationElements.filter(element => element.order > -1);

    return ({
      items: navigationElements,
      isLoading: false,
    });
  }

  /**
   * Renders a row for every X columns
   */
  private renderRows() {
    const settings = this.props.settings ? this.props.settings : {};
    const linkTextColor = settings.linkTextColor ? settings.linkTextColor : '';
    const navHeaderTextColor = settings.navHeaderTextColor ? settings.navHeaderTextColor : '';
    const navContentBackgroundColor = settings.navContentBackgroundColor ? settings.navContentBackgroundColor : '';
    const navColumns = settings.navColumns ? +settings.navColumns : 5;
    const elementsWithItems = this.state.items.filter(item => item.links.length > 0);
    const rows = [];
    const flexDirection = this.state.isXtraLarge ? styles.isLarge : styles.isSmall;
    let numberOfColumns = navColumns;
    if (!numberOfColumns || numberOfColumns < 0) {
      numberOfColumns = 5;
    }

    for (let i = 0; i < elementsWithItems.length; i += numberOfColumns) {
      const row = (
        <div className={[styles.navElementsRow, flexDirection].join(" ")} style={{ color: linkTextColor, backgroundColor: navContentBackgroundColor }}>
          {[].concat(elementsWithItems).slice(i, i + numberOfColumns).map(item => {
            const props: INavigationElementProps = {
              ...item,
              navHeaderTextColor: navHeaderTextColor,
              linkTextColor: linkTextColor
            };
            return <NavigationElement {...props} />;
          })}
        </div>
      );
      rows.push(row);
    }
    return rows;
  }

  private _onToggle(_, isExpanded?: boolean): void {
    if (typeof isExpanded !== 'undefined') {
      // hide id clicked outside the menu
      this.setState({ isExpanded: false });
    } else {
      this.setState(prevState => ({ isExpanded: !prevState.isExpanded }));
    }
  }

  /**
   * On window resize
   */
  private _onWindowResize(): void {
    const isXtraLarge: boolean = this.isLargerThanLarge();
    this.setState({ isXtraLarge: isXtraLarge });
  }

  private isLargerThanLarge(): boolean {
    const deviceWidthStr: string = Breakpoints.GetCurrentBreakpoint();
    let isXtraLarge: boolean;
    switch (deviceWidthStr) {
      case "sm": case "md": case "lg": isXtraLarge = false;
        break;
      default: isXtraLarge = true;
    }
    return isXtraLarge;
  }

  private renderHomeButton = (): JSX.Element => {
    const settings = this.props.settings ? this.props.settings : {};
    const textColor = settings.homeButtonTextColor ? settings.homeButtonTextColor : '#ffffff';
    const backgroundColor = settings.homeButtonColor ? settings.homeButtonColor : '';
    const homeButtonUrl = settings.homeButtonUrl ? settings.homeButtonUrl : '';
    const isHidden = (settings.homeButtonEnabled === 'false' || (settings.homeButtonMobileOnly === 'true' && !this.state.isXtraLarge));

    return (
      <div className={`${styles.homeButtonContainer}`} style={{ color: textColor, backgroundColor: backgroundColor }} hidden={isHidden}>
        <a href={homeButtonUrl} title={settings.homeButtonText ? settings.homeButtonText : ''}>
          <Icon style={{ color: textColor }} iconName={settings.homeButtonIcon ? settings.homeButtonIcon : 'Home'} />
        </a>
      </div>
    );
  }

  private toggleFocus(): void {
    const toggleFocusSelectors: string[] = [
      "div.od-SuiteNav",
      "div.commandBarWrapper",
      "#SuiteNavPlaceHolder",
      "div[role='banner']",
      "div.sp-pageLayout-sideNav div[class^='spNav_']",
      "div.sp-App--hasLeftNav .Files-leftNav"
    ];
    toggleFocusSelectors.forEach((selector: string): void => this.toggleElement(selector));
  }

  private toggleElement(selector): void {
    const element = document.querySelector(selector);
    element ? 
      (
        element.style.display == "none" ? 
          element.style.display = "block" : 
          element.style.display = "none"
      ) : 
      console.log(`Pzl.Megamenu.FocusOnContent: element ${selector} not found.`);
  }

  private renderFocusButton = (): JSX.Element => {
    const settings = this.props.settings ? this.props.settings : {};
    const textColor = settings.focusButtonTextColor ? settings.focusButtonTextColor : '#ffffff';
    const backgroundColor = settings.focusButtonColor ? settings.focusButtonColor : '';
    const backgroundColorWhenActive = settings.focusButtonActiveColor ? settings.focusButtonActiveColor : '#000000';
    const isHidden = settings.focusButtonEnabled !== 'true';

    return (
      <div className={`${styles.focusButtonContainer}`} style={{ color: textColor, backgroundColor: this.state.isFocusToggled ? backgroundColorWhenActive : backgroundColor }} hidden={isHidden}>
        <Icon onClick={() => {
          this.setState({ isFocusToggled: !this.state.isFocusToggled });
          this.toggleFocus();
        }} title={settings.focusButtonText ? settings.focusButtonText : 'Focus on content'} style={{ color: textColor, backgroundColor: this.state.isFocusToggled ? backgroundColorWhenActive : backgroundColor }} iconName={settings.focusButtonIcon ? settings.focusButtonIcon : 'ZoomToFit'} />
      </div>
    );
  }

  private renderHelpButton = (): JSX.Element => {
    const settings = this.props.settings ? this.props.settings : {};
    const textColor = settings.helpButtonTextColor ? settings.helpButtonTextColor : '#ffffff';
    const backgroundColor = settings.helpButtonColor ? settings.helpButtonColor : 'ff0000';
    const buttonText = settings.helpButtonText ? settings.helpButtonText : '';
    const helpButtonUrl = settings.helpButtonUrl ? settings.helpButtonUrl : '';
    return (
      <div className={`${styles.supportLinkContainer}`} style={{ backgroundColor: backgroundColor }} hidden={!(settings.helpButtonEnabled === 'true')}>
        <a href={helpButtonUrl} style={{ color: textColor }} target="_blank" >
          <Icon iconName={settings.helpButtonIcon ? settings.helpButtonIcon : 'Home'} />
          <span>{this.state.isXtraLarge ? <span className={styles.supportLinkText}>{buttonText}</span> : null}</span>
        </a>
      </div>
    );
  }

  private renderSearchBox = (): JSX.Element => {
    const settings = this.props.settings ? this.props.settings : {};
    return (
      <div className={`${styles.searchBoxContainer}`} hidden={settings.searchBarEnabled !== 'true' || !this.state.isXtraLarge}>
        <SearchBox className={styles.searchBox} onSearch={(query) => this.onSearch(query, settings.searchBarUrlParam)} placeholder={settings.searchBarPlaceholder ? settings.searchBarPlaceholder : 'Search'} />
      </div>
    );
  }

  private onSearch(searchValue, queryParam?: string): void {
    const safeQueryParam = (queryParam && queryParam.length > 0) ? queryParam : 'q';
    const searchUrl = this.props.settings && this.props.settings.searchBarSearchUrl ? this.props.settings.searchBarSearchUrl : `${this.props.currentSiteUrl}/_layouts/15/search.aspx`;
    const searchUrlWithParams = searchUrl.indexOf('?') > -1 ? searchUrl : `${searchUrl}?${safeQueryParam}=`;
    location.href = `${searchUrlWithParams}${searchValue}`;
  }
}
export {
  IGlobalNavigationProps,
  IGlobalNavigationState,
};
