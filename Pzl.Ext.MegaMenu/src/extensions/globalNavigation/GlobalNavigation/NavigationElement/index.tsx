import * as React from "react";
import INavigationElementProps from "./INavigationElementProps";
import INavigationElementState from "./INavigationElementState";
import NavigationLink from "./NavigationLink";
import styles from './NavigationElement.module.scss';
import * as Breakpoints from "../Breakpoints";
import { Icon } from "office-ui-fabric-react/lib/Icon";

export default class NavigationElement extends React.PureComponent<INavigationElementProps, INavigationElementState> {
    constructor(props: INavigationElementProps) {
        super(props);
        this._onWindowResize(true);
        this._onHeaderClick = this._onHeaderClick.bind(this);
        this._onWindowResize = this._onWindowResize.bind(this);
    }
    public componentDidMount() {
        window.addEventListener("resize", _ => this._onWindowResize());
    }

    public componentWillUnmount() {
        window.removeEventListener("resize", _ => this._onWindowResize());
    }

    public shouldComponentUpdate(nextProps: INavigationElementProps, nextState: INavigationElementState) {
        const shouldUpdate = (this.state.isExpanded != nextState.isExpanded || this.state.deviceWidthStr !== nextState.deviceWidthStr);
        return shouldUpdate;
    }

    public render() {
        let flexDirection = this.state.isExpanded ? styles.isLarge : styles.isSmall;
        let headerStyle = this.getHeaderStyle();
        let links = this.state.isExpanded ?
            <div className={this.isMobile() ? styles.linksMobile : styles.links}>
                <div className={`${styles.linkrow} ${flexDirection}`}>
                    {this.props.links.map(lnk => {
                        const props = { ...lnk, linkTextColor: this.props.linkTextColor };
                        return <NavigationLink {...props} />;
                    })}
                </div>
            </div> : null;

        let headerMarkup = <span>{this.props.header}</span>;
        if (this.props.headerLink && !this.isMobile()) {
            const props = { text: this.props.header, url: this.props.headerLink, linkTextColor: this.props.navHeaderTextColor, isHeader:true };
            headerMarkup = <NavigationLink {...props} />;
        } else if (this.props.headerLink && this.isMobile()) {
            headerMarkup = <span>{this.props.header}<a className={styles.mobileHeadingLink} href={this.props.headerLink}><Icon iconName='Link' title={`Navigate to ${this.props.headerLink}`}></Icon></a></span>;
        }
        return (
            <div className={`${styles.navigationElement} ${this.isMobile() ? 'isMobile' : 'isDesktop'}`}>
                <div onClick={this._onHeaderClick} className={styles.header} style={headerStyle}>
                    {this.isMobile() && this.state.isExpanded && <Icon iconName='ChevronDown' title='Collapse'></Icon>}
                    {this.isMobile() && !this.state.isExpanded && <Icon iconName='ChevronUp' title='Expand'></Icon>}               
                    {headerMarkup}
                </div>
                {links}
            </div >
        );
    }
    /**
 * On window resize
 * 
 * @param {boolean} initial True if initial call from constructor()
 */
    private _onWindowResize(initial = false) {
        const deviceWidthStr = Breakpoints.GetCurrentBreakpoint();
        let isExpanded;
        switch (deviceWidthStr) {
            case "sm": case "md": case "lg": isExpanded = false;
                break;
            default: isExpanded = true;
        }
        const state = { deviceWidthStr, isExpanded };
        if (initial) {
            this.state = state;
        } else {
            this.setState(state);
        }
    }

    private isMobile() {
        return ["sm", "md", "lg"].indexOf(this.state.deviceWidthStr) !== -1;
    }
    /**
     * On header click
     */
    private _onHeaderClick(e) {
        e.preventDefault();
        if (this.isMobile()) {
            this.setState(prevState => ({ isExpanded: !prevState.isExpanded }));
        }
    }

    /**
     * Get header style
     */
    private getHeaderStyle(): React.CSSProperties {
        let style: React.CSSProperties = {};
        switch (this.state.deviceWidthStr) {
            case "sm": case "md": case "lg": style.cursor = "pointer";
                break;
        }
        style.color = this.props.navHeaderTextColor;
        return style;
    }
}

export {
    INavigationElementProps,
    INavigationElementState,
};
