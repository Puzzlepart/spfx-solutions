import * as React from "react";
import INavigationLinkProps from "./INavigationLinkProps";
import styles from './NavigationLink.module.scss';
import * as strings from 'GlobalNavigationApplicationCustomizerStrings';
import { Icon } from "office-ui-fabric-react/lib/Icon";

/**
 * See https://julieturner.net/2018/08/spfx-anchor-tags-hitting-the-target/ for explanation of handling external links
 */
const NavigationLink = ({ text, url, linkTextColor, isHeader }: INavigationLinkProps) => {
    function isLinkExternal(): boolean {
        return url && url.match(document.location.host) == null && url.indexOf("/") != 0;
    }

    function getTarget(): string {
        return isLinkExternal() ? "_blank" : "_self";
    }

    let className = isHeader ? styles.headerLink : styles.navigationLink;
    return (
        <div className={className} >
            <a href={url} target={getTarget()} data-interception={isLinkExternal() ? "off" : "on" } className={styles.anchor} style={{ color: linkTextColor }}>
                <span><span className={styles.anchorText}>{text}</span>{isLinkExternal() && <Icon iconName='NavigateExternalInline' title={strings.Anchor_External_Title} className={styles.pzlExternalIcon} />}</span>
            </a>
        </div>
    );
};

export default NavigationLink;
export { INavigationLinkProps };
