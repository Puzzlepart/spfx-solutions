import { INavLink, INavLinkGroup } from 'office-ui-fabric-react';
import { includes } from 'lodash';

/**
 * Gets navigation links from the given heading selector.
 * 
 * @param selector Heading selector from webpart settings.
 * @returns Navigation links grouped by their position on the page.
 */
export const getNavLinks = (selector: string[]): INavLinkGroup => {
    let navLinks: INavLink[] = [];
    const nodes: NodeList = document.querySelectorAll(selector.join(','));
    let currentHeader: INavLink;
    let prevLink: INavLink;
    let currentSubheader: INavLink;
    let prevPosition: number;
    nodes.forEach((node: any, key, parent) => {
        if (node.id && node.id !== '') {
            const currentPosition: number = parseInt(node.localName.substring(1));
            const navLink: INavLink = { name: node.innerText, key: `#${node.id}`, url: `#${node.id}`, links: [], isExpanded: true };
            if (prevPosition) {
                if (currentPosition > prevPosition) {
                    if (!currentHeader) {
                        currentHeader = prevLink;
                        currentHeader.links.push(navLink);
                    } else if (currentSubheader) {
                        currentSubheader.link.push(navLink);
                    } else {
                        if (prevLink !== currentHeader) {
                            prevLink.links.push(navLink);
                            if (!currentSubheader) currentSubheader = prevLink;
                        } else currentHeader.links.push(navLink);
                    }
                }
                if (currentPosition < prevPosition) {
                    if (currentSubheader) currentHeader.links.push(currentSubheader);
                    navLinks.push(currentHeader);
                    currentHeader = navLink;
                    currentSubheader = null;
                }
                if (currentPosition === prevPosition) {
                    if (currentSubheader) {
                        currentSubheader.links.push(navLink);
                    } else if (currentHeader) {
                        currentHeader.links.push(navLink);
                    } else navLinks.push(prevLink, navLink);
                }
            }
            prevPosition = currentPosition;
            prevLink = navLink;
        }
        if (parent.length === key + 1 && currentHeader && !includes(navLinks, currentHeader)) {
            navLinks.push(currentHeader);
        }
    });
    return { links: navLinks };
};