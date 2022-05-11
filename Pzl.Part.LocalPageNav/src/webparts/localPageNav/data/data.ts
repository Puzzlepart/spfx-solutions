import { INavLink, INavLinkGroup } from 'office-ui-fabric-react';
import { includes } from 'lodash';

/**
 * Gets navigation links from the given heading selector.
 * 
 * @param selector Heading selector from webpart settings.
 * @returns Navigation links grouped by their position on the page.
 */
export const getNavLinks = (selector: string[]): INavLinkGroup => {
    let map = {}, links = [];

    const nodes: any = Array.from<any>(document.querySelectorAll(selector.join(','))).map(node => {
        return {
            name: node.innerText,
            key: `#${node.id}`,
            url: `#${node.id}`,
            links: [],
            isExpanded: true,
            linkStyle: parseInt(node.localName.substring(1)) // The level
        };
    });

    for (let i = 0; i < nodes.length; i++) {

        map[nodes[i].key] = i; // Initialize the map

        if (i === 0 || nodes[i].linkStyle === 1) {
            nodes[i].parent = ''; // First node and level 1 nodes doesn't have a parent
        } else if (nodes[i].linkStyle > nodes[i - 1].linkStyle) {
            nodes[i].parent = nodes[i - 1].key; // Level 2 or 3, below previous, which is the parent
        } else if (nodes[i].linkStyle === nodes[i - 1].linkStyle) {
            nodes[i].parent = nodes[i - 1].parent; // Level 2 or 3, like the previous, share the parent
        }
        else {
            // Level 2, previous was level 3. Need to traverse back to find parent
            let p = i - 1;
            while (nodes[p].linkStyle > nodes[i].linkStyle && 0 < p) {
                p--;
            }
            nodes[i].parent = nodes[p].parent; // Use the parent of the first previous heading not on lower level
        }
    }

    for (let i = 0; i < nodes.length; i++) {
        let node = nodes[i];
        if (node.parent !== '') {
            // if you have dangling branches check that map[node.parent] exists
            nodes[map[node.parent]].links.push(node);
        } else {
            links.push(node);
        }
    }
    return { links: links };
};
