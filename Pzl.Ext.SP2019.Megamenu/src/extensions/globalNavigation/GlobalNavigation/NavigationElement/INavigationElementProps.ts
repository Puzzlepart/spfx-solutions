import { INavigationLinkProps } from "./NavigationLink";

export default interface INavigationElementProps extends React.HTMLAttributes<HTMLElement> {
    header: string;
    searchQuery?: string;
    links: INavigationLinkProps[];
    columnWidth?: { [key: string]: number };
    navHeaderTextColor?: string;
    headerLink?: string;
    linkTextColor?: string;
    isExpanded?: boolean;
    order: number;
}