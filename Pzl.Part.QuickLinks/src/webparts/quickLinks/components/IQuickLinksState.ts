export interface IQuickLinksState {
    linkStructure: Array<ICategory>;
}

export interface ILink {
    id?: number;
    displayText: string;
    url: string;
    icon: string;
    category: string;
    priority: string;
    openInSameTab: boolean;
}
export interface ICategory {
    links: Array<ILink>;
    displayText: string;
}