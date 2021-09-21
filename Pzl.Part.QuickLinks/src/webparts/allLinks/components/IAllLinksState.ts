
export interface IAllLinksState {
    editorLinks?: Array<Link>;
    favouriteLinks?: Array<Link>;
    mandatoryLinks?: Array<Link>;
    categoryLinks?: Array<ICategory>;
    isLoading?: boolean;
    showErrorMessage?: boolean;
    showSuccessMessage?: boolean;
    showLoadingSpinner?: boolean;
    currentUser?: User;
    showModal?: boolean;
    modalData?: Link;
    isFirstUpdate?: boolean;
    saveButtonDisabled?: boolean;
}

export interface Link {
    id?: number;
    displayText: string;
    url: string;
    icon?: string;
    priority?: string;
    mandatory?: number;
    linkType: LinkType;
}

export enum LinkType {
    editorLink = "EditorLink",
    favouriteLinks = "FavouriteLink",
    mandatoryLinks = "MandatoryLinks"
}

export interface User {
    id: number;
    linkFieldId?: string;
}

export interface ILink {
    id?: number;
    displayText: string;
    url: string;
    icon: string;
    category: string;
    priority: string;
    mandatory?: number;
    linkType: LinkType;
    openInSameTab?: boolean;
}
export interface ICategory {
    links: Array<ILink>;
    displayText: string;
}
