import { INavigationElementProps } from './NavigationElement';

export default interface IGlobalNavigationState {
    isExpanded?: boolean;
    isXtraLarge?: boolean;
    isFocusToggled?: boolean;
    items?: INavigationElementProps[];
    error?: any;
    isLoading?: boolean;
    hasHomeButton?: boolean;
    hasPromotedButton?: boolean;
    toggleBarItem?: any;
    homeButtonLink?: any;
}
