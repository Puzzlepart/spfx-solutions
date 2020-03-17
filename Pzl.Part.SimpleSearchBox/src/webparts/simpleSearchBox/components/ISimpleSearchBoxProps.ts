import { DisplayMode } from '@microsoft/sp-core-library';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ISimpleSearchBoxProps {
  searchurl: string;
  title: string;
  openInNewTab: boolean;
  placeholder: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  themeVariant: IReadonlyTheme | undefined;
}
