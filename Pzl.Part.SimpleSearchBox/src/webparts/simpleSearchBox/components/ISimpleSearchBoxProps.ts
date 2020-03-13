
import { DisplayMode } from '@microsoft/sp-core-library';
export interface ISimpleSearchBoxProps {
  searchurl: string;
  title: string;
  placeholder: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
