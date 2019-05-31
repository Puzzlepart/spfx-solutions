import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IBirthdayWebPartProps {
  title: string;
  itemsCount: number;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateProperty: (value: string) => void;
}
