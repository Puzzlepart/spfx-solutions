import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-webpart-base/node_modules/@microsoft/sp-component-base";

export interface IRssFeedProps {
  title: string;
  seeAllUrl: string;
  rssFeedUrl: string;
  apiKey: string;
  itemsCount: number;
  officeUIFabricIcon: string;
  displayMode: DisplayMode;
  context: WebPartContext;
  updateProperty: (value: string) => void;
  cacheDuration: number;
  instanceId: string;
}
