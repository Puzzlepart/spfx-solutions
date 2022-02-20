import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'LocalPageNavWebPartStrings';
import LocalPageNav from './components/LocalPageNav';
import { ILocalPageNavProps } from './components/ILocalPageNavProps';
import { INavLinkGroup } from 'office-ui-fabric-react/lib/Nav';
import { getNavLinks } from './data/data';

export interface ILocalPageNavWebPartProps {
  title: string;
  selector: string[];
}

export default class LocalPageNavWebPart extends BaseClientSideWebPart<ILocalPageNavWebPartProps> {
  private _navLinks: INavLinkGroup;

  public render(): void {
    const element: React.ReactElement<ILocalPageNavProps> = React.createElement(
      LocalPageNav,
      {
        title: this.properties.title,
        navLinks: this._navLinks
      }
    );

    ReactDom.render(element, this.domElement);
  }


  protected onInit(): Promise<void> {
    super.onInit();
    this._navLinks = getNavLinks(this.properties.selector);
    return Promise.resolve();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.SettingsGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyFieldMultiSelect('selector', {
                  key: 'selector',
                  label: "Included headings",
                  options: [
                    {
                      key: "h2",
                      text: "Heading 1"
                    },
                    {
                      key: "h3",
                      text: "Heading 2"
                    },
                    {
                      key: "h4",
                      text: "Heading 3"
                    }
                  ],
                  selectedKeys: this.properties.selector
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
