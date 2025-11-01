import { Version } from '@microsoft/sp-core-library';
import { 
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'WelcomeMessageWebPartStrings';

import styles from './WelcomeMessageWebPart.module.scss';

export interface IWelcomeMessageWebPartProps {
  welcomeText: string;
  removeWebPartMarginPadding: boolean;
}

export default class WelcomeMessageWebPart extends BaseClientSideWebPart<IWelcomeMessageWebPartProps> {
  public render(): void {
    let userDisplayName = this.context.pageContext.user.displayName;
    if (userDisplayName.indexOf(',') !== -1) {
      userDisplayName = userDisplayName.split(',')[1].trim();
    }
    
    // Use the configurable welcome text (default value is set in manifest)
    const welcomeText = this.properties.welcomeText;
    
    // Replace the {user} token with the actual user name
    const finalWelcomeText = welcomeText.replace('{user}', userDisplayName);
    
    this.domElement.innerHTML = `<div id="pzl-welcomeMessage" class="${ styles.welcomeMessage }"><h2>${finalWelcomeText}</h2></div>`
    
    // Apply or remove margin/padding styles based on property
    this._applyMarginPaddingStyles();
  }

  private _applyMarginPaddingStyles(): void {
    // Remove any existing style element we might have added
    const existingStyle = document.getElementById('pzl-welcomeMessage-margin-padding-style');
    if (existingStyle) {
      existingStyle.remove();
    }

    if (this.properties.removeWebPartMarginPadding) {
      // Find the CanvasControl parent element by traversing up the DOM tree
      let currentElement = this.domElement.parentElement;
      while (currentElement) {
        if (currentElement.getAttribute('data-automation-id') === 'CanvasControl') {
          // Create a unique CSS rule for this specific CanvasControl
          const webPartId = this.context.instanceId;
          const styleElement = document.createElement('style');
          styleElement.id = 'pzl-welcomeMessage-margin-padding-style';
          styleElement.innerHTML = `
            [data-automation-id="CanvasControl"][data-sp-web-part-id="${webPartId}"] {
              margin-top: 0 !important;
              margin-bottom: 0 !important;
            }
          `;
          
          // Add the data attribute to make our CSS selector work
          currentElement.setAttribute('data-sp-web-part-id', webPartId);
          
          // Add the style to the document head
          document.head.appendChild(styleElement);
          break;
        }
        currentElement = currentElement.parentElement;
      }
    }
  }

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: string | boolean, newValue: string | boolean): void {
    if (propertyPath === 'welcomeText') {
      this.render();
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('welcomeText', {
                  label: strings.WelcomeTextFieldLabel,
                  placeholder: strings.WelcomeTextPlaceholder,
                  multiline: false
                }),
                PropertyPaneToggle('removeWebPartMarginPadding', {
                  label: strings.RemoveWebPartMarginPaddingFieldLabel,
                  onText: strings.On,
                  offText: strings.Off
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
