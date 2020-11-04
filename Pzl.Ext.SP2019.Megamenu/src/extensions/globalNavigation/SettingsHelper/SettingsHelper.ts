import { Web } from '@pnp/sp';
import * as urljoin from 'url-join';



interface NavigationSettings {
  navToggleText?: string;
  navToggleTextColor?: string;
  navToggleBackgroundColor?: string;
  navHeaderTextColor?: string;
  navContentBackgroundColor?: string;
  navHideTheMenu?: string;
  linkTextColor?: string;
  navColumns?: string;
  helpButtonEnabled?: string;
  helpButtonText?: string;
  helpButtonUrl?: string;
  helpButtonColor?: string;
  helpButtonTextColor?: string;
  helpButtonIcon?: string;
  homeButtonEnabled?: string;
  homeButtonMobileOnly?: string;
  homeButtonText?: string;
  homeButtonUrl?: string;
  homeButtonColor?: string;
  homeButtonTextColor?: string;
  homeButtonIcon?: string;
  homeButtonFloatLeft?: string;
  searchBarEnabled?: string;
  searchBarPlaceholder?: string;
  searchBarSearchUrl?: string;
  searchBarUrlParam?: string;
  focusButtonEnabled?: string;
  focusButtonTextColor?: string;
  focusButtonText?: string;
  focusButtonColor?: string;
  focusButtonIcon?: string;
  focusButtonActiveColor?: string;
  announcementLevels?: string;
}


async function fetchGlobalNavigationSettings(serverRelativeWebUrl: string, settingsListUrl: string) {
  const webUrl = urljoin(`${document.location.protocol}//${document.location.hostname}`, serverRelativeWebUrl);
  const listUrl = urljoin(serverRelativeWebUrl, settingsListUrl);
  let navSettings: NavigationSettings = {};

  try {
    const web = new Web(webUrl);
    const settings = await web.getList(listUrl).items.select('Title', 'PzlSettingValue').usingCaching().get();
    settings.forEach(setting => {
      navSettings[setting['Title']] = setting['PzlSettingValue'];
    });
  } catch(e) {
    console.log(e);
  }

  return navSettings;
}


export {NavigationSettings, fetchGlobalNavigationSettings};
