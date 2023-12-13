import { BrandVariants, createDarkTheme, createLightTheme } from '@fluentui/react-components'
import pSBC from 'shade-blend-color'

const themeColors: any = (window as any).__themeState__.theme
const primaryColor = themeColors.themePrimary

const brandVariants: BrandVariants = {
  10: pSBC(-0.5, primaryColor),
  20: pSBC(-0.45, primaryColor),
  30: pSBC(-0.4, primaryColor),
  40: pSBC(-0.35, primaryColor),
  50: pSBC(-0.3, primaryColor),
  60: pSBC(-0.25, primaryColor),
  70: pSBC(-0.2, primaryColor),
  80: pSBC(-0.1, primaryColor),
  90: primaryColor,
  100: pSBC(0.05, primaryColor),
  110: pSBC(0.1, primaryColor),
  120: pSBC(0.15, primaryColor),
  130: pSBC(0.2, primaryColor),
  140: pSBC(0.25, primaryColor),
  150: pSBC(0.3, primaryColor),
  160: pSBC(0.35, primaryColor)
}

export const customLightTheme = createLightTheme(brandVariants)
export const customDarkTheme = createDarkTheme(brandVariants)
