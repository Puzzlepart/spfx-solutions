import React, { FC, ReactNode } from 'react'
import { FluentProvider, Theme, webDarkTheme, webLightTheme } from '@fluentui/react-components'
import type { IReadonlyTheme } from '@microsoft/sp-component-base'

/**
 * Fluent theme for what the web part draws directly on the page section: the
 * header, the button for new announcements and the "no announcements" text.
 * On a section with a dark background (`isInverted`) the text, icons and
 * button outline take the section's own text colour, so they are readable on
 * it. On light sections the light theme is used as before.
 *
 * @param themeVariant Theme of the page section the web part is in
 */
export function getSectionTheme(themeVariant: IReadonlyTheme | undefined): Theme {
  if (!themeVariant?.isInverted) return webLightTheme
  const text = themeVariant.semanticColors?.bodyText || '#ffffff'
  // A light veil for hover and pressed, whatever the section's colour is
  const hover = 'rgba(255, 255, 255, 0.16)'
  const pressed = 'rgba(255, 255, 255, 0.24)'
  return {
    ...webDarkTheme,
    colorNeutralForeground1: text,
    colorNeutralForeground1Hover: text,
    colorNeutralForeground1Pressed: text,
    colorNeutralForeground2: text,
    colorNeutralForeground2Hover: text,
    colorNeutralForeground2Pressed: text,
    colorNeutralForeground2BrandHover: text,
    colorNeutralForeground2BrandPressed: text,
    colorNeutralForeground3: text,
    colorNeutralForeground4: text,
    colorNeutralStroke1: text,
    colorNeutralStroke1Hover: text,
    colorNeutralStroke1Pressed: text,
    colorStrokeFocus2: text,
    colorTransparentBackgroundHover: hover,
    colorTransparentBackgroundPressed: pressed,
    colorSubtleBackgroundHover: hover,
    colorSubtleBackgroundPressed: pressed
  }
}

export interface ISectionThemeProps {
  themeVariant: IReadonlyTheme | undefined
  className?: string
  children?: ReactNode
}

/**
 * Renders its children with the theme of the page section (see `getSectionTheme`).
 * Dialogs and the announcements themselves are not inside it: they have their
 * own light surfaces.
 */
export const SectionTheme: FC<ISectionThemeProps> = ({ themeVariant, className, children }) => (
  <FluentProvider
    theme={getSectionTheme(themeVariant)}
    className={className}
    style={{ backgroundColor: 'transparent' }}>
    {children}
  </FluentProvider>
)
