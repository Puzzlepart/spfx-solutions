import React, { FC } from 'react'
import * as strings from 'QuickLinksWebPartStrings'
import styles from './QuickLinks.module.scss'
import { IQuickLinksProps, ILink, ICategory } from './types'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import { stringIsNullOrEmpty } from '@pnp/common'
import { useQuickLinks } from './useQuickLinks'
import {
  Button,
  FluentProvider,
  IdPrefixProvider,
  InfoLabel,
  Link,
  useId
} from '@fluentui/react-components'

export const QuickLinks: FC<IQuickLinksProps> = (props) => {
  const { state, callWebHook, backgroundColor, theme } = useQuickLinks(props)
  const fluentProviderId = useId('fp-your-links')

  const generateLinks = (categories: Array<ICategory>) => {
    const uncategorizedLinks = categories
      .filter((category) => category.displayText !== undefined)
      .map((category) => category.links)
    const favoriteLinks = categories
      .filter((category) => category.displayText === undefined)
      .map((category) => category.links)
    const sortedLinks = uncategorizedLinks
      .reduce((acc, val) => acc.concat(val), [])
      .sort((a, b) => Number(a.priority) - Number(b.priority))

    sortedLinks.push(...favoriteLinks.reduce((acc, val) => acc.concat(val), []))

    if (props.groupByCategory) {
      return categories.map((category: ICategory, idx) => {
        const linkItems = category.links.map((link: ILink, idx) => {
          return (
            <Button
              key={`link_${idx}`}
              title={link.displayText}
              style={{
                lineHeight: `${props.lineHeight}px`,
                width: props.responsiveButtons || props.iconsOnly ? 'auto' : '100%'
              }}
              className={styles.link}
              appearance={props.buttonAppearance}
              size={props.iconSize >= 26 ? 'large' : props.iconSize <= 16 ? 'small' : 'medium'}
              icon={
                <Icon
                  className={styles.icon}
                  style={{ fontSize: props.iconSize }}
                  iconName={link.icon ? link.icon : props.defaultIcon}
                />
              }
              onClick={() => {
                callWebHook(link.id, link.url, link.category)
                window.open(link.url, link.openInSameTab ? '_self' : '_blank')
              }}
            >
              {!props.iconsOnly && <span className={styles.label}>{link.displayText}</span>}
            </Button>
          )
        })

        if (props.groupByCategory) {
          return (
            <div className={styles.categorySection} key={`category_${idx}`}>
              <div className={styles.heading}>
                {category.displayText !== undefined ? category.displayText : 'Mine lenker'}
              </div>
              <div className={styles.links} style={{ gap: props.groupByCategory && props.gapSize }}>
                {linkItems}
              </div>
            </div>
          )
        }
        return linkItems
      })
    } else {
      return sortedLinks.map((link: ILink, idx) => {
        return (
          <Button
            key={`link_${idx}`}
            title={link.displayText}
            style={{
              lineHeight: `${props.lineHeight}px`,
              width: props.responsiveButtons || props.iconsOnly ? 'auto' : '100%'
            }}
            className={styles.link}
            appearance={props.buttonAppearance}
            size={props.iconSize >= 26 ? 'large' : props.iconSize <= 16 ? 'small' : 'medium'}
            icon={
              <Icon
                className={styles.icon}
                style={{ fontSize: props.iconSize }}
                iconName={link.icon ? link.icon : props.defaultIcon}
              />
            }
            onClick={() => {
              callWebHook(link.id, link.url, link.category)
              window.open(link.url, link.openInSameTab ? '_self' : '_blank')
            }}
          >
            {!props.iconsOnly && <span className={styles.label}>{link.displayText}</span>}
          </Button>
        )
      })
    }
  }

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider
        theme={theme}
        className={styles.quickLinks}
        style={{ backgroundColor, boxShadow: props.renderShadow && 'var(--shadow2)' }}
      >
        <div className={styles.header} style={{ display: props.hideHeader && 'none' }}>
          <InfoLabel
            className={styles.title}
            info={props.description}
            style={{ display: props.hideTitle && 'none' }}
          >
            <span>{props.title}</span>
          </InfoLabel>
          <Link
            onClick={() => window.open(props.allLinksUrl, '_blank')}
            style={{ display: props.hideShowAll && 'none' }}
          >
            {stringIsNullOrEmpty(props.allLinksText) ? strings.AllLinksLabel : props.allLinksText}
          </Link>
        </div>
        <div className={styles.links} style={{ gap: !props.groupByCategory && props.gapSize }}>
          {generateLinks(state.linkStructure)}
        </div>
      </FluentProvider>
    </IdPrefixProvider>
  )
}

QuickLinks.defaultProps = {
  defaultIcon: 'Link',
  title: strings.Title,
  description: strings.Description,
  lineHeight: 20,
  gapSize: 8,
  iconSize: 20,
  buttonAppearance: 'subtle'
}
