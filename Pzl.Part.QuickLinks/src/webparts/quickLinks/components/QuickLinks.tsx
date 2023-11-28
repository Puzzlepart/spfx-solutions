import React, { FC } from 'react'
import * as strings from 'QuickLinksWebPartStrings'
import styles from './QuickLinks.module.scss'
import { IQuickLinksProps, ILink, ICategory } from './types'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import { Text } from 'office-ui-fabric-react/lib/Text'
import { useQuickLinks } from './useQuickLinks'
import { Button, FluentProvider, InfoLabel, Link } from '@fluentui/react-components'
import { customLightTheme } from '../../../util/theme'

export const QuickLinks: FC<IQuickLinksProps> = (props) => {
  const { state, callWebHook, backgroundColor } = useQuickLinks(props)

  const generateLinks = (categories: Array<ICategory>) => {
    return categories.map((cat: ICategory, catIndex: number) => {
      const linkItems = cat.links.map((link: ILink, linkIndex) => {
        return (
          <div
            key={`link_${linkIndex}`}
            className={styles.linkGridColumn}
            style={{ lineHeight: `${props.lineHeight}px` }}
          >
            <Text
              className={styles.linkContainer}
              onClick={() => {
                callWebHook(link.url, link.category)
                window.open(link.url, link.openInSameTab ? '_self' : '_blank')
              }}
            >
              <Icon
                className={styles.icon}
                style={{ opacity: props.iconOpacity / 100 }}
                iconName={link.icon ? link.icon : props.defaultIcon}
              />
              <span style={{ width: props.maxLinkLength }}>{link.displayText}</span>
            </Text>
          </div>
        )
      })

      if (props.groupByCategory) {
        return (
          <div className={styles.categorySection}>
            <div className={styles.linkCategoryHeading}>{cat.displayText}</div>
            {linkItems}
          </div>
        )
      }
      return <div key={`category_${catIndex}`}>{linkItems}</div>
    })
  }

  return (
    <FluentProvider
      theme={customLightTheme}
      className={styles.quickLinks}
      style={{ backgroundColor }}
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
          {strings.AllLinksLabel}
        </Link>
      </div>
      <div className={styles.linkGrid}>{generateLinks(state.linkStructure)}</div>
    </FluentProvider>
  )
}

QuickLinks.defaultProps = {
  defaultIcon: 'Link',
  title: strings.Title,
  description: strings.Description,
  iconOpacity: 100,
  lineHeight: 40,
  maxLinkLength: 130
}
