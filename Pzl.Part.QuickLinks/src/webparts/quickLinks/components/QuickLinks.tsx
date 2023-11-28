import React, { FC } from 'react'
import * as strings from 'QuickLinksWebPartStrings'
import styles from './QuickLinks.module.scss'
import { IQuickLinksProps, ILink, ICategory } from './types'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import { Text } from 'office-ui-fabric-react/lib/Text'
import { useQuickLinks } from './useQuickLinks'

export const QuickLinks: FC<IQuickLinksProps> = (props) => {
  const { state, callWebHook, backgroundColor } = useQuickLinks(props)

  const generateLinks = (categories: Array<ICategory>) => {
    return categories.map((cat: ICategory, catIndex: number) => {
      const linkItems = cat.links.map((link: ILink, linkIndex) => {
        const linkIcon = (
          <Icon
            className={styles.icon}
            style={{ opacity: props.iconOpacity / 100 }}
            iconName={link.icon ? link.icon : props.defaultIcon}
          />
        )
        const linkStyle = { width: props.maxLinkLength }
        const linkTarget = link.openInSameTab ? '_self' : '_blank'
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
                window.open(link.url, linkTarget)
              }}
            >
              {linkIcon}
              <span style={linkStyle}>{link.displayText}</span>
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

  const links = generateLinks(state.linkStructure)

  return (
    <div className={styles.quickLinks} style={{ backgroundColor }}>
      <div className={styles.webpartHeader}>
        <span>{props.title}</span>
        <span className={styles.showAll}>
          <Text onClick={() => window.open(props.allLinksUrl, '_blank')}>
            {strings.component_AllLinksLabel}
          </Text>
        </span>
      </div>
      <div className={styles.linkGrid}>{links}</div>
    </div>
  )
}
