import React, { FC } from 'react'
import * as strings from 'QuickLinksWebPartStrings'
import styles from './QuickLinks.module.scss'
import { IQuickLinksProps, ILink, ICategory } from './types'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import { useQuickLinks } from './useQuickLinks'
import { Button, FluentProvider, InfoLabel, Link } from '@fluentui/react-components'
import { customLightTheme } from '../../../util/theme'

export const QuickLinks: FC<IQuickLinksProps> = (props) => {
  const { state, callWebHook, backgroundColor } = useQuickLinks(props)

  const generateLinks = (categories: Array<ICategory>) => {

    return categories.map((category: ICategory, idx) => {
      const linkItems = category.links.map((link: ILink, idx) => {
        return (
          <Button
            key={`link_${idx}`}
            title={link.displayText}
            style={{ lineHeight: `${props.lineHeight}px`, width: (props.responsiveButtons || props.iconsOnly) ? 'auto' : '100%' }}
            className={styles.link}
            appearance='subtle'
            size='medium'
            icon={
              <Icon
                className={styles.icon}
                style={{ opacity: props.iconOpacity / 100 }}
                iconName={link.icon ? link.icon : props.defaultIcon}
              />
            }
            onClick={() => {
              callWebHook(link.url, link.category)
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
            <div className={styles.heading}>{category.displayText !== undefined ? category.displayText : 'Mine lenker'}</div>
            {linkItems}
          </div>
        )
      }
      return linkItems
    })
  }

  return (
    <FluentProvider
      theme={customLightTheme}
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
          {strings.AllLinksLabel}
        </Link>
      </div>
      <div className={styles.links}>{generateLinks(state.linkStructure)}</div>
    </FluentProvider>
  )
}

QuickLinks.defaultProps = {
  defaultIcon: 'Link',
  title: strings.Title,
  description: strings.Description,
  iconOpacity: 100,
  lineHeight: 20
}
