import * as React from 'react'
import styles from './AllLinks.module.scss'
import { IAllLinksProps, Link, LinkType, ILink, ICategory } from './types'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import { TextField } from 'office-ui-fabric-react/lib/TextField'
import * as strings from 'AllLinksWebPartStrings'
import { Text } from 'office-ui-fabric-react/lib/Text'
import { stringIsNullOrEmpty } from '@pnp/common'
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  DialogTrigger,
  Field,
  FluentProvider,
  Input,
  MessageBar,
  Spinner
} from '@fluentui/react-components'
import { customLightTheme } from '../../../util/theme'
import { useAllLinks } from './useAllLinks'
import { AddFilled, AddRegular, bundleIcon } from '@fluentui/react-icons'

const Icons = {
  Add: bundleIcon(AddFilled, AddRegular)
}

export const AllLinks: React.FC<IAllLinksProps> = (props) => {
  const {
    state,
    setState,
    backgroundColor,
    openNewItemModal,
    appendToFavourites,
    removeFromFavourites,
    removeCustomFromFavourites,
    addNewLink,
    onModalValueChanged,
    validateUrl
  } = useAllLinks(props)

  console.log({ state, props })

  const generateEditorLinkComponents = (links: Array<Link>): JSX.Element[] => {
    return links.map((link: Link, index: number): JSX.Element => {
      return (
        <div key={`editor_link_${index}`} className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, '_blank')}>
            <Icon iconName={link.icon ? link.icon : props.defaultIcon} className={styles.icon} />
            <span title={link.displayText}>{link.displayText}</span>
          </Text>
          <Icon
            className={styles.actionIcon}
            iconName='CirclePlus'
            onClick={() => appendToFavourites(link)}
          />
        </div>
      )
    })
  }

  const generateMandatoryLinkComponents = (links: Array<Link>): JSX.Element[] => {
    return links.map((link: Link, index: number): JSX.Element => {
      return (
        <div key={`required_link_${index}`} className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, '_blank')}>
            <Icon iconName={link.icon ? link.icon : props.defaultIcon} className={styles.icon} />
            <span>{link.displayText}</span>
          </Text>
        </div>
      )
    })
  }

  const generateFavouriteLinkComponents = (links: Array<Link>): JSX.Element[] => {
    return links.map((link: Link, index: number): JSX.Element => {
      const linkIcon: JSX.Element = (
        <Icon iconName={link.icon ? link.icon : props.defaultIcon} className={styles.icon} />
      )
      const removeLinkButton: JSX.Element =
        link.linkType === LinkType.editorLink ? (
          <Icon
            className={styles.actionIcon}
            iconName='SkypeCircleMinus'
            onClick={() => removeFromFavourites(link)}
          />
        ) : (
          <Icon
            className={styles.actionIcon}
            iconName='SkypeCircleMinus'
            onClick={() => removeCustomFromFavourites(link)}
          />
        )
      return (
        <div key={`favourite_link_${index}`} className={styles.linkParent}>
          <Text className={styles.linkContainer} onClick={() => window.open(link.url, '_blank')}>
            {linkIcon}
            <div>{link.displayText}</div>
          </Text>
          {removeLinkButton}
        </div>
      )
    })
  }

  const generateLinks = (categories: Array<ICategory>): JSX.Element[] => {
    return categories.map((cat: ICategory, index: number): JSX.Element => {
      const linkItems: JSX.Element[] = cat.links.map(
        (link: ILink, subIndex: number): JSX.Element => {
          const linkIcon: JSX.Element = (
            <Icon iconName={link.icon ? link.icon : props.defaultIcon} className={styles.icon} />
          )
          const linkTarget: string = link.openInSameTab ? '_self' : '_blank'
          return (
            <div key={`link_cat_sub_${subIndex}`} className={styles.linkGridColumn}>
              <a
                className={styles.linkContainer}
                data-interception='off'
                href={link.url}
                title={link.displayText}
                target={linkTarget}
              >
                {linkIcon}
                <span>{link.displayText}</span>
              </a>
              {link.mandatory ? (
                <Icon
                  className={styles.icon}
                  iconName='Lock'
                  title={strings.ActionRemoveMandatory}
                />
              ) : (
                <Icon
                  className={styles.actionIcon}
                  iconName='CirclePlus'
                  onClick={() => appendToFavourites(link)}
                />
              )}
            </div>
          )
        }
      )
      if (props.listingByCategory) {
        return (
          <div key={`link_cat_${index}`} className={styles.categorySection}>
            <div className={styles.linkCategoryHeading}>{cat.displayText}</div>
            {linkItems}
          </div>
        )
      }
      return <div key={`link_no_cat_${index}`}>{linkItems}</div>
    })
  }

  const mandatoryLinks: JSX.Element[] = state.mandatoryLinks
    ? generateMandatoryLinkComponents(state.mandatoryLinks)
    : null
  const editorLinks: JSX.Element[] = state.editorLinks
    ? generateEditorLinkComponents(state.editorLinks)
    : null
  const favouriteLinks: JSX.Element[] = state.favouriteLinks
    ? generateFavouriteLinkComponents(state.favouriteLinks)
    : null

  const links: JSX.Element = props.listingByCategory ? (
    <div className={styles.allLinks}>
      <div className={styles.webpartHeader}>
        <span>{props.listingByCategoryTitle}</span>
      </div>
      <div className={styles.linkGrid}>{generateLinks(state.categoryLinks)}</div>
    </div>
  ) : (
    <div>
      <div className={styles.webpartHeading}>
        {stringIsNullOrEmpty(props.mandatoryLinksTitle)
          ? strings.MandatoryLinksLabel
          : props.mandatoryLinksTitle}
      </div>
      <div className={styles.editorLinksContainer}>{mandatoryLinks}</div>
      <div className={styles.webpartHeading}>
        {stringIsNullOrEmpty(props.recommendedLinksTitle)
          ? strings.RecommendedLinksLabel
          : props.recommendedLinksTitle}
      </div>
      <div className={styles.editorLinksContainer}>{editorLinks}</div>
    </div>
  )

  const yourLinks: JSX.Element = (
    <div>
      <div className={styles.webpartHeading}>
        {stringIsNullOrEmpty(props.yourLinksTitle) ? strings.YourLinksLabel : props.yourLinksTitle}
      </div>
      <div className={styles.editorLinksContainer}>{favouriteLinks}</div>
      <div className={styles.buttonRow}>
        <Dialog>
          <DialogTrigger disableButtonEnhancement>
            <Button
              title={strings.NewLinkLabel}
              className={styles.button}
              appearance='subtle'
              icon={<Icons.Add />}
              onClick={() => openNewItemModal()}
            >
              <span className={styles.label}>{strings.NewLinkLabel}</span>
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{strings.NewLinkLabel}</DialogTitle>
              <DialogContent>
                <div className={styles.modalBody}>
                  <TextField
                    label={strings.TitleLabel}
                    onChange={(_, newVal: any) => onModalValueChanged('displayText', newVal)}
                    value={state.showModal && state.modalData['displayText']}
                  />
                  <div>
                    <Field
                      label='Url'
                      validationState={state.validationError ? 'error' : 'none'}
                      validationMessage={state.validationError && strings.UrlValidationLabel}
                    >
                      <Input
                        type='url'
                        placeholder='Angi url her...'
                        onChange={(_, data): void => {
                          onModalValueChanged('url', data.value)
                          validateUrl(data.value)
                        }}
                      />
                    </Field>
                  </div>
                </div>
              </DialogContent>
              <DialogActions>
                <DialogTrigger disableButtonEnhancement>
                  <Button
                    title={strings.CancelLabel}
                    className={styles.button}
                    onClick={() => setState({ modalData: null, showModal: false })}
                  >
                    <span className={styles.label}>{strings.CancelLabel}</span>
                  </Button>
                </DialogTrigger>
                <Button
                  title={strings.AddLabel}
                  className={styles.button}
                  appearance='primary'
                  icon={<Icons.Add />}
                  onClick={() => addNewLink()}
                >
                  <span className={styles.label}>{strings.AddLabel}</span>
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </div>
  )

  return (
    <FluentProvider
      theme={customLightTheme}
      className={styles.allLinks}
      style={{ backgroundColor }}
    >
      {state.loading ? (
        <Spinner label='Laster inn lenker' />
      ) : (
        <div className={styles.allLinks} style={{ backgroundColor }}>
          {state.error && <MessageBar intent='error'>{strings.SaveErrorLabel}</MessageBar>}
          {props.yourLinksOnTop ? (
            <div>
              {yourLinks}
              {links}
            </div>
          ) : (
            <div>
              {links}
              {yourLinks}
            </div>
          )}
        </div>
      )}
    </FluentProvider>
  )
}
