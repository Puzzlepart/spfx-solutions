import * as React from 'react'
import styles from './AllLinks.module.scss'
import { IAllLinksProps, LinkType, ILink, ICategory } from './types'
import { Icon } from 'office-ui-fabric-react/lib/Icon'
import * as strings from 'AllLinksWebPartStrings'
import { stringIsNullOrEmpty } from '@pnp/common'
import { IconPicker } from '@pnp/spfx-controls-react/lib/IconPicker'
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
  IdPrefixProvider,
  InfoLabel,
  Input,
  MessageBar,
  Spinner,
  SplitButton,
  useId
} from '@fluentui/react-components'
import { useAllLinks } from './useAllLinks'
import { Icons } from '../../../util/icons'

export const AllLinks: React.FC<IAllLinksProps> = (props) => {
  const {
    state,
    setState,
    backgroundColor,
    openNewLinkDialog,
    appendToFavourites,
    removeFromFavourites,
    removeCustomFromFavourites,
    addNewLink,
    onDialogValueChanged,
    validateUrl,
    theme
  } = useAllLinks(props)
  const fluentProviderId = useId('fp-all-links')

  const generateEditorLinks = (links: Array<ILink>) => {
    return links.map((link: ILink, idx: number) => {
      return (
        <SplitButton
          key={`editor_link_${idx}`}
          title={link.displayText}
          className={styles.link}
          icon={
            <Icon className={styles.icon} iconName={link.icon ? link.icon : props.defaultIcon} />
          }
          menuIcon={null}
          menuButton={{
            style: { width: '30px' },
            children: (
              <Button
                title={`Legg til ${link.displayText} i dine lenker`}
                appearance='transparent'
                size='small'
                icon={<Icons.Add />}
              />
            ),
            onClick: () => appendToFavourites(link)
          }}
          primaryActionButton={{
            onClick: () => {
              window.open(link.url, link.openInSameTab ? '_self' : '_blank')
            }
          }}
        >
          <span className={styles.label}>{link.displayText}</span>
        </SplitButton>
      )
    })
  }

  const generateMandatoryLinks = (links: Array<ILink>) => {
    return links.map((link: ILink, idx: number) => {
      return (
        <Button
          key={`mandatory_link_${idx}`}
          title={link.displayText}
          className={styles.link}
          icon={
            <Icon className={styles.icon} iconName={link.icon ? link.icon : props.defaultIcon} />
          }
          onClick={() => {
            window.open(link.url, link.openInSameTab ? '_self' : '_blank')
          }}
        >
          <span className={styles.label}>{link.displayText}</span>
        </Button>
      )
    })
  }

  const generateFavouriteLinks = (links: Array<ILink>) => {
    return links.map((link: ILink, idx: number) => {
      return (
        <SplitButton
          key={`favourite_link_${idx}`}
          title={link.displayText}
          className={styles.link}
          icon={
            <Icon className={styles.icon} iconName={link.icon ? link.icon : props.defaultIcon} />
          }
          menuIcon={null}
          menuButton={{
            style: { width: '30px' },
            children: (
              <Button
                title={`Fjern ${link.displayText} fra dine lenker`}
                appearance='transparent'
                size='small'
                icon={<Icons.Subtract />}
              />
            ),
            onClick: () => {
              link.linkType === LinkType.editorLink
                ? removeFromFavourites(link)
                : removeCustomFromFavourites(link)
            }
          }}
          primaryActionButton={{
            onClick: () => {
              window.open(link.url, link.openInSameTab ? '_self' : '_blank')
            }
          }}
        >
          <span className={styles.label}>{link.displayText}</span>
        </SplitButton>
      )
    })
  }

  const generateCategorizedLinks = (categories: Array<ICategory>) => {
    return categories?.map((category: ICategory, idx: number) => {
      const linkItems = category.links.map((link: ILink, subIdx: number) => {
        return (
          <SplitButton
            key={`link_${subIdx}`}
            title={link.displayText}
            className={styles.link}
            icon={
              <Icon className={styles.icon} iconName={link.icon ? link.icon : props.defaultIcon} />
            }
            menuIcon={null}
            menuButton={{
              style: { width: '30px' },
              children: (
                <Button
                  key={`link_${subIdx}`}
                  title={
                    link.mandatory
                      ? strings.ActionRemoveMandatory
                      : `Legg til ${link.displayText} i dine lenker`
                  }
                  appearance='transparent'
                  size='small'
                  icon={link.mandatory ? <Icons.Lock /> : <Icons.Add />}
                  disabled={link.mandatory}
                />
              ),
              onClick: () => appendToFavourites(link)
            }}
            primaryActionButton={{
              onClick: () => {
                window.open(link.url, link.openInSameTab ? '_self' : '_blank')
              }
            }}
          >
            <span className={styles.label}>{link.displayText}</span>
          </SplitButton>
        )
      })

      if (props.groupByCategory) {
        return (
          <div className={styles.categorySection} key={`category_${idx}`}>
            <div className={styles.heading}>
              {category.displayText !== undefined ? category.displayText : 'Mine lenker'}
            </div>
            <div key={`links_${idx}`} className={styles.links}>
              {linkItems}
            </div>
          </div>
        )
      }

      return (
        <div key={`links_${idx}`} className={styles.links}>
          {linkItems}
        </div>
      )
    })
  }
  const links = props.groupByCategory ? (
    <div className={styles.links}>{generateCategorizedLinks(state.categoryLinks)}</div>
  ) : (
    <>
      <div style={{ display: props.hideMandatoryLinks && 'none' }}>
        <InfoLabel
          className={styles.linksTitle}
          info={
            stringIsNullOrEmpty(props.mandatoryLinksDescription)
              ? strings.MandatoryLinksDescription
              : props.mandatoryLinksDescription
          }
        >
          <span>
            {stringIsNullOrEmpty(props.mandatoryLinksTitle)
              ? strings.MandatoryLinksLabel
              : props.mandatoryLinksTitle}
          </span>
        </InfoLabel>
        {state.mandatoryLinks && (
          <div className={styles.links}>{generateMandatoryLinks(state.mandatoryLinks)}</div>
        )}
      </div>
      <div style={{ display: props.hideRecommendedLinks && 'none' }}>
        <InfoLabel
          className={styles.linksTitle}
          info={
            stringIsNullOrEmpty(props.recommendedLinksDescription)
              ? strings.RecommendedLinksDescription
              : props.recommendedLinksDescription
          }
        >
          <span>
            {stringIsNullOrEmpty(props.recommendedLinksTitle)
              ? strings.RecommendedLinksLabel
              : props.recommendedLinksTitle}
          </span>
        </InfoLabel>
        {state.editorLinks && (
          <div className={styles.links}>{generateEditorLinks(state.editorLinks)}</div>
        )}
      </div>
    </>
  )

  const yourLinks = (
    <div style={{ display: props.hideYourLinks && 'none' }}>
      <InfoLabel
        className={styles.linksTitle}
        info={
          stringIsNullOrEmpty(props.yourLinksDescription)
            ? strings.YourLinksDescription
            : props.yourLinksDescription
        }
      >
        <span>
          {stringIsNullOrEmpty(props.yourLinksTitle)
            ? strings.YourLinksLabel
            : props.yourLinksTitle}
        </span>
      </InfoLabel>
      {state.favouriteLinks && (
        <div className={styles.links}>{generateFavouriteLinks(state.favouriteLinks)}</div>
      )}
      <div className={styles.footer}>
        <Dialog>
          <DialogTrigger disableButtonEnhancement>
            <Button
              title={strings.NewLinkLabel}
              appearance='subtle'
              className={styles.footerButton}
              icon={<Icons.Add20 />}
              onClick={() => openNewLinkDialog()}
            >
              <span className={styles.footerButtonLabel}>{strings.NewLinkLabel}</span>
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{strings.NewLinkLabel}</DialogTitle>
              <DialogContent className={styles.dialogContent}>
                <Field label={strings.TitleLabel}>
                  <Input
                    placeholder={strings.TitlePlaceholder}
                    onChange={(_, data): void => onDialogValueChanged('displayText', data.value)}
                  />
                </Field>
                <Field
                  label='Url'
                  validationState={state.validationError ? 'error' : 'none'}
                  validationMessage={state.validationError && strings.UrlValidationLabel}
                >
                  <Input
                    type='url'
                    placeholder={strings.UrlPlaceholder}
                    onChange={(_, data): void => {
                      onDialogValueChanged('url', data.value)
                      validateUrl(data.value)
                    }}
                  />
                </Field>
                <Field label={strings.IconLabel}>
                  <div className={styles.iconField}>
                    <IconPicker
                      useStartsWithSearch
                      buttonLabel={strings.IconButtonLabel}
                      currentIcon={state.dialogData?.icon}
                      onChange={(icon: string) => {
                        onDialogValueChanged('icon', icon)
                      }}
                      panelClassName='iconPickerPanel'
                      onSave={(icon: string) => {
                        onDialogValueChanged('icon', icon)
                      }}
                    />
                    <MessageBar className={styles.iconMessage} intent='info' icon={null}>
                      <div className={styles.selectedIcon}>
                        {strings.SelectedIconLabel}
                        <Icon
                          iconName={
                            state.dialogData?.icon ? state.dialogData?.icon : props.defaultIcon
                          }
                        />
                        {`(${state.dialogData?.icon ? state.dialogData?.icon : props.defaultIcon})`}
                      </div>
                    </MessageBar>
                  </div>
                </Field>
              </DialogContent>
              <DialogActions>
                <DialogTrigger disableButtonEnhancement>
                  <Button
                    title={strings.CancelLabel}
                    onClick={() => setState({ dialogData: null, showDialog: false })}
                  >
                    <span className={styles.label}>{strings.CancelLabel}</span>
                  </Button>
                </DialogTrigger>
                <DialogTrigger disableButtonEnhancement>
                  <Button
                    title={strings.AddLabel}
                    appearance='primary'
                    icon={<Icons.Add />}
                    onClick={() => addNewLink()}
                  >
                    <span className={styles.label}>{strings.AddLabel}</span>
                  </Button>
                </DialogTrigger>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </div>
  )

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider theme={theme} className={styles.allLinks} style={{ backgroundColor }}>
        {state.loading ? (
          <Spinner label='Laster inn lenker' />
        ) : (
          <div className={styles.allLinks}>
            {state.error && <MessageBar intent='error'>{strings.SaveErrorLabel}</MessageBar>}
            {yourLinks}
            {links}
          </div>
        )}
      </FluentProvider>
    </IdPrefixProvider>
  )
}

AllLinks.defaultProps = {
  defaultIcon: 'Link'
}
