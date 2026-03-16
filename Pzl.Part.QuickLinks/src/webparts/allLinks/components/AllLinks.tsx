import * as React from 'react'
import styles from './AllLinks.module.scss'
import { IAllLinksProps, LinkType, ILink, ICategory } from './types'
import { Icon } from '@fluentui/react/lib/Icon'
import * as strings from 'AllLinksWebPartStrings'
import { isNullOrEmpty } from '../../../util/string'
import {
  Button,
  Checkbox,
  Combobox,
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
  Option,
  Spinner,
  SplitButton,
  useId
} from '@fluentui/react-components'
import { useAllLinks } from './useAllLinks'
import { Icons } from '../../../util/icons'
import { fluentIconNames } from '../../../util/fluentIconNames'

export const AllLinks: React.FC<IAllLinksProps> = (props) => {
  const [iconSearch, setIconSearch] = React.useState('')
  const [editorIconSearch, setEditorIconSearch] = React.useState('')
  const [editorDialogOpen, setEditorDialogOpen] = React.useState(false)
  const [editorIsSaving, setEditorIsSaving] = React.useState(false)
  const [editorValidationError, setEditorValidationError] = React.useState(false)
  const [editorSaveError, setEditorSaveError] = React.useState('')
  const [recentlyAddedLinkTokens, setRecentlyAddedLinkTokens] = React.useState<string[]>([])
  const {
    state,
    setState,
    backgroundColor,
    openNewLinkDialog,
    appendToFavourites,
    removeFromFavourites,
    removeCustomFromFavourites,
    addNewLink,
    addEditorLink,
    onDialogValueChanged,
    validateUrl,
    theme
  } = useAllLinks(props)
  const fluentProviderId = useId('fp-all-links')
  const createEmptyEditorLink = React.useCallback(
    (): ILink => ({
      displayText: '',
      url: '',
      icon: props.defaultIcon,
      category: '',
      priority: '1000',
      active: true,
      mandatory: false,
      linkType: LinkType.editorLink
    }),
    [props.defaultIcon]
  )
  const [editorDialogData, setEditorDialogData] = React.useState<ILink>(createEmptyEditorLink)
  const formatLabel = (template: string, value: string) => template.replace('{0}', value)
  const selectedIconName = state.dialogData?.icon || props.defaultIcon
  const selectedEditorIconName = editorDialogData.icon || props.defaultIcon
  const iconSearchValue = iconSearch.trim().toLowerCase()
  const editorIconSearchValue = editorIconSearch.trim().toLowerCase()
  const filteredIconNames = fluentIconNames
    .filter((iconName) => {
      if (!iconSearchValue) return true
      return iconName.toLowerCase().includes(iconSearchValue)
    })
    .sort((left, right) => {
      const leftStartsWith = left.toLowerCase().startsWith(iconSearchValue)
      const rightStartsWith = right.toLowerCase().startsWith(iconSearchValue)

      if (leftStartsWith === rightStartsWith) {
        return left.localeCompare(right)
      }

      return leftStartsWith ? -1 : 1
    })
  const filteredEditorIconNames = fluentIconNames
    .filter((iconName) => {
      if (!editorIconSearchValue) return true
      return iconName.toLowerCase().includes(editorIconSearchValue)
    })
    .sort((left, right) => {
      const leftStartsWith = left.toLowerCase().startsWith(editorIconSearchValue)
      const rightStartsWith = right.toLowerCase().startsWith(editorIconSearchValue)

      if (leftStartsWith === rightStartsWith) {
        return left.localeCompare(right)
      }

      return leftStartsWith ? -1 : 1
    })

  const updateEditorDialogValue = (field: keyof ILink, newValue: string | boolean): void => {
    setEditorDialogData((currentData) => ({
      ...currentData,
      [field]: newValue
    }))
  }

  const getLinkToken = React.useCallback((link: ILink): string | null => {
    if (typeof link.id === 'number') {
      return `id:${link.id}`
    }

    if (link.localKey) {
      return `local:${link.localKey}`
    }

    return null
  }, [])

  const markLinkAsRecentlyAdded = React.useCallback((link: ILink): void => {
    const linkToken = getLinkToken(link)
    if (!linkToken) {
      return
    }

    setRecentlyAddedLinkTokens((currentTokens) => {
      if (currentTokens.includes(linkToken)) {
        return currentTokens
      }

      return [...currentTokens, linkToken]
    })

    window.setTimeout(() => {
      setRecentlyAddedLinkTokens((currentTokens) =>
        currentTokens.filter((currentToken) => currentToken !== linkToken)
      )
    }, 5000)
  }, [getLinkToken])

  const getLinkClassName = (link: ILink): string => {
    const classNames = [styles.link]
    const linkToken = getLinkToken(link)

    if (linkToken && recentlyAddedLinkTokens.includes(linkToken)) {
      classNames.push(styles.recentlyAddedLink)
    }

    return classNames.join(' ')
  }

  const validateEditorUrl = (value: string): boolean => {
    const trimmedValue = value.trim()

    if (!trimmedValue) {
      setEditorValidationError(false)
      return true
    }

    const urlRegex: RegExp =
      /(http|ftp|https):\/\/([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-]*[\w@?^=%&/~+#-])?/
    const isValid = urlRegex.test(trimmedValue)
    setEditorValidationError(!isValid)
    return isValid
  }

  const resetEditorDialog = (): void => {
    setEditorDialogData(createEmptyEditorLink())
    setEditorIconSearch('')
    setEditorValidationError(false)
    setEditorSaveError('')
  }

  const openEditorDialog = (): void => {
    resetEditorDialog()
    setEditorDialogOpen(true)
  }

  const closeEditorDialog = (): void => {
    if (editorIsSaving) {
      return
    }

    setEditorDialogOpen(false)
    resetEditorDialog()
  }

  const submitEditorLink = async (): Promise<void> => {
    if (editorIsSaving) {
      return
    }

    const title = editorDialogData.displayText.trim()
    const url = editorDialogData.url.trim()
    const category = editorDialogData.category?.trim() || ''
    const priority = editorDialogData.priority?.trim() || '1000'

    if (!title || !url) {
      setEditorSaveError(strings.EditorValidationLabel)
      return
    }

    if (!validateEditorUrl(url)) {
      setEditorSaveError(strings.UrlValidationLabel)
      return
    }

    setEditorIsSaving(true)
    setEditorSaveError('')

    try {
      const createdLink = await addEditorLink({
        ...editorDialogData,
        displayText: title,
        url,
        category,
        priority,
        icon: editorDialogData.icon || props.defaultIcon,
        active: editorDialogData.active ?? true,
        mandatory: !!editorDialogData.mandatory
      })

      if (!createdLink) {
        setEditorSaveError(state.errorMessage || strings.SaveErrorLabel)
        return
      }

      markLinkAsRecentlyAdded(createdLink)
      setEditorDialogOpen(false)
      resetEditorDialog()
    } finally {
      setEditorIsSaving(false)
    }
  }

  const generateEditorLinks = (links: Array<ILink>) => {
    return links.map((link: ILink, idx: number) => {
      return (
        <SplitButton
          key={`editor_link_${idx}`}
          title={link.displayText}
          className={getLinkClassName(link)}
          icon={
            <Icon className={styles.icon} iconName={link.icon ? link.icon : props.defaultIcon} />
          }
          menuIcon={null}
          menuButton={{
            style: { width: '30px' },
            children: (
              <Button
                title={formatLabel(strings.AddToYourLinksLabel, link.displayText)}
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
          className={getLinkClassName(link)}
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
          className={getLinkClassName(link)}
          icon={
            <Icon className={styles.icon} iconName={link.icon ? link.icon : props.defaultIcon} />
          }
          menuIcon={null}
          menuButton={{
            style: { width: '30px' },
            children: (
              <Button
                title={formatLabel(strings.RemoveFromYourLinksLabel, link.displayText)}
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
            className={getLinkClassName(link)}
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
                      : formatLabel(strings.AddToYourLinksLabel, link.displayText)
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
              {category.displayText !== undefined ? category.displayText : strings.YourLinksLabel}
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
            isNullOrEmpty(props.mandatoryLinksDescription)
              ? strings.MandatoryLinksDescription
              : props.mandatoryLinksDescription
          }
        >
          <span>
            {isNullOrEmpty(props.mandatoryLinksTitle)
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
            isNullOrEmpty(props.recommendedLinksDescription)
              ? strings.RecommendedLinksDescription
              : props.recommendedLinksDescription
          }
        >
          <span>
            {isNullOrEmpty(props.recommendedLinksTitle)
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
          isNullOrEmpty(props.yourLinksDescription)
            ? strings.YourLinksDescription
            : props.yourLinksDescription
        }
      >
        <span>
          {isNullOrEmpty(props.yourLinksTitle) ? strings.YourLinksLabel : props.yourLinksTitle}
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
              onClick={() => {
                setIconSearch('')
                openNewLinkDialog()
              }}
            >
              <span className={styles.footerButtonLabel}>{strings.NewLinkLabel}</span>
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{strings.NewLinkLabel}</DialogTitle>
              <DialogContent className={styles.dialogContent}>
                <MessageBar intent='info'>{strings.PersonalLinkNoticeLabel}</MessageBar>
                <Field label={strings.TitleLabel}>
                  <Input
                    placeholder={strings.TitlePlaceholder}
                    onChange={(_, data): void => onDialogValueChanged('displayText', data.value)}
                  />
                </Field>
                <Field
                  label={strings.UrlLabel}
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
                    <Field label={strings.IconSearchLabel}>
                      <Input
                        value={iconSearch}
                        placeholder={strings.IconSearchPlaceholder}
                        onChange={(_, data) => setIconSearch(data.value)}
                      />
                    </Field>
                    <div className={styles.iconPickerActions}>
                      <Button
                        appearance={
                          selectedIconName === props.defaultIcon ? 'primary' : 'secondary'
                        }
                        onClick={() => onDialogValueChanged('icon', props.defaultIcon)}
                      >
                        <span>{strings.UseDefaultIconLabel}</span>
                      </Button>
                    </div>
                    <div className={styles.iconGrid}>
                      {filteredIconNames.map((iconName) => (
                        <Button
                          key={iconName}
                          className={styles.iconOption}
                          appearance='transparent'
                          onClick={() => onDialogValueChanged('icon', iconName)}
                          title={iconName}
                          aria-label={iconName}
                          aria-pressed={selectedIconName === iconName}
                        >
                          <span className={styles.iconOptionGlyph}>
                            <Icon iconName={iconName} />
                          </span>
                        </Button>
                      ))}
                    </div>
                    {filteredIconNames.length === 0 && (
                      <MessageBar intent='warning'>{strings.NoIconsFoundLabel}</MessageBar>
                    )}
                    <MessageBar className={styles.iconMessage} intent='info' icon={null}>
                      <div className={styles.selectedIcon}>
                        {strings.SelectedIconLabel}
                        <Icon iconName={selectedIconName} />
                        {`(${selectedIconName})`}
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
                    onClick={() => {
                      const createdLink = addNewLink()
                      markLinkAsRecentlyAdded(createdLink)
                    }}
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

  const editorLinksAdmin = state.canManageEditorLinks ? (
    <div className={styles.editorSection}>
      <InfoLabel className={styles.linksTitle} info={strings.EditorSectionDescription}>
        <span>{strings.EditorSectionLabel}</span>
      </InfoLabel>
      <div className={styles.footer}>
        <Dialog open={editorDialogOpen} onOpenChange={(_, data) => setEditorDialogOpen(data.open)}>
          <DialogTrigger disableButtonEnhancement>
            <Button
              title={strings.NewSharedLinkLabel}
              appearance='subtle'
              className={styles.footerButton}
              icon={<Icons.Add20 />}
              onClick={openEditorDialog}
            >
              <span className={styles.footerButtonLabel}>{strings.NewSharedLinkLabel}</span>
            </Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>{strings.NewSharedLinkLabel}</DialogTitle>
              <DialogContent className={styles.dialogContent}>
                <Field label={strings.TitleLabel} required>
                  <Input
                    value={editorDialogData.displayText}
                    placeholder={strings.TitlePlaceholder}
                    onChange={(_, data) => {
                      updateEditorDialogValue('displayText', data.value)
                      setEditorSaveError('')
                    }}
                  />
                </Field>
                <Field
                  label={strings.UrlLabel}
                  required
                  validationState={editorValidationError ? 'error' : 'none'}
                  validationMessage={editorValidationError && strings.UrlValidationLabel}
                >
                  <Input
                    type='url'
                    value={editorDialogData.url}
                    placeholder={strings.UrlPlaceholder}
                    onChange={(_, data) => {
                      updateEditorDialogValue('url', data.value)
                      validateEditorUrl(data.value)
                      setEditorSaveError('')
                    }}
                  />
                </Field>
                <Field label={strings.CategoryLabel}>
                  <Combobox
                    freeform
                    value={editorDialogData.category || ''}
                    placeholder={strings.CategoryPlaceholder}
                    onChange={(event) =>
                      updateEditorDialogValue('category', event.target.value || '')
                    }
                    onOptionSelect={(_, data) =>
                      updateEditorDialogValue('category', data.optionText || '')
                    }
                  >
                    {(state.categoryOptions ?? []).map((categoryOption) => (
                      <Option key={categoryOption} text={categoryOption}>
                        {categoryOption}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label={strings.PriorityLabel}>
                  <Input
                    type='number'
                    value={editorDialogData.priority || '1000'}
                    placeholder={strings.PriorityPlaceholder}
                    onChange={(_, data) => updateEditorDialogValue('priority', data.value)}
                  />
                </Field>
                <div className={styles.editorOptions}>
                  <Checkbox
                    checked={editorDialogData.active ?? true}
                    label={strings.ActiveLabel}
                    onChange={(_, data) => updateEditorDialogValue('active', !!data.checked)}
                  />
                  <Checkbox
                    checked={!!editorDialogData.mandatory}
                    label={strings.MandatoryOptionLabel}
                    onChange={(_, data) =>
                      updateEditorDialogValue('mandatory', !!data.checked)
                    }
                  />
                </div>
                <Field label={strings.IconLabel}>
                  <div className={styles.iconField}>
                    <Field label={strings.IconSearchLabel}>
                      <Input
                        value={editorIconSearch}
                        placeholder={strings.IconSearchPlaceholder}
                        onChange={(_, data) => setEditorIconSearch(data.value)}
                      />
                    </Field>
                    <div className={styles.iconPickerActions}>
                      <Button
                        appearance={
                          selectedEditorIconName === props.defaultIcon ? 'primary' : 'secondary'
                        }
                        onClick={() => updateEditorDialogValue('icon', props.defaultIcon)}
                      >
                        <span>{strings.UseDefaultIconLabel}</span>
                      </Button>
                    </div>
                    <div className={styles.iconGrid}>
                      {filteredEditorIconNames.map((iconName) => (
                        <Button
                          key={iconName}
                          className={styles.iconOption}
                          appearance='transparent'
                          onClick={() => updateEditorDialogValue('icon', iconName)}
                          title={iconName}
                          aria-label={iconName}
                          aria-pressed={selectedEditorIconName === iconName}
                        >
                          <span className={styles.iconOptionGlyph}>
                            <Icon iconName={iconName} />
                          </span>
                        </Button>
                      ))}
                    </div>
                    {filteredEditorIconNames.length === 0 && (
                      <MessageBar intent='warning'>{strings.NoIconsFoundLabel}</MessageBar>
                    )}
                    <MessageBar className={styles.iconMessage} intent='info' icon={null}>
                      <div className={styles.selectedIcon}>
                        {strings.SelectedIconLabel}
                        <Icon iconName={selectedEditorIconName} />
                        {`(${selectedEditorIconName})`}
                      </div>
                    </MessageBar>
                  </div>
                </Field>
                {editorSaveError && <MessageBar intent='error'>{editorSaveError}</MessageBar>}
              </DialogContent>
              <DialogActions>
                <Button title={strings.CancelLabel} onClick={closeEditorDialog} disabled={editorIsSaving}>
                  <span className={styles.label}>{strings.CancelLabel}</span>
                </Button>
                <Button
                  title={strings.CreateSharedLinkLabel}
                  appearance='primary'
                  icon={<Icons.Add />}
                  disabled={editorIsSaving}
                  onClick={() => submitEditorLink()}
                >
                  <span className={styles.label}>
                    {editorIsSaving ? strings.LoadingLabel : strings.CreateSharedLinkLabel}
                  </span>
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </div>
  ) : null

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider theme={theme} className={styles.allLinks} style={{ backgroundColor }}>
        {state.loading ? (
          <Spinner label={strings.LoadingLabel} />
        ) : (
          <div className={styles.allLinks}>
            {state.error && (
              <MessageBar intent='error'>
                {state.errorMessage || strings.SaveErrorLabel}
              </MessageBar>
            )}
            {yourLinks}
            {links}
            {editorLinksAdmin}
          </div>
        )}
      </FluentProvider>
    </IdPrefixProvider>
  )
}

AllLinks.defaultProps = {
  defaultIcon: 'Link'
}
