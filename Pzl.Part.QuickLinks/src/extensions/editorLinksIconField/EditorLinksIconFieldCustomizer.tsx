import * as React from 'react'
import * as ReactDom from 'react-dom'
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters,
  ListItemAccessor
} from '@microsoft/sp-listview-extensibility'
import { Icon } from '@fluentui/react/lib/Icon'
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Field,
  FluentProvider,
  IdPrefixProvider,
  Input,
  MessageBar,
  useId,
  webLightTheme
} from '@fluentui/react-components'
import * as strings from 'EditorLinksIconFieldCustomizerStrings'
import { fluentIconNames } from '../../util/fluentIconNames'
import styles from './EditorLinksIconFieldCustomizer.module.scss'

const iconFieldInternalName = 'PzlOfficeUIFabricIcon'

interface IIconFieldCellProps {
  itemId: number
  listId: string
  spHttpClient: SPHttpClient
  webUrl: string
  value?: string
}

const IconFieldCell: React.FC<IIconFieldCellProps> = ({
  itemId,
  listId,
  spHttpClient,
  webUrl,
  value
}) => {
  const fluentProviderId = useId(`fp-editor-links-icon-${itemId}`)
  const [dialogOpen, setDialogOpen] = React.useState(false)
  const [iconSearch, setIconSearch] = React.useState('')
  const [selectedIcon, setSelectedIcon] = React.useState(value || '')
  const [currentIcon, setCurrentIcon] = React.useState(value || '')
  const [isSaving, setIsSaving] = React.useState(false)
  const [errorMessage, setErrorMessage] = React.useState('')

  const iconSearchValue = iconSearch.trim().toLowerCase()
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

  const openDialog = (): void => {
    setSelectedIcon(currentIcon)
    setIconSearch('')
    setErrorMessage('')
    setDialogOpen(true)
  }

  const closeDialog = (): void => {
    if (isSaving) return
    setDialogOpen(false)
    setErrorMessage('')
  }

  const saveIcon = async (): Promise<void> => {
    setIsSaving(true)
    setErrorMessage('')

    try {
      const response: SPHttpClientResponse = await spHttpClient.post(
        `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'content-type': 'application/json;odata.metadata=none',
            'if-match': '*',
            'x-http-method': 'MERGE'
          },
          body: JSON.stringify({
            [iconFieldInternalName]: selectedIcon || ''
          })
        }
      )

      if (!response.ok) {
        throw new Error('Save failed')
      }

      setCurrentIcon(selectedIcon)
      setDialogOpen(false)
    } catch {
      setErrorMessage(strings.SaveErrorLabel)
    } finally {
      setIsSaving(false)
    }
  }

  const displayedIcon = currentIcon || 'Link'

  return (
    <IdPrefixProvider value={fluentProviderId}>
      <FluentProvider theme={webLightTheme} className={styles.editorLinksIconField}>
        <Button
          className={styles.trigger}
          appearance='transparent'
          title={currentIcon || strings.SelectIconLabel}
          aria-label={currentIcon || strings.SelectIconLabel}
          onClick={openDialog}
        >
          <span className={styles.triggerGlyph}>
            <Icon iconName={displayedIcon} />
          </span>
        </Button>
        <Dialog open={dialogOpen} onOpenChange={(_, data) => !data.open && closeDialog()}>
          <DialogSurface className={styles.dialogSurface}>
            <DialogBody>
              <DialogTitle>{strings.DialogTitle}</DialogTitle>
              <DialogContent className={styles.dialogContent}>
                <Field label={strings.IconSearchLabel}>
                  <Input
                    value={iconSearch}
                    placeholder={strings.IconSearchPlaceholder}
                    onChange={(_, data) => setIconSearch(data.value)}
                  />
                </Field>
                <div className={styles.pickerActions}>
                  <Button appearance='secondary' onClick={() => setSelectedIcon('')}>
                    <span>{strings.ClearIconLabel}</span>
                  </Button>
                </div>
                <div className={styles.iconGrid}>
                  {filteredIconNames.map((iconName) => (
                    <Button
                      key={iconName}
                      className={styles.iconOption}
                      appearance='transparent'
                      aria-label={iconName}
                      aria-pressed={selectedIcon === iconName}
                      title={iconName}
                      onClick={() => setSelectedIcon(iconName)}
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
                <MessageBar icon={null}>
                  <div className={styles.selectedIcon}>
                    {strings.SelectedIconLabel}
                    <Icon iconName={selectedIcon || 'Link'} />
                    {selectedIcon || 'Link'}
                  </div>
                </MessageBar>
                {errorMessage && <div className={styles.error}>{errorMessage}</div>}
              </DialogContent>
              <DialogActions>
                <Button appearance='secondary' onClick={closeDialog} disabled={isSaving}>
                  <span>{strings.CancelLabel}</span>
                </Button>
                <Button appearance='primary' onClick={saveIcon} disabled={isSaving}>
                  <span>{isSaving ? strings.LoadingLabel : strings.SaveLabel}</span>
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </FluentProvider>
    </IdPrefixProvider>
  )
}

export default class EditorLinksIconFieldCustomizer extends BaseFieldCustomizer<
  Record<string, never>
> {
  public onInit(): Promise<void> {
    return Promise.resolve()
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const itemId = Number((event.listItem as ListItemAccessor).getValueByName('ID'))
    const listId = this.context.pageContext.list.id.toString()
    const webUrl = this.context.pageContext.web.absoluteUrl

    ReactDom.render(
      <IconFieldCell
        itemId={itemId}
        listId={listId}
        spHttpClient={this.context.spHttpClient}
        webUrl={webUrl}
        value={event.fieldValue}
      />,
      event.domElement
    )
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDom.unmountComponentAtNode(event.domElement)
    super.onDisposeCell(event)
  }
}
