import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Icon } from '@fluentui/react/lib/Icon'
import {
  Button,
  FluentProvider,
  IdPrefixProvider,
  Input,
  MessageBar,
  webLightTheme,
  useId
} from '@fluentui/react-components'
import {
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-property-pane'
import { PropertyPaneCustomField } from '@microsoft/sp-property-pane/lib/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField'
import { fluentIconNames } from '../util/fluentIconNames'
import styles from './PropertyPaneFluentIconPicker.module.scss'

interface IPropertyPaneFluentIconPickerRenderProps {
  label: string
  selectedIconLabel: string
  searchPlaceholder: string
  noIconsFoundLabel: string
  currentIcon: string
  onChange: (icon: string) => void
}

interface IPropertyPaneFluentIconPickerProps extends IPropertyPaneFluentIconPickerRenderProps {
  targetProperty: string
  key: string
}

const PropertyPaneFluentIconPickerControl: React.FC<IPropertyPaneFluentIconPickerRenderProps> = ({
  label,
  selectedIconLabel,
  searchPlaceholder,
  noIconsFoundLabel,
  currentIcon,
  onChange
}) => {
  const idPrefix = useId('pp-icon-picker')
  const [iconSearch, setIconSearch] = React.useState('')

  const filteredIconNames = React.useMemo(() => {
    const iconSearchValue = iconSearch.trim().toLowerCase()

    return fluentIconNames
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
  }, [iconSearch])

  return (
    <IdPrefixProvider value={idPrefix}>
      <FluentProvider theme={webLightTheme} className={styles.iconPickerField}>
        <label className={styles.pickerLabel}>{label}</label>
        <div className={styles.searchField}>
          <Input
            value={iconSearch}
            placeholder={searchPlaceholder}
            onChange={(_, data) => setIconSearch(data.value)}
          />
        </div>
        <div className={styles.iconGrid}>
          {filteredIconNames.map((iconName) => (
            <Button
              key={iconName}
              className={styles.iconOption}
              appearance='transparent'
              title={iconName}
              aria-label={iconName}
              aria-pressed={currentIcon === iconName}
              onClick={() => onChange(iconName)}
            >
              <span className={styles.iconOptionGlyph}>
                <Icon iconName={iconName} />
              </span>
            </Button>
          ))}
        </div>
        {filteredIconNames.length === 0 && <MessageBar intent='warning'>{noIconsFoundLabel}</MessageBar>}
        <div className={styles.selectedIcon}>
          <span>{selectedIconLabel}</span>
          <Icon iconName={currentIcon || 'Link'} />
          <span>{currentIcon || 'Link'}</span>
        </div>
      </FluentProvider>
    </IdPrefixProvider>
  )
}

export const PropertyPaneFluentIconPicker = (
  props: IPropertyPaneFluentIconPickerProps
) => {
  const onRender: IPropertyPaneCustomFieldProps['onRender'] = (domElement, _, changeCallback) => {
    const handleChange = (icon: string): void => {
      changeCallback?.(props.targetProperty, icon, true)
      props.onChange(icon)
    }

    ReactDom.render(
      <PropertyPaneFluentIconPickerControl
        label={props.label}
        selectedIconLabel={props.selectedIconLabel}
        searchPlaceholder={props.searchPlaceholder}
        noIconsFoundLabel={props.noIconsFoundLabel}
        currentIcon={props.currentIcon}
        onChange={handleChange}
      />,
      domElement
    )
  }

  const onDispose: IPropertyPaneCustomFieldProps['onDispose'] = (domElement) => {
    ReactDom.unmountComponentAtNode(domElement)
  }

  return PropertyPaneCustomField({
    key: props.key,
    onRender,
    onDispose
  }) as any
}