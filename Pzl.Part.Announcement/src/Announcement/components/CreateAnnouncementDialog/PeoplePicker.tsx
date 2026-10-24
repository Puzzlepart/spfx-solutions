import React, { FC, useEffect, useState } from 'react'
import { Avatar, Combobox, Option, Persona, Tag, TagGroup, Text } from '@fluentui/react-components'
import strings from 'AnnouncementStrings'
import styles from './CreateAnnouncementDialog.module.scss'
import { IPersonOption } from '../../services/AnnouncementListService'

/**
 * Delay (ms) after the last keystroke before searching
 */
const SEARCH_DEBOUNCE = 300

export interface IPeoplePickerProps {
  id?: string
  value: IPersonOption[]
  /**
   * Allow more than one person
   */
  multiple: boolean
  /**
   * Include groups in the search
   */
  allowGroups: boolean
  onChange: (value: IPersonOption[]) => void
  search: (query: string, allowGroups: boolean) => Promise<IPersonOption[]>
}

/**
 * A simple people picker built on the Fluent UI v9 `Combobox`, searching with
 * the SharePoint people picker API.
 */
export const PeoplePicker: FC<IPeoplePickerProps> = (props) => {
  const [query, setQuery] = useState('')
  const [options, setOptions] = useState<IPersonOption[]>([])
  const [searching, setSearching] = useState(false)

  useEffect(() => {
    if (query.trim().length < 2) {
      setOptions([])
      setSearching(false)
      return
    }
    setSearching(true)
    let cancelled = false
    const timer = setTimeout(() => {
      props
        .search(query, props.allowGroups)
        .then((result) => {
          if (!cancelled) setOptions(result)
        })
        .catch(() => {
          if (!cancelled) setOptions([])
        })
        .then(() => {
          if (!cancelled) setSearching(false)
        })
    }, SEARCH_DEBOUNCE)
    return () => {
      cancelled = true
      clearTimeout(timer)
    }
  }, [query])

  const select = (key: string) => {
    const option = options.find((o) => o.key === key)
    if (!option) return
    if (props.value.some((v) => v.key === option.key)) return
    props.onChange(props.multiple ? [...props.value, option] : [option])
    setQuery('')
    setOptions([])
  }

  const remove = (key: string) => props.onChange(props.value.filter((v) => v.key !== key))

  const inputDisabled = !props.multiple && props.value.length > 0

  return (
    <div className={styles.peoplePicker}>
      {props.value.length > 0 && (
        <TagGroup
          onDismiss={(_, data) => remove(data.value)}
          aria-label={strings.FieldTitles.PzlResponsible}>
          {props.value.map((person) => (
            <Tag
              key={person.key}
              value={person.key}
              dismissible
              media={<Avatar name={person.text} color='colorful' />}>
              {person.text}
            </Tag>
          ))}
        </TagGroup>
      )}
      {!inputDisabled && (
        <Combobox
          id={props.id}
          freeform
          clearable
          value={query}
          selectedOptions={[]}
          placeholder={strings.PeoplePickerPlaceholder}
          onChange={(event) => setQuery(event.target.value)}
          onOptionSelect={(_, data) => select(data.optionValue)}>
          {options.map((option) => (
            <Option key={option.key} value={option.key} text={option.text}>
              <Persona
                name={option.text}
                secondaryText={option.email}
                avatar={{ color: 'colorful' }}
                size='small'
              />
            </Option>
          ))}
          {options.length === 0 && query.trim().length >= 2 && (
            <Option key='empty' value='' text='' disabled>
              <Text size={200}>
                {searching ? strings.PeoplePickerSearchingText : strings.PeoplePickerNoResultsText}
              </Text>
            </Option>
          )}
        </Combobox>
      )}
    </div>
  )
}
