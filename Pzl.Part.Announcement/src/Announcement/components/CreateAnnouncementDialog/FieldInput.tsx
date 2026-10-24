import React, { FC } from 'react'
import { Dropdown, Field, Input, Option, Switch, Textarea } from '@fluentui/react-components'
import strings from 'AnnouncementStrings'
import { IListField } from '../types'
import { IPersonOption } from '../../services/AnnouncementListService'
import { PeoplePicker } from './PeoplePicker'

export interface IFieldInputProps {
  field: IListField
  value: any
  error?: string
  disabled?: boolean
  onChange: (value: any) => void
  searchPeople: (query: string, allowGroups: boolean) => Promise<IPersonOption[]>
}

/**
 * Renders an input for a list field based on its type
 */
export const FieldInput: FC<IFieldInputProps> = (props) => {
  const { field, value, onChange, disabled } = props

  const renderControl = () => {
    switch (field.type) {
      case 'Note':
        return (
          <Textarea
            value={value ?? ''}
            rows={field.richText ? 6 : 3}
            resize='vertical'
            disabled={disabled}
            onChange={(_, data) => onChange(data.value)}
          />
        )
      case 'Choice':
        return (
          <Dropdown
            value={value ?? ''}
            selectedOptions={value ? [value] : []}
            placeholder={strings.ChoicePlaceholder}
            disabled={disabled}
            onOptionSelect={(_, data) => onChange(data.optionValue)}>
            {(field.choices ?? []).map((choice) => (
              <Option key={choice} value={choice}>
                {choice}
              </Option>
            ))}
          </Dropdown>
        )
      case 'MultiChoice': {
        const selected: string[] = Array.isArray(value) ? value : []
        return (
          <Dropdown
            multiselect
            value={selected.join(', ')}
            selectedOptions={selected}
            placeholder={strings.ChoicePlaceholder}
            disabled={disabled}
            onOptionSelect={(_, data) => onChange(data.selectedOptions)}>
            {(field.choices ?? []).map((choice) => (
              <Option key={choice} value={choice}>
                {choice}
              </Option>
            ))}
          </Dropdown>
        )
      }
      case 'DateTime':
        return (
          <Input
            type={field.dateOnly ? 'date' : 'datetime-local'}
            value={value ?? ''}
            disabled={disabled}
            onChange={(_, data) => onChange(data.value)}
          />
        )
      case 'Boolean':
        return (
          <Switch
            checked={value === true}
            disabled={disabled}
            onChange={(_, data) => onChange(data.checked)}
          />
        )
      case 'Number':
      case 'Currency':
        return (
          <Input
            type='number'
            value={value ?? ''}
            disabled={disabled}
            onChange={(_, data) => onChange(data.value)}
          />
        )
      case 'URL':
        return (
          <Input
            type='url'
            value={value ?? ''}
            disabled={disabled}
            onChange={(_, data) => onChange(data.value)}
          />
        )
      case 'User':
      case 'UserMulti':
        return (
          <PeoplePicker
            value={Array.isArray(value) ? value : []}
            multiple={field.type === 'UserMulti'}
            allowGroups={field.allowGroups}
            onChange={onChange}
            search={props.searchPeople}
          />
        )
      default:
        return (
          <Input
            value={value ?? ''}
            disabled={disabled}
            onChange={(_, data) => onChange(data.value)}
          />
        )
    }
  }

  return (
    <Field
      label={field.title}
      required={field.required}
      hint={field.description || undefined}
      validationMessage={props.error}
      validationState={props.error ? 'error' : 'none'}>
      {renderControl()}
    </Field>
  )
}
