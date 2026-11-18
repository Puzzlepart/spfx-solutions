import { useEffect, useRef, useState } from 'react'
import strings from 'AnnouncementStrings'
import { IAnnouncement, IListField } from '../types'
import { AnnouncementListService, isFormField } from '../../services/AnnouncementListService'
import { format } from '@fluentui/react'
import {
  isFormFieldEnabled,
  NOTIFY_TEAMS_FIELD,
  SEVERITY_FIELD,
  SHOW_AS_BANNER_FIELD
} from '../../constants'
import { isAlwaysNotified, isCriticalValue, Severity } from '../../model'
import {
  dateInputValueToIso,
  getFormValueFromItem,
  getInitialFormValue,
  htmlToText,
  isEmptyValue
} from '../../util/fieldValue'

/**
 * Whether a form value counts as filled in (for required validation)
 */
function hasValue(field: IListField, value: any): boolean {
  if (field.type === 'Boolean') return true
  if (field.type === 'Note' && field.richText) return htmlToText(value).length > 0
  if (typeof value === 'string') return value.trim().length > 0
  return !isEmptyValue(value)
}

/**
 * Converts the form values to the REST payload for `items.add`, or for
 * `update` when the values of the existing item are given: empty fields are
 * then cleared on the item instead of left out.
 *
 * @param fields Fields in the form
 * @param values Form values keyed by internal name
 * @param service List service
 * @param original Values of the item being edited, keyed by entity property name
 */
async function buildPayload(
  fields: IListField[],
  values: Record<string, any>,
  service: AnnouncementListService,
  original?: Record<string, any>
): Promise<Record<string, any>> {
  const payload: Record<string, any> = {}
  const userId = (person: { key: string; id?: number }) =>
    person.id ? Promise.resolve(person.id) : service.ensureUserId(person.key)
  for (const field of fields) {
    const value = values[field.internalName]
    const name = field.entityPropertyName
    const filled = hasValue(field, value)
    /** Sets the field, or clears it when an existing item is edited */
    const set = (key: string, filledValue: () => any, emptyValue: any = null) => {
      if (filled) payload[key] = filledValue()
      else if (original) payload[key] = emptyValue
    }
    switch (field.type) {
      case 'Boolean':
        payload[name] = value === true
        break
      case 'DateTime': {
        const iso = dateInputValueToIso(value)
        if (iso) payload[name] = iso
        else if (original) payload[name] = null
        break
      }
      case 'Number':
      case 'Currency':
        set(name, () => Number(value))
        break
      case 'URL':
        // The description of an unchanged link is kept
        set(name, () => ({
          Url: value,
          Description: original?.[name]?.Url === value ? original[name].Description : value
        }))
        break
      case 'MultiChoice':
        set(name, () => value, [])
        break
      case 'User':
        if (filled) payload[`${name}Id`] = await userId(value[0])
        else if (original) payload[`${name}Id`] = null
        break
      case 'UserMulti':
        if (filled) payload[`${name}Id`] = await Promise.all(value.map(userId))
        else if (original) payload[`${name}Id`] = []
        break
      default:
        set(name, () => value)
    }
  }
  return payload
}

/**
 * Form logic for `CreateAnnouncementDialog`: values, validation and save.
 * Creates a new announcement, or updates the one given in `announcement`.
 *
 * @param open Whether the dialog is open (the form is reset when it opens)
 * @param allFields All list fields
 * @param hiddenFormFields Fields hidden in the web part properties (`false` = hidden)
 * @param alwaysNotified Levels the notification add-on always notifies about
 * @param service List service
 * @param onSaved Called after the item has been created or updated
 * @param announcement The announcement to edit. Without it a new one is created.
 */
export function useCreateAnnouncementForm(
  open: boolean,
  allFields: IListField[],
  hiddenFormFields: Record<string, boolean> | undefined,
  alwaysNotified: Severity[],
  service: AnnouncementListService,
  onSaved: () => void,
  announcement?: IAnnouncement
) {
  const formFields = allFields.filter(
    (field) =>
      isFormField(field) &&
      (field.required || isFormFieldEnabled(field.internalName, hiddenFormFields))
  )
  const [values, setValues] = useState<Record<string, any>>({})
  const [errors, setErrors] = useState<Record<string, string>>({})
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string>(null)
  /**
   * Whether the user has changed «Vis som banner» in this form. Until then it
   * follows the message level: on for «Kritisk», off otherwise.
   */
  const bannerTouched = useRef(false)
  const hasBannerField = formFields.some((field) => field.internalName === SHOW_AS_BANNER_FIELD)

  // Reset the form each time the dialog opens
  useEffect(() => {
    if (!open) return
    const initial: Record<string, any> = {}
    formFields.forEach((field) => {
      initial[field.internalName] = announcement
        ? getFormValueFromItem(field, announcement.values[field.entityPropertyName])
        : getInitialFormValue(field)
    })
    // An existing announcement keeps its own «Vis som banner»
    bannerTouched.current = !!announcement
    if (!announcement && hasBannerField && isCriticalValue(initial[SEVERITY_FIELD])) {
      initial[SHOW_AS_BANNER_FIELD] = true
    }
    setValues(initial)
    setErrors({})
    setSaveError(null)
    setSaving(false)
  }, [open, allFields, announcement?.id])

  /**
   * «Varsle i Teams» is selected and locked while the chosen level is one of
   * «Nivåer som alltid varsles»: the service notifies about it anyway. The
   * user's own choice is kept in `values` and comes back for other levels.
   */
  const notifyLocked =
    formFields.some((field) => field.internalName === NOTIFY_TEAMS_FIELD) &&
    isAlwaysNotified(values[SEVERITY_FIELD], alwaysNotified)
  const shownValues = notifyLocked ? { ...values, [NOTIFY_TEAMS_FIELD]: true } : values
  /**
   * Fields that cannot be changed, with the explanation shown under them
   */
  const lockedFields: Record<string, string> = notifyLocked
    ? {
        [NOTIFY_TEAMS_FIELD]: format(
          strings.Notification.AlwaysNotifiedHint,
          values[SEVERITY_FIELD]
        )
      }
    : {}

  const setValue = (internalName: string, value: any) => {
    if (internalName === SHOW_AS_BANNER_FIELD) bannerTouched.current = true
    const preselectBanner =
      internalName === SEVERITY_FIELD && hasBannerField && !bannerTouched.current
    setValues((current) => ({
      ...current,
      [internalName]: value,
      ...(preselectBanner ? { [SHOW_AS_BANNER_FIELD]: isCriticalValue(value) } : {})
    }))
    setErrors((current) => {
      if (!current[internalName]) return current
      const next = { ...current }
      delete next[internalName]
      return next
    })
  }

  const validate = (): boolean => {
    const next: Record<string, string> = {}
    formFields.forEach((field) => {
      if (field.required && !hasValue(field, values[field.internalName])) {
        next[field.internalName] = strings.RequiredFieldText
      }
    })
    setErrors(next)
    return Object.keys(next).length === 0
  }

  const submit = async () => {
    if (saving || !validate()) return
    setSaving(true)
    setSaveError(null)
    try {
      const payload = await buildPayload(formFields, shownValues, service, announcement?.values)
      if (announcement) await service.updateItem(announcement.id, payload)
      else await service.addItem(payload)
      onSaved()
    } catch {
      setSaveError(strings.CreateErrorText)
    } finally {
      setSaving(false)
    }
  }

  return {
    formFields,
    values: shownValues,
    lockedFields,
    errors,
    saving,
    saveError,
    setValue,
    submit
  } as const
}
