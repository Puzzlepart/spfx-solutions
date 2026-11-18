import { useEffect, useRef, useState } from 'react'
import strings from 'AnnouncementStrings'
import { IListField } from '../types'
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
 * Converts the form values to the REST payload for `items.add`
 */
async function buildPayload(
  fields: IListField[],
  values: Record<string, any>,
  service: AnnouncementListService
): Promise<Record<string, any>> {
  const payload: Record<string, any> = {}
  for (const field of fields) {
    const value = values[field.internalName]
    const name = field.entityPropertyName
    switch (field.type) {
      case 'Boolean':
        payload[name] = value === true
        break
      case 'DateTime': {
        const iso = dateInputValueToIso(value)
        if (iso) payload[name] = iso
        break
      }
      case 'Number':
      case 'Currency':
        if (hasValue(field, value)) payload[name] = Number(value)
        break
      case 'URL':
        if (hasValue(field, value)) payload[name] = { Url: value, Description: value }
        break
      case 'MultiChoice':
        if (hasValue(field, value)) payload[name] = value
        break
      case 'User': {
        if (hasValue(field, value)) {
          payload[`${name}Id`] = await service.ensureUserId(value[0].key)
        }
        break
      }
      case 'UserMulti': {
        if (hasValue(field, value)) {
          payload[`${name}Id`] = await Promise.all(
            value.map((person: { key: string }) => service.ensureUserId(person.key))
          )
        }
        break
      }
      default:
        if (hasValue(field, value)) payload[name] = value
    }
  }
  return payload
}

/**
 * Form logic for `CreateAnnouncementDialog`: values, validation and save.
 *
 * @param open Whether the dialog is open (the form is reset when it opens)
 * @param allFields All list fields
 * @param hiddenFormFields Fields hidden in the web part properties (`false` = hidden)
 * @param alwaysNotified Levels the notification add-on always notifies about
 * @param service List service
 * @param onCreated Called after the item has been created
 */
export function useCreateAnnouncementForm(
  open: boolean,
  allFields: IListField[],
  hiddenFormFields: Record<string, boolean> | undefined,
  alwaysNotified: Severity[],
  service: AnnouncementListService,
  onCreated: () => void
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
      initial[field.internalName] = getInitialFormValue(field)
    })
    bannerTouched.current = false
    if (hasBannerField && isCriticalValue(initial[SEVERITY_FIELD])) {
      initial[SHOW_AS_BANNER_FIELD] = true
    }
    setValues(initial)
    setErrors({})
    setSaveError(null)
    setSaving(false)
  }, [open, allFields])

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
      const payload = await buildPayload(formFields, shownValues, service)
      await service.addItem(payload)
      onCreated()
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
