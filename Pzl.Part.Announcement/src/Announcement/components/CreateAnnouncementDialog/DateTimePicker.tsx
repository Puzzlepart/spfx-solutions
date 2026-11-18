import React, { FC, useRef, useState } from 'react'
import {
  Button,
  Divider,
  Input,
  Label,
  Popover,
  PopoverSurface,
  Select,
  useId
} from '@fluentui/react-components'
import { Calendar, CalendarStrings, DayOfWeek } from '@fluentui/react-calendar-compat'
import { CalendarMonthRegular } from '@fluentui/react-icons'
import strings from 'AnnouncementStrings'
import styles from './CreateAnnouncementDialog.module.scss'
import { formatDateInputValue, parseDateInputValue, toDateInputValue } from '../../util/fieldValue'

/**
 * Minutes that can be chosen (the value's own minute is added if it is not one of them)
 */
const MINUTE_STEP = 5

const pad = (n: number) => (n < 10 ? `0${n}` : `${n}`)

const capitalize = (text: string) => text.charAt(0).toUpperCase() + text.slice(1)

/**
 * Month and day names in Norwegian for the calendar, from `Intl`
 */
function getCalendarStrings(): CalendarStrings {
  const name = (options: Intl.DateTimeFormatOptions, date: Date) =>
    capitalize(new Intl.DateTimeFormat('nb-NO', options).format(date))
  const months = Array.from({ length: 12 }, (_, month) => new Date(2024, month, 1))
  // 7 January 2024 is a Sunday: the calendar wants the days from Sunday
  const days = Array.from({ length: 7 }, (_, day) => new Date(2024, 0, 7 + day))
  return {
    months: months.map((date) => name({ month: 'long' }, date)),
    shortMonths: months.map((date) => name({ month: 'short' }, date).replace('.', '')),
    days: days.map((date) => name({ weekday: 'long' }, date)),
    shortDays: days.map((date) => name({ weekday: 'narrow' }, date)),
    goToToday: strings.DatePicker.GoToToday,
    prevMonthAriaLabel: strings.DatePicker.PrevMonth,
    nextMonthAriaLabel: strings.DatePicker.NextMonth,
    prevYearAriaLabel: strings.DatePicker.PrevYear,
    nextYearAriaLabel: strings.DatePicker.NextYear
  }
}

export interface IDateTimePickerProps {
  /**
   * `yyyy-mm-dd` (date only) or `yyyy-mm-ddThh:mm`, local time. Empty for no value.
   */
  value: string
  dateOnly: boolean
  /**
   * Required fields cannot be emptied
   */
  required?: boolean
  disabled?: boolean
  onChange: (value: string) => void
}

/**
 * Date (and time) picker for the form: a calendar, the time of day and an OK
 * button that closes the picker. The browser's own date and time picker has no
 * such button, and has to be closed by clicking outside it.
 */
export const DateTimePicker: FC<IDateTimePickerProps> = ({
  value,
  dateOnly,
  required,
  disabled,
  onChange
}) => {
  const [open, setOpen] = useState(false)
  /**
   * The field. The picker is placed under it, and focus returns to it when the
   * picker is closed with OK or the keyboard.
   */
  const fieldRef = useRef<HTMLSpanElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  const close = () => {
    setOpen(false)
    inputRef.current?.focus()
  }
  const timeLabelId = useId('announcement-time')
  const date = parseDateInputValue(value)

  /**
   * The value with a new date, keeping the time of day. Without a value the
   * time is now, rounded down to the minute step.
   */
  const withDate = (day: Date): Date => {
    const now = new Date()
    const hours = date ? date.getHours() : now.getHours()
    const minutes = date ? date.getMinutes() : now.getMinutes() - (now.getMinutes() % MINUTE_STEP)
    return new Date(day.getFullYear(), day.getMonth(), day.getDate(), hours, minutes)
  }

  /**
   * The value (or today) with a new time of day
   */
  const withTime = (hours: number, minutes: number): Date => {
    const day = date ?? new Date()
    return new Date(day.getFullYear(), day.getMonth(), day.getDate(), hours, minutes)
  }

  const change = (next: Date) => onChange(toDateInputValue(next, dateOnly))

  const minutes = Array.from({ length: 60 / MINUTE_STEP }, (_, index) => index * MINUTE_STEP)
  if (date && !minutes.includes(date.getMinutes())) {
    minutes.push(date.getMinutes())
    minutes.sort((a, b) => a - b)
  }

  return (
    <>
      <Input
        root={{ ref: fieldRef }}
        ref={inputRef}
        readOnly
        value={formatDateInputValue(value)}
        placeholder={strings.DatePicker.Placeholder}
        disabled={disabled}
        contentAfter={<CalendarMonthRegular />}
        aria-haspopup='dialog'
        aria-expanded={open}
        title={strings.DatePicker.OpenTitle}
        onClick={() => setOpen(true)}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ' ' || event.key === 'ArrowDown') {
            event.preventDefault()
            setOpen(true)
          }
        }}
      />
      <Popover
        open={open}
        trapFocus
        positioning={{ target: fieldRef.current, position: 'below', align: 'start' }}
        onOpenChange={(event, data) => {
          if (data.open) setOpen(true)
          else if (event.type === 'keydown') close()
          else setOpen(false)
        }}>
        <PopoverSurface className={styles.datePicker} aria-label={strings.DatePicker.OpenTitle}>
          <Calendar
            // Mounted again on the first choice: the calendar cannot go from no value to a value
            key={date ? 'value' : 'empty'}
            value={date}
            firstDayOfWeek={DayOfWeek.Monday}
            strings={getCalendarStrings()}
            showMonthPickerAsOverlay
            highlightSelectedMonth
            showGoToToday
            onSelectDate={(day) => change(withDate(day))}
            // Escape inside the calendar is stopped by the calendar and reported here
            onDismiss={close}
          />
          {!dateOnly && (
            <div className={styles.time} role='group' aria-labelledby={timeLabelId}>
              <Label id={timeLabelId}>{strings.DatePicker.TimeLabel}</Label>
              <Select
                className={styles.timeSelect}
                aria-label={strings.DatePicker.HourLabel}
                value={date ? pad(date.getHours()) : ''}
                onChange={(_, data) =>
                  change(withTime(Number(data.value), date?.getMinutes() ?? 0))
                }>
                {!date && <option value='' />}
                {Array.from({ length: 24 }, (_, hour) => (
                  <option key={hour} value={pad(hour)}>
                    {pad(hour)}
                  </option>
                ))}
              </Select>
              <span aria-hidden='true'>:</span>
              <Select
                className={styles.timeSelect}
                aria-label={strings.DatePicker.MinuteLabel}
                value={date ? pad(date.getMinutes()) : ''}
                onChange={(_, data) => change(withTime(date?.getHours() ?? 0, Number(data.value)))}>
                {!date && <option value='' />}
                {minutes.map((minute) => (
                  <option key={minute} value={pad(minute)}>
                    {pad(minute)}
                  </option>
                ))}
              </Select>
            </div>
          )}
          <Divider />
          <div className={styles.datePickerActions}>
            {!required && (
              <Button appearance='subtle' disabled={!value} onClick={() => onChange('')}>
                {strings.DatePicker.ClearButtonText}
              </Button>
            )}
            <Button appearance='primary' onClick={close}>
              {strings.DatePicker.OkButtonText}
            </Button>
          </div>
        </PopoverSurface>
      </Popover>
    </>
  )
}
