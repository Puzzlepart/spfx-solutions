import {
  compareAnnouncements,
  compareWithPeriods,
  isAlwaysNotified,
  isCriticalValue,
  Period,
  periodFilter,
  Severity,
  Status,
  statusChoiceValue,
  toKnownSeverity,
  toPeriod,
  toSeverity,
  toStatus
} from './model'

describe('toSeverity', () => {
  it('maps the 2.0 values in Norwegian and English', () => {
    expect(toSeverity('Informasjon')).toBe('info')
    expect(toSeverity('Advarsel')).toBe('warning')
    expect(toSeverity('Kritisk')).toBe('critical')
    expect(toSeverity('Critical')).toBe('critical')
  })

  it('maps the 1.x values', () => {
    expect(toSeverity('Varsel')).toBe('critical')
    expect(toSeverity('Alert')).toBe('critical')
    expect(toSeverity('Normal')).toBe('info')
  })

  it('ignores case and whitespace, and shows unknown values as information', () => {
    expect(toSeverity(' kritisk ')).toBe('critical')
    expect(toSeverity('Ukjent')).toBe('info')
    expect(toSeverity(null)).toBe('info')
  })
})

describe('toStatus', () => {
  it('maps the values in Norwegian and English', () => {
    expect(toStatus('Planlagt')).toBe('planned')
    expect(toStatus('Pågår')).toBe('ongoing')
    expect(toStatus('Løst')).toBe('resolved')
    expect(toStatus('In progress')).toBe('ongoing')
  })

  it('treats empty and unknown values as ongoing', () => {
    expect(toStatus(undefined)).toBe('ongoing')
    expect(toStatus('')).toBe('ongoing')
    expect(toStatus('Noe annet')).toBe('ongoing')
  })
})

describe('isCriticalValue', () => {
  it('is true for critical values only', () => {
    expect(isCriticalValue('Kritisk')).toBe(true)
    expect(isCriticalValue('Varsel')).toBe(true)
    expect(isCriticalValue('Advarsel')).toBe(false)
    expect(isCriticalValue('')).toBe(false)
  })
})

describe('statusChoiceValue', () => {
  it('uses the choice of the list when it has one', () => {
    expect(statusChoiceValue('resolved', ['Planned', 'In progress', 'Resolved'])).toBe('Resolved')
  })

  it('falls back to the template value', () => {
    expect(statusChoiceValue('resolved')).toBe('Løst')
    expect(statusChoiceValue('ongoing', [])).toBe('Pågår')
  })
})

describe('compareAnnouncements', () => {
  const item = (
    id: string,
    status: Status,
    severity: Severity,
    startDate: string,
    pinned = false
  ) => ({ id, status, severity, startDate, pinned })

  it('sorts pinned, then status, then level, then newest start date', () => {
    const items = [
      item('resolved-critical', 'resolved', 'critical', '2026-11-05'),
      item('planned-info', 'planned', 'info', '2026-11-05'),
      item('ongoing-info-old', 'ongoing', 'info', '2026-11-01'),
      item('ongoing-info-new', 'ongoing', 'info', '2026-11-04'),
      item('ongoing-critical', 'ongoing', 'critical', '2026-11-01'),
      item('pinned-resolved', 'resolved', 'info', '2026-10-01', true)
    ]
    expect(items.sort(compareAnnouncements).map((a) => a.id)).toEqual([
      'pinned-resolved',
      'ongoing-critical',
      'ongoing-info-new',
      'ongoing-info-old',
      'planned-info',
      'resolved-critical'
    ])
  })
})

describe('isAlwaysNotified', () => {
  it('is true for a level in «Nivåer som alltid varsles», in any language', () => {
    expect(isAlwaysNotified('Kritisk', ['critical'])).toBe(true)
    expect(isAlwaysNotified('Critical', ['critical', 'warning'])).toBe(true)
    expect(isAlwaysNotified('Varsel', ['critical'])).toBe(true)
  })

  it('is false for other levels, no level and no configuration', () => {
    expect(isAlwaysNotified('Advarsel', ['critical'])).toBe(false)
    expect(isAlwaysNotified('', ['info'])).toBe(false)
    expect(isAlwaysNotified(null, ['info'])).toBe(false)
    expect(isAlwaysNotified('Kritisk', [])).toBe(false)
    expect(isAlwaysNotified('Kritisk', undefined)).toBe(false)
  })
})

describe('toKnownSeverity', () => {
  it('maps known values and leaves out free text', () => {
    expect(toKnownSeverity('Kritisk')).toBe('critical')
    expect(toKnownSeverity('Varsel')).toBe('critical')
    expect(toKnownSeverity('Alle')).toBeNull()
    expect(toKnownSeverity(null)).toBeNull()
  })
})

describe('toPeriod', () => {
  const now = new Date('2026-11-18T12:00:00Z')

  it('tells upcoming, active and past apart on the display period', () => {
    expect(toPeriod('2026-11-19T00:00:00Z', '2026-11-20T00:00:00Z', now)).toBe('upcoming')
    expect(toPeriod('2026-11-17T00:00:00Z', '2026-11-20T00:00:00Z', now)).toBe('active')
    expect(toPeriod('2026-11-10T00:00:00Z', '2026-11-17T00:00:00Z', now)).toBe('past')
  })

  it('counts an announcement without dates as active', () => {
    expect(toPeriod(null, undefined, now)).toBe('active')
  })
})

describe('compareWithPeriods', () => {
  const item = (id: string, period: Period, startDate: string, endDate: string) => ({
    id,
    period,
    startDate,
    endDate,
    pinned: false,
    status: 'ongoing' as Status,
    severity: 'info' as Severity
  })

  it('lists active first, then upcoming by start, then past by latest end', () => {
    const sorted = [
      item('past-old', 'past', '2026-01-01', '2026-01-02'),
      item('upcoming-late', 'upcoming', '2026-12-10', '2026-12-11'),
      item('active', 'active', '2026-11-17', '2026-11-20'),
      item('past-new', 'past', '2026-11-01', '2026-11-02'),
      item('upcoming-soon', 'upcoming', '2026-11-25', '2026-11-26')
    ].sort(compareWithPeriods)
    expect(sorted.map((a) => a.id)).toEqual([
      'active',
      'upcoming-soon',
      'upcoming-late',
      'past-new',
      'past-old'
    ])
  })
})

describe('periodFilter', () => {
  const now = '2026-11-18T12:00:00.000Z'
  const started = `PzlStartDate le datetime'${now}'`
  const notEnded = `PzlEndDate ge datetime'${now}'`

  it('fetches only active announcements by default', () => {
    expect(periodFilter(now)).toBe(`${started} and ${notEnded}`)
    expect(periodFilter(now, {})).toBe(`${started} and ${notEnded}`)
  })

  it('drops the end date limit for past and the start date limit for upcoming', () => {
    expect(periodFilter(now, { past: true })).toBe(started)
    expect(periodFilter(now, { upcoming: true })).toBe(notEnded)
    expect(periodFilter(now, { past: true, upcoming: true })).toBe('')
  })
})
