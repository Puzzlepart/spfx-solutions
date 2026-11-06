import {
  compareAnnouncements,
  isCriticalValue,
  Severity,
  Status,
  statusChoiceValue,
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
