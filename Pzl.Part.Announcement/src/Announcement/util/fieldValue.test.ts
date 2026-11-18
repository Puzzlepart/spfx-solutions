import {
  dateInputValueToIso,
  htmlToText,
  formatDateInputValue,
  parseDateInputValue,
  toDateInputValue
} from './fieldValue'

describe('date form values', () => {
  it('round-trips a date and time in local time', () => {
    const date = new Date(2026, 10, 17, 9, 5)
    expect(toDateInputValue(date, false)).toBe('2026-11-17T09:05')
    expect(parseDateInputValue('2026-11-17T09:05').getTime()).toBe(date.getTime())
  })

  it('reads a date-only value as local midnight', () => {
    expect(toDateInputValue(new Date(2026, 0, 3, 14, 30), true)).toBe('2026-01-03')
    expect(parseDateInputValue('2026-01-03').getTime()).toBe(new Date(2026, 0, 3).getTime())
  })

  it('formats for display, with the time only when the value has one', () => {
    expect(formatDateInputValue('2026-11-17T09:05')).toBe('17.11.2026 kl. 09:05')
    expect(formatDateInputValue('2026-01-03')).toBe('03.01.2026')
  })

  it('treats empty and invalid values as no value', () => {
    expect(parseDateInputValue('')).toBeUndefined()
    expect(parseDateInputValue('ikke en dato')).toBeUndefined()
    expect(formatDateInputValue('')).toBe('')
    expect(dateInputValueToIso('')).toBeNull()
    expect(dateInputValueToIso('2026-11-17T09:05')).toBe(new Date(2026, 10, 17, 9, 5).toISOString())
  })
})

describe('htmlToText', () => {
  it('removes tags and decodes the character references SharePoint stores in rich text', () => {
    expect(
      htmlToText(
        '<div class="ExternalClass"><p>E-post&#58; nede &amp; treg&#160;&quot;i dag&quot;&#8203;</p><p>Bl&aring;b&#xE6;r &lt;ikke en tagg&gt;</p></div>'
      )
    ).toBe('E-post: nede & treg "i dag" Blåbær <ikke en tagg>')
  })

  it('leaves plain text alone and treats markup without text as empty', () => {
    expect(htmlToText('Tom & Jerry 100 %')).toBe('Tom & Jerry 100 %')
    expect(htmlToText('<p>&nbsp;</p><br>')).toBe('')
    expect(htmlToText(null)).toBe('')
  })
})
