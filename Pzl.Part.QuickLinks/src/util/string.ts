export const isNullOrEmpty = (value?: string | null): boolean => {
  return value === undefined || value === null || value.length === 0
}
