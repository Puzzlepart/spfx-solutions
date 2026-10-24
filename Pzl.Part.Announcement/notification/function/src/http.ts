/**
 * Minimal fetch wrapper with retry on throttling (429) and transient errors (502/503/504).
 */
export class HttpError extends Error {
  constructor(
    public readonly status: number,
    public readonly url: string,
    public readonly body: string
  ) {
    super(`HTTP ${status} for ${url}: ${body.substring(0, 500)}`)
  }
}

export interface RequestOptions {
  method?: string
  headers?: Record<string, string>
  body?: unknown
  /** Number of retries on 429/5xx. Default 4 */
  retries?: number
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

export async function request<T = unknown>(url: string, options: RequestOptions = {}): Promise<T> {
  const retries = options.retries ?? 4
  let attempt = 0
  for (;;) {
    const response = await fetch(url, {
      method: options.method ?? 'GET',
      headers: {
        Accept: 'application/json',
        ...(options.body !== undefined ? { 'Content-Type': 'application/json' } : {}),
        ...options.headers
      },
      body: options.body !== undefined ? JSON.stringify(options.body) : undefined
    })

    if (response.ok) {
      if (response.status === 204) return undefined as T
      const text = await response.text()
      return (text ? JSON.parse(text) : undefined) as T
    }

    const retryable = response.status === 429 || [502, 503, 504].includes(response.status)
    if (retryable && attempt < retries) {
      const retryAfter = Number(response.headers.get('Retry-After'))
      const delay = retryAfter > 0 ? retryAfter * 1000 : Math.min(2 ** attempt * 1000, 30_000)
      attempt++
      await sleep(delay)
      continue
    }

    throw new HttpError(response.status, url, await response.text())
  }
}
