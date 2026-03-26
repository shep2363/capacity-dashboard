import type { SmartsheetProgressPayload } from './smartsheetProgress'

export class SmartsheetProgressApiError extends Error {
  status: number

  constructor(status: number, message: string) {
    super(message)
    this.status = status
  }
}

const API_BASE =
  import.meta.env.VITE_SHARED_DATA_API_URL ??
  (typeof window !== 'undefined' && window.location.hostname === 'localhost' ? 'http://127.0.0.1:8000' : '')

function buildApiUrl(path: string): string {
  if (!API_BASE) {
    return path
  }
  return `${API_BASE.replace(/\/$/, '')}${path}`
}

async function extractErrorMessage(response: Response): Promise<string> {
  try {
    const payload = await response.json()
    if (payload && typeof payload.error === 'string') {
      return payload.error
    }
    if (payload && typeof payload.message === 'string') {
      return payload.message
    }
  } catch {
    // no-op
  }
  return `Request failed (${response.status}).`
}

export async function fetchSmartsheetProgress(): Promise<SmartsheetProgressPayload> {
  const response = await fetch(buildApiUrl('/api/smartsheet-progress'), { cache: 'no-store' })
  if (!response.ok) {
    throw new SmartsheetProgressApiError(response.status, await extractErrorMessage(response))
  }
  return (await response.json()) as SmartsheetProgressPayload
}
