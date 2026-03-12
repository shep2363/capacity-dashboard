export interface ActiveWorkbookStatus {
  hasWorkbook: boolean
  dataset: WorkbookDataset
  fileName: string
  uploadedAt: string | null
  sizeBytes: number
}

export type WorkbookDataset = 'main' | 'sales'

const ACTIVE_WORKBOOK_API_BASE =
  import.meta.env.VITE_SHARED_DATA_API_URL ??
  (typeof window !== 'undefined' && window.location.hostname === 'localhost' ? 'http://127.0.0.1:8000' : '')

function buildApiUrl(path: string): string {
  if (!ACTIVE_WORKBOOK_API_BASE) {
    return path
  }
  return `${ACTIVE_WORKBOOK_API_BASE.replace(/\/$/, '')}${path}`
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

async function fetchJson<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(buildApiUrl(path), { cache: 'no-store', ...init })
  if (!response.ok) {
    throw new Error(await extractErrorMessage(response))
  }
  return (await response.json()) as T
}

export async function fetchActiveWorkbookStatus(dataset: WorkbookDataset): Promise<ActiveWorkbookStatus> {
  return fetchJson<ActiveWorkbookStatus>(`/api/workbook-state?dataset=${dataset}`)
}

export async function downloadActiveWorkbook(dataset: WorkbookDataset): Promise<ArrayBuffer> {
  const response = await fetch(buildApiUrl(`/api/workbook-file?dataset=${dataset}`), { cache: 'no-store' })
  if (!response.ok) {
    throw new Error(await extractErrorMessage(response))
  }
  return response.arrayBuffer()
}

export async function uploadActiveWorkbook(dataset: WorkbookDataset, file: File): Promise<ActiveWorkbookStatus> {
  const formData = new FormData()
  formData.set('file', file)
  return fetchJson<ActiveWorkbookStatus>(`/api/upload-workbook?dataset=${dataset}`, {
    method: 'POST',
    body: formData,
  })
}
