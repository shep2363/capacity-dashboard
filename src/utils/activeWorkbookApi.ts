import { buildSharedApiUrl, withAuth } from './apiBase'

export interface ActiveWorkbookStatus {
  hasWorkbook: boolean
  dataset: WorkbookDataset
  fileName: string
  uploadedAt: string | null
  sizeBytes: number
}

export type WorkbookDataset = 'main' | 'sales'

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
  const response = await fetch(buildSharedApiUrl(path), withAuth(init))
  if (!response.ok) {
    throw new Error(await extractErrorMessage(response))
  }
  return (await response.json()) as T
}

export async function fetchActiveWorkbookStatus(dataset: WorkbookDataset): Promise<ActiveWorkbookStatus> {
  return fetchJson<ActiveWorkbookStatus>(`/api/workbook-state?dataset=${dataset}`)
}

export async function downloadActiveWorkbook(dataset: WorkbookDataset): Promise<ArrayBuffer> {
  const response = await fetch(buildSharedApiUrl(`/api/workbook-file?dataset=${dataset}`), withAuth())
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
