import type { AppFilters } from '../types'

export type WorkbookDataset = 'main' | 'sales'

export interface WorkbookStoreStatus {
  dataset: WorkbookDataset
  hasWorkbook: boolean
  fileName: string
  uploadedAt: string | null
  sizeBytes: number
}

export interface SharedPlanningState {
  version: number
  main: {
    file: string
    overrides: Record<string, number>
    enabled: Record<string, boolean>
    weeklyCaps: Record<string, number>
    filters: AppFilters
    weekendDates: string[]
    weekendExtras: Record<string, number>
  }
  sales: {
    file: string
    overrides: Record<string, number>
    enabled: Record<string, boolean>
  }
}

const DEFAULT_FILTERS: AppFilters = {
  dateFrom: '',
  dateTo: '',
  year: '',
  resources: [],
}

const DEFAULT_SHARED_STATE: SharedPlanningState = {
  version: 1,
  main: {
    file: '',
    overrides: {},
    enabled: {},
    weeklyCaps: {},
    filters: DEFAULT_FILTERS,
    weekendDates: [],
    weekendExtras: {},
  },
  sales: {
    file: '',
    overrides: {},
    enabled: {},
  },
}

const SHARED_API_BASE =
  import.meta.env.VITE_SHARED_DATA_API_URL ??
  (typeof window !== 'undefined' && window.location.hostname === 'localhost' ? 'http://127.0.0.1:8000' : '')

function buildApiUrl(path: string): string {
  if (!SHARED_API_BASE) {
    return path
  }
  return `${SHARED_API_BASE.replace(/\/$/, '')}${path}`
}

async function extractResponseMessage(response: Response): Promise<string> {
  try {
    const payload = await response.json()
    if (payload && typeof payload.error === 'string') {
      return payload.error
    }
    if (payload && typeof payload.message === 'string') {
      return payload.message
    }
  } catch {
    // ignore and try plain text fallback below
  }

  try {
    const text = await response.text()
    if (text) {
      return text
    }
  } catch {
    // ignore
  }

  return `Request failed (${response.status}).`
}

async function fetchJson<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(buildApiUrl(path), {
    cache: 'no-store',
    ...init,
  })

  if (!response.ok) {
    throw new Error(await extractResponseMessage(response))
  }

  return (await response.json()) as T
}

export function createDefaultSharedPlanningState(): SharedPlanningState {
  return {
    version: DEFAULT_SHARED_STATE.version,
    main: {
      file: DEFAULT_SHARED_STATE.main.file,
      overrides: {},
      enabled: {},
      weeklyCaps: {},
      filters: { ...DEFAULT_FILTERS, resources: [] },
      weekendDates: [],
      weekendExtras: {},
    },
    sales: {
      file: DEFAULT_SHARED_STATE.sales.file,
      overrides: {},
      enabled: {},
    },
  }
}

export async function fetchWorkbookStoreStatus(dataset: WorkbookDataset): Promise<WorkbookStoreStatus> {
  return fetchJson<WorkbookStoreStatus>(`/api/workbook-state?dataset=${dataset}`)
}

export async function downloadWorkbookFile(dataset: WorkbookDataset): Promise<ArrayBuffer> {
  const response = await fetch(buildApiUrl(`/api/workbook-file?dataset=${dataset}`), { cache: 'no-store' })
  if (!response.ok) {
    throw new Error(await extractResponseMessage(response))
  }
  return response.arrayBuffer()
}

export async function uploadWorkbookFile(dataset: WorkbookDataset, file: File): Promise<WorkbookStoreStatus> {
  const formData = new FormData()
  formData.set('file', file)

  return fetchJson<WorkbookStoreStatus>(`/api/upload-workbook?dataset=${dataset}`, {
    method: 'POST',
    body: formData,
  })
}

export async function fetchSharedPlanningState(): Promise<SharedPlanningState> {
  const payload = await fetchJson<SharedPlanningState>('/api/shared-state')
  const defaults = createDefaultSharedPlanningState()
  return {
    ...defaults,
    ...payload,
    main: {
      ...defaults.main,
      ...(payload.main ?? {}),
    },
    sales: {
      ...defaults.sales,
      ...(payload.sales ?? {}),
    },
  }
}

export async function saveSharedPlanningState(state: SharedPlanningState): Promise<void> {
  await fetchJson<SharedPlanningState>('/api/shared-state', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(state),
  })
}
