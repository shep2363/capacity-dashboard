import type { WorkbookDataset } from './activeWorkbookApi'

export interface RevenueRateEntry {
  revenuePerHour: number
  grossProfitPerHour: number
}

export type RevenueRateMap = Record<string, RevenueRateEntry>

export interface RevenueRatesPayload {
  dataset: WorkbookDataset
  version: number
  updatedAt: string | null
  source: string
  rateCount: number
  rates: RevenueRateMap
}

export interface SaveRevenueRatesOptions {
  baseVersion?: number | null
  source?: string
}

export class RevenueRatesApiError extends Error {
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

async function parseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    throw new RevenueRatesApiError(response.status, await extractErrorMessage(response))
  }
  return (await response.json()) as T
}

export async function fetchRevenueRates(dataset: WorkbookDataset): Promise<RevenueRatesPayload> {
  const response = await fetch(buildApiUrl(`/api/revenue-rates?dataset=${dataset}`), { cache: 'no-store' })
  return parseJson<RevenueRatesPayload>(response)
}

export async function saveRevenueRates(
  dataset: WorkbookDataset,
  rates: RevenueRateMap,
  options: SaveRevenueRatesOptions = {},
): Promise<RevenueRatesPayload> {
  const payload: Record<string, unknown> = {
    rates,
    source: options.source ?? 'revenue-ui',
  }
  if (typeof options.baseVersion === 'number' && Number.isFinite(options.baseVersion) && options.baseVersion >= 0) {
    payload.baseVersion = options.baseVersion
  }
  const response = await fetch(buildApiUrl(`/api/revenue-rates?dataset=${dataset}`), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    cache: 'no-store',
  })
  return parseJson<RevenueRatesPayload>(response)
}
