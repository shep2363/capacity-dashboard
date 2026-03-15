import type { WorkbookDataset } from './activeWorkbookApi'
import { buildSharedApiUrl, withAuth } from './apiBase'

export interface PlanningStatePayload {
  dataset: WorkbookDataset
  version: number
  updatedAt: string | null
  source: string
  overrideCount: number
  overrides: Record<string, number>
}

export interface SavePlanningStateOptions {
  baseVersion?: number | null
  source?: string
}

export class PlanningStateApiError extends Error {
  status: number

  constructor(status: number, message: string) {
    super(message)
    this.status = status
  }
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
    throw new PlanningStateApiError(response.status, await extractErrorMessage(response))
  }
  return (await response.json()) as T
}

export async function fetchPlanningState(dataset: WorkbookDataset): Promise<PlanningStatePayload> {
  const response = await fetch(buildSharedApiUrl(`/api/planning-state?dataset=${dataset}`), withAuth())
  return parseJson<PlanningStatePayload>(response)
}

export async function savePlanningState(
  dataset: WorkbookDataset,
  overrides: Record<string, number>,
  options: SavePlanningStateOptions = {},
): Promise<PlanningStatePayload> {
  const payload: Record<string, unknown> = {
    overrides,
    source: options.source ?? 'planning-ui',
  }
  if (typeof options.baseVersion === 'number' && Number.isFinite(options.baseVersion) && options.baseVersion >= 0) {
    payload.baseVersion = options.baseVersion
  }
  const response = await fetch(
    buildSharedApiUrl(`/api/planning-state?dataset=${dataset}`),
    withAuth({
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    }),
  )
  return parseJson<PlanningStatePayload>(response)
}
