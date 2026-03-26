export interface SmartsheetStationProgress {
  process?: number | null
  manualProcess?: number | null
  weld?: number | null
  paint?: number | null
  qc?: number | null
  ship?: number | null
}

export interface SmartsheetProgressEntry {
  rowId: string
  rowNumber: number
  job?: string | null
  sequence?: string | null
  weight?: string | null
  sourceSheet?: string | null
  stationProgress: SmartsheetStationProgress
}

export interface SmartsheetProgressPayload {
  sheetId: string
  updatedAt: string
  rowCount: number
  rows: SmartsheetProgressEntry[]
  matchedColumns?: Record<string, string | null>
}

export interface ResolvedDepartmentProgress {
  rowId: string
  rowNumber: number
  percentComplete: number
  sourceColumn: keyof SmartsheetStationProgress
}

export interface DepartmentProgressMatcher {
  entryCount: number
  resolve: (resource: string, project: string, sequence: string) => ResolvedDepartmentProgress | null
}

function normalizeText(value: string | null | undefined): string {
  return (value ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
}

function canonicalizeResource(value: string | null | undefined): string {
  const normalized = normalizeText(value)
  if (!normalized) return ''
  if (normalized.includes('process')) return 'processing'
  if (normalized.includes('fabric') || normalized === 'fab') return 'fabrication'
  if (normalized.includes('paint')) return 'paint'
  if (normalized.includes('ship')) return 'shipping'
  if (normalized.includes('assembl') || normalized === 'assy') return 'assembly'
  return normalized
}

function uniqueValues(values: Array<string | null | undefined>): string[] {
  const seen = new Set<string>()
  const result: string[] = []
  values.forEach((value) => {
    const normalized = normalizeText(value)
    if (!normalized || seen.has(normalized)) {
      return
    }
    seen.add(normalized)
    result.push(normalized)
  })
  return result
}

function extractDigitBearingTokens(value: string | null | undefined): string[] {
  const normalized = normalizeText(value)
  if (!normalized) {
    return []
  }
  return uniqueValues(
    normalized
      .split(' ')
      .map((token) => token.trim())
      .filter((token) => token.length > 0 && /\d/.test(token)),
  )
}

function entryJobCandidates(entry: SmartsheetProgressEntry): string[] {
  return uniqueValues([entry.job, ...extractDigitBearingTokens(entry.job)])
}

function entrySequenceCandidates(entry: SmartsheetProgressEntry): string[] {
  return uniqueValues([entry.sequence, ...extractDigitBearingTokens(entry.sequence)])
}

function buildUniqueLookup(entries: SmartsheetProgressEntry[], keyFactory: (entry: SmartsheetProgressEntry) => string[]) {
  const unique = new Map<string, SmartsheetProgressEntry>()
  const duplicates = new Set<string>()

  entries.forEach((entry) => {
    keyFactory(entry).forEach((key) => {
      if (!key || duplicates.has(key)) {
        return
      }
      if (unique.has(key)) {
        unique.delete(key)
        duplicates.add(key)
        return
      }
      unique.set(key, entry)
    })
  })

  return unique
}

function toResolvedProgress(
  entry: SmartsheetProgressEntry,
  sourceColumn: keyof SmartsheetStationProgress,
  value: number | null | undefined,
): ResolvedDepartmentProgress | null {
  if (!Number.isFinite(value)) {
    return null
  }
  return {
    rowId: entry.rowId,
    rowNumber: entry.rowNumber,
    percentComplete: Math.min(100, Math.max(0, Number(value))),
    sourceColumn,
  }
}

function resolveStationProgress(resource: string, entry: SmartsheetProgressEntry): ResolvedDepartmentProgress | null {
  const stationProgress = entry.stationProgress ?? {}
  const canonicalResource = canonicalizeResource(resource)

  if (canonicalResource === 'processing') {
    return (
      toResolvedProgress(entry, 'manualProcess', stationProgress.manualProcess) ??
      toResolvedProgress(entry, 'process', stationProgress.process)
    )
  }
  if (canonicalResource === 'fabrication') {
    return toResolvedProgress(entry, 'weld', stationProgress.weld)
  }
  if (canonicalResource === 'paint') {
    return toResolvedProgress(entry, 'paint', stationProgress.paint)
  }
  if (canonicalResource === 'shipping') {
    return toResolvedProgress(entry, 'ship', stationProgress.ship)
  }
  return null
}

export function buildDepartmentProgressMatcher(entries: SmartsheetProgressEntry[]): DepartmentProgressMatcher {
  const jobSequence = buildUniqueLookup(entries, (entry) => {
    const jobs = entryJobCandidates(entry)
    const sequences = entrySequenceCandidates(entry)
    const keys: string[] = []

    jobs.forEach((job) => {
      sequences.forEach((sequence) => {
        keys.push(`${job}|${sequence}`)
      })
    })

    return keys
  })

  return {
    entryCount: entries.length,
    resolve(resource: string, project: string, sequence: string) {
      const jobCandidates = uniqueValues([project, ...extractDigitBearingTokens(project)])
      const sequenceCandidates = uniqueValues([sequence, ...extractDigitBearingTokens(sequence)])

      if (jobCandidates.length === 0 || sequenceCandidates.length === 0) {
        return null
      }

      for (const job of jobCandidates) {
        for (const sequenceValue of sequenceCandidates) {
          const matched = jobSequence.get(`${job}|${sequenceValue}`)
          if (!matched) {
            continue
          }
          const resolved = resolveStationProgress(resource, matched)
          if (resolved) {
            return resolved
          }
        }
      }

      return null
    },
  }
}

export function formatSmartsheetSyncLabel(
  status: 'idle' | 'loading' | 'loaded' | 'error',
  updatedAt: string | null,
  errorMessage: string,
  matchedCount?: number,
  totalCount?: number,
): string {
  if (status === 'loading') {
    return 'Syncing Smartsheet...'
  }
  if (status === 'error') {
    return errorMessage || 'Smartsheet sync failed'
  }
  const matchSummary =
    typeof matchedCount === 'number' && typeof totalCount === 'number'
      ? ` - ${matchedCount}/${totalCount} matched`
      : ''
  if (status === 'loaded' && updatedAt) {
    return `Smartsheet synced ${new Date(updatedAt).toLocaleString()}${matchSummary}`
  }
  return `Smartsheet not synced${matchSummary}`
}
