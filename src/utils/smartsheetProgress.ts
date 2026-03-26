export interface SmartsheetProgressIdentifiers {
  sir?: string | null
  quote?: string | null
  title?: string | null
  jobNumber?: string | null
}

export interface SmartsheetProgressEntry {
  rowId: string
  rowNumber: number
  percentComplete: number
  project?: string | null
  sequence?: string | null
  resource?: string | null
  identifiers?: SmartsheetProgressIdentifiers
}

export interface SmartsheetProgressPayload {
  sheetId: string
  updatedAt: string
  rowCount: number
  rows: SmartsheetProgressEntry[]
  matchedColumns?: Record<string, string | null>
}

export interface DepartmentProgressMatcher {
  entryCount: number
  resolve: (resource: string, project: string, sequence: string) => SmartsheetProgressEntry | null
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

function entryProjectCandidates(entry: SmartsheetProgressEntry): string[] {
  const composedProject = [
    entry.identifiers?.sir,
    entry.identifiers?.quote,
    entry.identifiers?.title,
  ]
    .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
    .join(' - ')

  return uniqueValues([entry.project, composedProject, entry.identifiers?.jobNumber])
}

function entrySequenceCandidates(entry: SmartsheetProgressEntry): string[] {
  return uniqueValues([entry.sequence])
}

function entryResourceCandidates(entry: SmartsheetProgressEntry): string[] {
  const canonical = canonicalizeResource(entry.resource)
  return canonical ? [canonical] : []
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

export function buildDepartmentProgressMatcher(entries: SmartsheetProgressEntry[]): DepartmentProgressMatcher {
  const resourceProjectSequence = buildUniqueLookup(entries, (entry) => {
    const resources = entryResourceCandidates(entry)
    const projects = entryProjectCandidates(entry)
    const sequences = entrySequenceCandidates(entry)
    const keys: string[] = []
    resources.forEach((resource) => {
      projects.forEach((project) => {
        sequences.forEach((sequence) => {
          keys.push(`${resource}|${project}|${sequence}`)
        })
      })
    })
    return keys
  })

  const projectSequence = buildUniqueLookup(entries, (entry) => {
    const projects = entryProjectCandidates(entry)
    const sequences = entrySequenceCandidates(entry)
    const keys: string[] = []
    projects.forEach((project) => {
      sequences.forEach((sequence) => {
        keys.push(`${project}|${sequence}`)
      })
    })
    return keys
  })

  const resourceSequence = buildUniqueLookup(entries, (entry) => {
    const resources = entryResourceCandidates(entry)
    const sequences = entrySequenceCandidates(entry)
    const keys: string[] = []
    resources.forEach((resource) => {
      sequences.forEach((sequence) => {
        keys.push(`${resource}|${sequence}`)
      })
    })
    return keys
  })

  const sequenceOnly = buildUniqueLookup(entries, (entry) => entrySequenceCandidates(entry))

  return {
    entryCount: entries.length,
    resolve(resource: string, project: string, sequence: string) {
      const canonicalResource = canonicalizeResource(resource)
      const normalizedProject = normalizeText(project)
      const normalizedSequence = normalizeText(sequence)

      if (!normalizedSequence) {
        return null
      }

      if (canonicalResource && normalizedProject) {
        const exact = resourceProjectSequence.get(`${canonicalResource}|${normalizedProject}|${normalizedSequence}`)
        if (exact) return exact
      }

      if (normalizedProject) {
        const byProject = projectSequence.get(`${normalizedProject}|${normalizedSequence}`)
        if (byProject) return byProject
      }

      if (canonicalResource) {
        const byResource = resourceSequence.get(`${canonicalResource}|${normalizedSequence}`)
        if (byResource) return byResource
      }

      return sequenceOnly.get(normalizedSequence) ?? null
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
      ? ` • ${matchedCount}/${totalCount} matched`
      : ''
  if (status === 'loaded' && updatedAt) {
    return `Smartsheet synced ${new Date(updatedAt).toLocaleString()}${matchSummary}`
  }
  return `Smartsheet not synced${matchSummary}`
}
