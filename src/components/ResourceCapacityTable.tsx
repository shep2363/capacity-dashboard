import { useEffect, useMemo, useState } from 'react'
import type { Dispatch, SetStateAction } from 'react'

interface ResourceCapacityTableProps {
  resources: string[]
  enabledResources: Record<string, boolean>
  weeklyCapacitiesByResource: Record<string, number>
  onWeeklyCapacityChange: (resource: string, weeklyCapacity: number) => void
  onToggleResource: (resource: string, enabled: boolean) => void
  weekendExtraByResource: Record<string, number>
  onWeekendExtraChange: (resource: string, weekendHours: number) => void
}

const WEEKS_PER_MONTH = 52 / 12

function monthlyFromWeekly(weekly: number): number {
  return weekly * WEEKS_PER_MONTH
}

export function ResourceCapacityTable({
  resources,
  enabledResources,
  weeklyCapacitiesByResource,
  onWeeklyCapacityChange,
  onToggleResource,
  weekendExtraByResource,
  onWeekendExtraChange,
}: ResourceCapacityTableProps) {
  const [isCollapsed, setIsCollapsed] = useState(false)
  const [draftWeeklyByResource, setDraftWeeklyByResource] = useState<Record<string, string>>({})
  const [draftWeekendByResource, setDraftWeekendByResource] = useState<Record<string, string>>({})

  const resourceSet = useMemo(() => new Set(resources), [resources])

  useEffect(() => {
    const next: Record<string, string> = {}
    for (const resource of resources) {
      const weekly = weeklyCapacitiesByResource[resource] ?? 0
      next[resource] = Number.isFinite(weekly) ? String(weekly) : '0'
    }
    setDraftWeeklyByResource(next)
  }, [resources, weeklyCapacitiesByResource])

  useEffect(() => {
    const next: Record<string, string> = {}
    for (const resource of resources) {
      const weekend = weekendExtraByResource[resource] ?? 0
      // Keep weekend contribution visually blank when value is zero.
      if (!Number.isFinite(weekend) || weekend === 0) {
        next[resource] = ''
      } else {
        next[resource] = String(weekend)
      }
    }
    setDraftWeekendByResource(next)
  }, [resources, weekendExtraByResource])

  function parseCommittedHours(draft: string): number {
    if (draft.trim() === '') {
      return 0
    }
    const parsed = Number(draft)
    if (!Number.isFinite(parsed)) {
      return 0
    }
    return Math.max(0, parsed)
  }

  function updateDraftValue(
    setter: Dispatch<SetStateAction<Record<string, string>>>,
    resource: string,
    rawValue: string,
  ): void {
    // Allow empty string while editing; otherwise keep numeric-ish values only.
    if (rawValue !== '' && !/^\d*\.?\d*$/.test(rawValue)) {
      return
    }

    setter((current) => ({
      ...current,
      [resource]: rawValue,
    }))
  }

  function commitWeeklyDraft(resource: string): void {
    const draft = draftWeeklyByResource[resource] ?? ''
    const committed = parseCommittedHours(draft)
    onWeeklyCapacityChange(resource, committed)
    setDraftWeeklyByResource((current) => ({
      ...current,
      [resource]: committed === 0 ? (draft.trim() === '' ? '' : '0') : String(committed),
    }))
  }

  function commitWeekendDraft(resource: string): void {
    const draft = draftWeekendByResource[resource] ?? ''
    const committed = parseCommittedHours(draft)
    onWeekendExtraChange(resource, committed)
    setDraftWeekendByResource((current) => ({
      ...current,
      [resource]: committed === 0 ? '' : String(committed),
    }))
  }

  function resetWeeklyDraft(resource: string): void {
    if (!resourceSet.has(resource)) {
      return
    }
    const value = weeklyCapacitiesByResource[resource] ?? 0
    setDraftWeeklyByResource((current) => ({ ...current, [resource]: String(Math.max(0, value)) }))
  }

  function resetWeekendDraft(resource: string): void {
    if (!resourceSet.has(resource)) {
      return
    }
    const value = weekendExtraByResource[resource] ?? 0
    const normalized = Math.max(0, value)
    setDraftWeekendByResource((current) => ({
      ...current,
      [resource]: normalized === 0 ? '' : String(normalized),
    }))
  }

  return (
    <section className="panel table-panel resource-panel">
      <div className="section-header section-header-row">
        <div>
          <h2>Resource Capacity Input</h2>
          <p>
            Enter weekly capacity by resource. Weekend contribution is based on selected working weekend days in the
            planning period.
          </p>
        </div>
        <div className="section-actions">
          <button type="button" className="ghost-btn" onClick={() => setIsCollapsed((current) => !current)}>
            {isCollapsed ? 'Expand Table' : 'Collapse Table'}
          </button>
        </div>
      </div>

      {!isCollapsed && (
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Enabled</th>
                <th>Resource Name</th>
                <th>Weekly Capacity Hours</th>
                <th>Weekend Capacity Contribution</th>
                <th>Effective Weekly Capacity</th>
                <th>Monthly Capacity Hours</th>
              </tr>
            </thead>
            <tbody>
              {resources.map((resource) => {
                const weekly = weeklyCapacitiesByResource[resource] ?? 0
                const manualWeekend = weekendExtraByResource[resource] ?? 0
                const effectiveWeekly = weekly + manualWeekend
                const monthly = monthlyFromWeekly(effectiveWeekly)
                const isEnabled = enabledResources[resource] !== false

                return (
                  <tr key={resource} className={isEnabled ? 'capacity-selected-row' : ''}>
                    <td>
                      <label className="toggle-switch">
                        <input
                          type="checkbox"
                          checked={isEnabled}
                          onChange={(event) => onToggleResource(resource, event.target.checked)}
                          aria-label={`Toggle ${resource}`}
                        />
                        <span className="toggle-slider" />
                      </label>
                    </td>
                    <td>{resource}</td>
                    <td>
                      <input
                        className="capacity-cell"
                        type="text"
                        inputMode="decimal"
                        value={draftWeeklyByResource[resource] ?? String(Number.isFinite(weekly) ? weekly : 0)}
                        onChange={(event) => updateDraftValue(setDraftWeeklyByResource, resource, event.target.value)}
                        onBlur={() => commitWeeklyDraft(resource)}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') {
                            event.preventDefault()
                            commitWeeklyDraft(resource)
                            ;(event.currentTarget as HTMLInputElement).blur()
                          }
                          if (event.key === 'Escape') {
                            event.preventDefault()
                            resetWeeklyDraft(resource)
                            ;(event.currentTarget as HTMLInputElement).blur()
                          }
                        }}
                        aria-label={`Weekly capacity for ${resource}`}
                      />
                    </td>
                    <td>
                      <input
                        className="capacity-cell"
                        type="text"
                        inputMode="decimal"
                        value={draftWeekendByResource[resource] ?? String(Number.isFinite(manualWeekend) ? manualWeekend : 0)}
                        onChange={(event) => updateDraftValue(setDraftWeekendByResource, resource, event.target.value)}
                        onBlur={() => commitWeekendDraft(resource)}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter') {
                            event.preventDefault()
                            commitWeekendDraft(resource)
                            ;(event.currentTarget as HTMLInputElement).blur()
                          }
                          if (event.key === 'Escape') {
                            event.preventDefault()
                            resetWeekendDraft(resource)
                            ;(event.currentTarget as HTMLInputElement).blur()
                          }
                        }}
                        aria-label={`Weekend capacity for ${resource}`}
                      />
                    </td>
                    <td>{effectiveWeekly.toFixed(2)}</td>
                    <td>{monthly.toFixed(2)}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}
    </section>
  )
}
