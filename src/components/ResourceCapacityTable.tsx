import { useEffect, useMemo, useState } from 'react'
import type { Dispatch, SetStateAction } from 'react'
import { weekRangeLabel } from '../utils/planner'

interface ResourceCapacityTableProps {
  resources: string[]
  enabledResources: Record<string, boolean>
  weeklyCapacitiesByResource: Record<string, number>
  weekKeys: string[]
  weeklyCapacityOverridesByResource: Record<string, Record<string, number>>
  onWeeklyCapacityChange: (resource: string, weeklyCapacity: number) => void
  onSetWeeklyCapacityOverride: (resource: string, weekStartIso: string, weeklyCapacity: number) => void
  onSetWeeklyCapacityOverrides: (resource: string, weekStartIsos: string[], weeklyCapacity: number) => void
  onClearWeeklyCapacityOverride: (resource: string, weekStartIso: string) => void
  onClearWeeklyCapacityOverrides: (resource: string, weekStartIsos: string[]) => void
  onClearAllWeeklyCapacityOverrides: () => void
  onToggleResource: (resource: string, enabled: boolean) => void
  weekendExtraByResource: Record<string, number>
  onWeekendExtraChange: (resource: string, weekendHours: number) => void
  // holidayWeeks is unused here but kept for future per-week breakdowns
  holidayWeeks?: Map<string, number>
}

const WEEKS_PER_MONTH = 52 / 12

function monthlyFromWeekly(weekly: number): number {
  return weekly * WEEKS_PER_MONTH
}

export function ResourceCapacityTable({
  resources,
  enabledResources,
  weeklyCapacitiesByResource,
  weekKeys,
  weeklyCapacityOverridesByResource,
  onWeeklyCapacityChange,
  onSetWeeklyCapacityOverride,
  onSetWeeklyCapacityOverrides,
  onClearWeeklyCapacityOverride,
  onClearWeeklyCapacityOverrides,
  onClearAllWeeklyCapacityOverrides,
  onToggleResource,
  weekendExtraByResource,
  onWeekendExtraChange,
}: ResourceCapacityTableProps) {
  const [isCollapsed, setIsCollapsed] = useState(true)
  const [draftWeeklyByResource, setDraftWeeklyByResource] = useState<Record<string, string>>({})
  const [draftWeekendByResource, setDraftWeekendByResource] = useState<Record<string, string>>({})
  const [selectedOverrideResource, setSelectedOverrideResource] = useState('')
  const [selectedOverrideWeek, setSelectedOverrideWeek] = useState('')
  const [selectedOverrideWeeks, setSelectedOverrideWeeks] = useState<string[]>([])
  const [draftWeekOverrideHours, setDraftWeekOverrideHours] = useState('')
  const [bulkValidationMessage, setBulkValidationMessage] = useState('')

  const resourceSet = useMemo(() => new Set(resources), [resources])
  const weekSet = useMemo(() => new Set(weekKeys), [weekKeys])

  const weeklyOverrideRows = useMemo(
    () =>
      Object.entries(weeklyCapacityOverridesByResource)
        .flatMap(([resource, weekMap]) =>
          Object.entries(weekMap).map(([weekIso, hours]) => ({
            resource,
            weekIso,
            hours,
            defaultWeekly: weeklyCapacitiesByResource[resource] ?? 0,
          })),
        )
        .filter((row) => Number.isFinite(row.hours))
        .sort((a, b) => (a.weekIso === b.weekIso ? a.resource.localeCompare(b.resource) : a.weekIso.localeCompare(b.weekIso))),
    [weeklyCapacityOverridesByResource, weeklyCapacitiesByResource],
  )

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
      if (!Number.isFinite(weekend) || weekend === 0) {
        next[resource] = ''
      } else {
        next[resource] = String(weekend)
      }
    }
    setDraftWeekendByResource(next)
  }, [resources, weekendExtraByResource])

  useEffect(() => {
    if (resources.length === 0) {
      setSelectedOverrideResource('')
      return
    }
    setSelectedOverrideResource((current) =>
      current && resourceSet.has(current) ? current : resources[0],
    )
  }, [resources, resourceSet])

  useEffect(() => {
    if (weekKeys.length === 0) {
      setSelectedOverrideWeek('')
      setSelectedOverrideWeeks([])
      return
    }
    setSelectedOverrideWeek((current) => (current && weekSet.has(current) ? current : weekKeys[0]))
    setSelectedOverrideWeeks((current) => {
      if (current.length === 0) {
        return [weekKeys[0]]
      }
      const filtered = current.filter((weekIso) => weekSet.has(weekIso))
      return filtered.length > 0 ? filtered : [weekKeys[0]]
    })
  }, [weekKeys, weekSet])

  useEffect(() => {
    if (!selectedOverrideResource || !selectedOverrideWeek) {
      setDraftWeekOverrideHours('')
      return
    }
    const override = weeklyCapacityOverridesByResource[selectedOverrideResource]?.[selectedOverrideWeek]
    if (Number.isFinite(override)) {
      setDraftWeekOverrideHours(String(override))
      return
    }
    setDraftWeekOverrideHours('')
  }, [selectedOverrideResource, selectedOverrideWeek, weeklyCapacityOverridesByResource])

  useEffect(() => {
    setBulkValidationMessage('')
  }, [selectedOverrideResource, selectedOverrideWeek, selectedOverrideWeeks])

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

  function updateWeekOverrideDraft(rawValue: string): void {
    if (rawValue !== '' && !/^\d*\.?\d*$/.test(rawValue)) {
      return
    }
    setDraftWeekOverrideHours(rawValue)
  }

  function applyWeekOverride(): void {
    if (!selectedOverrideResource || !selectedOverrideWeek) {
      return
    }
    if (!weekSet.has(selectedOverrideWeek)) {
      return
    }
    if (draftWeekOverrideHours.trim() === '') {
      return
    }
    const committed = parseCommittedHours(draftWeekOverrideHours)
    onSetWeeklyCapacityOverride(selectedOverrideResource, selectedOverrideWeek, committed)
    setDraftWeekOverrideHours(String(committed))
    setBulkValidationMessage('')
  }

  function clearSelectedWeekOverride(): void {
    if (!selectedOverrideResource || !selectedOverrideWeek) {
      return
    }
    onClearWeeklyCapacityOverride(selectedOverrideResource, selectedOverrideWeek)
    setDraftWeekOverrideHours('')
    setBulkValidationMessage('')
  }

  function toggleSelectedWeek(weekIso: string): void {
    setSelectedOverrideWeeks((current) => {
      if (current.includes(weekIso)) {
        return current.filter((value) => value !== weekIso)
      }
      return [...current, weekIso].sort((a, b) => a.localeCompare(b))
    })
  }

  function selectAllWeeks(): void {
    setSelectedOverrideWeeks([...weekKeys])
    setBulkValidationMessage('')
  }

  function clearWeekSelection(): void {
    setSelectedOverrideWeeks([])
    setBulkValidationMessage('')
  }

  function applyOverrideToSelectedWeeks(): void {
    if (!selectedOverrideResource) {
      setBulkValidationMessage('Select a resource before applying override hours.')
      return
    }
    if (selectedOverrideWeeks.length === 0) {
      setBulkValidationMessage('Select at least one week for bulk override.')
      return
    }
    if (draftWeekOverrideHours.trim() === '') {
      setBulkValidationMessage('Enter override capacity hours before bulk apply.')
      return
    }

    const committed = parseCommittedHours(draftWeekOverrideHours)
    onSetWeeklyCapacityOverrides(selectedOverrideResource, selectedOverrideWeeks, committed)
    setDraftWeekOverrideHours(String(committed))
    setBulkValidationMessage('')
  }

  function clearSelectedWeeksOverrides(): void {
    if (!selectedOverrideResource) {
      setBulkValidationMessage('Select a resource before clearing overrides.')
      return
    }
    if (selectedOverrideWeeks.length === 0) {
      setBulkValidationMessage('Select at least one week to clear overrides.')
      return
    }
    onClearWeeklyCapacityOverrides(selectedOverrideResource, selectedOverrideWeeks)
    if (selectedOverrideWeeks.includes(selectedOverrideWeek)) {
      setDraftWeekOverrideHours('')
    }
    setBulkValidationMessage('')
  }

  function clearAllOverrides(): void {
    if (weeklyOverrideRows.length === 0) {
      return
    }
    const confirmed = window.confirm('Clear all weekly capacity overrides?')
    if (!confirmed) {
      return
    }
    onClearAllWeeklyCapacityOverrides()
    setDraftWeekOverrideHours('')
    setBulkValidationMessage('')
  }

  const selectedDefaultWeekly = selectedOverrideResource ? (weeklyCapacitiesByResource[selectedOverrideResource] ?? 0) : 0
  const selectedOverrideWeekly =
    selectedOverrideResource && selectedOverrideWeek
      ? weeklyCapacityOverridesByResource[selectedOverrideResource]?.[selectedOverrideWeek]
      : undefined
  const hasSelectedOverride = Number.isFinite(selectedOverrideWeekly)
  const selectedEffectiveWeekly = hasSelectedOverride ? Number(selectedOverrideWeekly) : selectedDefaultWeekly
  const selectedOverrideWeekSet = useMemo(() => new Set(selectedOverrideWeeks), [selectedOverrideWeeks])

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
            <span className={`chevron ${isCollapsed ? 'chevron-closed' : 'chevron-open'}`} aria-hidden="true">
              ▾
            </span>
            {isCollapsed ? 'Expand Table' : 'Collapse Table'}
          </button>
        </div>
      </div>

      <div className="weekly-capacity-override-panel">
        <div className="section-header">
          <h3>Weekly Capacity Overrides</h3>
          <p>Set a specific weekly capacity for a resource. Overrides replace the default weekly capacity for that week.</p>
        </div>

        {resources.length === 0 || weekKeys.length === 0 ? (
          <div className="status">No resources or weeks are available for override editing.</div>
        ) : (
          <>
            <div className="weekly-capacity-controls">
              <label>
                Resource
                <select
                  value={selectedOverrideResource}
                  onChange={(event) => setSelectedOverrideResource(event.target.value)}
                >
                  {resources.map((resource) => (
                    <option key={resource} value={resource}>
                      {resource}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Week
                <select value={selectedOverrideWeek} onChange={(event) => setSelectedOverrideWeek(event.target.value)}>
                  {weekKeys.map((weekKey) => (
                    <option key={weekKey} value={weekKey}>
                      {weekRangeLabel(weekKey)}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Override Capacity Hours
                <input
                  type="text"
                  inputMode="decimal"
                  value={draftWeekOverrideHours}
                  onChange={(event) => updateWeekOverrideDraft(event.target.value)}
                  onBlur={applyWeekOverride}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault()
                      applyWeekOverride()
                      ;(event.currentTarget as HTMLInputElement).blur()
                    }
                    if (event.key === 'Escape') {
                      event.preventDefault()
                      if (hasSelectedOverride) {
                        setDraftWeekOverrideHours(String(selectedOverrideWeekly))
                      } else {
                        setDraftWeekOverrideHours('')
                      }
                      ;(event.currentTarget as HTMLInputElement).blur()
                    }
                  }}
                  placeholder={hasSelectedOverride ? '' : `${selectedDefaultWeekly.toFixed(2)} default`}
                  aria-label="Weekly capacity override hours"
                />
              </label>
            </div>

            <div className="weekly-capacity-week-picker">
              <div className="weekly-capacity-week-picker-header">
                <strong>Weeks for Bulk Apply ({selectedOverrideWeeks.length} selected)</strong>
                <div className="weekly-capacity-week-picker-actions">
                  <button type="button" className="ghost-btn" onClick={selectAllWeeks}>
                    Select All Visible Weeks
                  </button>
                  <button type="button" className="ghost-btn" onClick={clearWeekSelection}>
                    Clear Week Selection
                  </button>
                </div>
              </div>
              <div className="weekly-capacity-week-list">
                {weekKeys.map((weekKey) => (
                  <label key={weekKey} className="weekly-capacity-week-option">
                    <input
                      type="checkbox"
                      checked={selectedOverrideWeekSet.has(weekKey)}
                      onChange={() => toggleSelectedWeek(weekKey)}
                    />
                    <span>{weekRangeLabel(weekKey)}</span>
                  </label>
                ))}
              </div>
            </div>

            <div className="weekly-capacity-summary">
              <span>
                <strong>Default:</strong> {selectedDefaultWeekly.toFixed(2)} h
              </span>
              <span>
                <strong>Override:</strong> {hasSelectedOverride ? `${Number(selectedOverrideWeekly).toFixed(2)} h` : 'None'}
              </span>
              <span>
                <strong>Effective:</strong> {selectedEffectiveWeekly.toFixed(2)} h
              </span>
              <span className={hasSelectedOverride ? 'weekly-capacity-tag weekly-capacity-tag-override' : 'weekly-capacity-tag'}>
                {hasSelectedOverride ? 'Override Active' : 'Using Default'}
              </span>
            </div>

            <div className="weekly-capacity-actions">
              <button type="button" onClick={applyWeekOverride} disabled={!selectedOverrideResource || !selectedOverrideWeek}>
                Save Override
              </button>
              <button type="button" onClick={applyOverrideToSelectedWeeks} disabled={!selectedOverrideResource}>
                Apply To Selected Weeks
              </button>
              <button
                type="button"
                className="ghost-btn"
                onClick={clearSelectedWeekOverride}
                disabled={!hasSelectedOverride}
              >
                Clear Selected Override
              </button>
              <button
                type="button"
                className="ghost-btn"
                onClick={clearSelectedWeeksOverrides}
                disabled={!selectedOverrideResource}
              >
                Clear Selected Weeks Overrides
              </button>
              <button
                type="button"
                className="ghost-btn"
                onClick={clearAllOverrides}
                disabled={weeklyOverrideRows.length === 0}
              >
                Clear All Overrides
              </button>
            </div>

            {bulkValidationMessage && <div className="error-text">{bulkValidationMessage}</div>}

            {weeklyOverrideRows.length > 0 && (
              <div className="table-wrap weekly-capacity-override-table-wrap">
                <table className="weekly-capacity-override-table">
                  <thead>
                    <tr>
                      <th>Resource</th>
                      <th>Week</th>
                      <th>Default</th>
                      <th>Override</th>
                      <th>Effective</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {weeklyOverrideRows.map((row) => (
                      <tr key={`${row.resource}::${row.weekIso}`}>
                        <td>{row.resource}</td>
                        <td>{weekRangeLabel(row.weekIso)}</td>
                        <td>{row.defaultWeekly.toFixed(2)}</td>
                        <td>{row.hours.toFixed(2)}</td>
                        <td>{row.hours.toFixed(2)}</td>
                        <td>
                          <button
                            type="button"
                            className="ghost-btn"
                            onClick={() => onClearWeeklyCapacityOverride(row.resource, row.weekIso)}
                          >
                            Clear
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
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
                <th>Holiday Impact (per holiday day)</th>
                <th>Effective Weekly Capacity</th>
                <th>Monthly Capacity Hours</th>
              </tr>
            </thead>
            <tbody>
              {resources.map((resource) => {
                const weekly = weeklyCapacitiesByResource[resource] ?? 0
                const manualWeekend = weekendExtraByResource[resource] ?? 0
                const holidayImpact = weekly / 5
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
                    <td>{holidayImpact.toFixed(2)}</td>
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
