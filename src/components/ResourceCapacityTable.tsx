import { useEffect, useMemo, useState } from 'react'
import type { Dispatch, SetStateAction } from 'react'
import { weekRangeLabel } from '../utils/planner'

interface ResourceCapacityTableProps {
  resources: string[]
  enabledResources: Record<string, boolean>
  weeklyCapacitiesByResource: Record<string, number>
  weekKeys: string[]
  totalWeekCapacitySchedule: Record<string, number>
  baseTotalCapacityByWeek: Record<string, number>
  effectiveTotalCapacityByWeek: Record<string, number>
  onWeeklyCapacityChange: (resource: string, weeklyCapacity: number) => void
  onSetTotalWeekCapacityForWeek: (weekStartIso: string, weeklyCapacity: number) => void
  onSetTotalWeekCapacityFromWeekForward: (weekStartIso: string, weeklyCapacity: number) => void
  onClearTotalWeekCapacityEntry: (weekStartIso: string) => void
  onClearAllTotalWeekCapacityEntries: () => void
  onToggleResource: (resource: string, enabled: boolean) => void
  weekendExtraByResource: Record<string, number>
  onWeekendExtraChange: (resource: string, weekendHours: number) => void
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
  totalWeekCapacitySchedule,
  baseTotalCapacityByWeek,
  effectiveTotalCapacityByWeek,
  onWeeklyCapacityChange,
  onSetTotalWeekCapacityForWeek,
  onSetTotalWeekCapacityFromWeekForward,
  onClearTotalWeekCapacityEntry,
  onClearAllTotalWeekCapacityEntries,
  onToggleResource,
  weekendExtraByResource,
  onWeekendExtraChange,
}: ResourceCapacityTableProps) {
  const [isCollapsed, setIsCollapsed] = useState(true)
  const [draftWeeklyByResource, setDraftWeeklyByResource] = useState<Record<string, string>>({})
  const [draftWeekendByResource, setDraftWeekendByResource] = useState<Record<string, string>>({})
  const [selectedCapacityWeek, setSelectedCapacityWeek] = useState('')
  const [draftScheduledCapacityHours, setDraftScheduledCapacityHours] = useState('')
  const [scheduleValidationMessage, setScheduleValidationMessage] = useState('')

  const resourceSet = useMemo(() => new Set(resources), [resources])
  const weekSet = useMemo(() => new Set(weekKeys), [weekKeys])

  const scheduleRows = useMemo(() => {
    const explicitWeeks = weekKeys
      .filter((weekIso) => Object.prototype.hasOwnProperty.call(totalWeekCapacitySchedule, weekIso))
      .sort((a, b) => a.localeCompare(b))

    return explicitWeeks.map((weekIso, index) => ({
      weekIso,
      hours: totalWeekCapacitySchedule[weekIso] ?? 0,
      nextWeekIso: explicitWeeks[index + 1] ?? null,
    }))
  }, [totalWeekCapacitySchedule, weekKeys])

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
    if (weekKeys.length === 0) {
      setSelectedCapacityWeek('')
      return
    }
    setSelectedCapacityWeek((current) => (current && weekSet.has(current) ? current : weekKeys[0]))
  }, [weekKeys, weekSet])

  useEffect(() => {
    if (!selectedCapacityWeek) {
      setDraftScheduledCapacityHours('')
      return
    }
    const scheduledCapacity = totalWeekCapacitySchedule[selectedCapacityWeek]
    if (Number.isFinite(scheduledCapacity)) {
      setDraftScheduledCapacityHours(String(scheduledCapacity))
      return
    }
    setDraftScheduledCapacityHours('')
  }, [selectedCapacityWeek, totalWeekCapacitySchedule])

  useEffect(() => {
    setScheduleValidationMessage('')
  }, [selectedCapacityWeek])

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

  function updateScheduledCapacityDraft(rawValue: string): void {
    if (rawValue !== '' && !/^\d*\.?\d*$/.test(rawValue)) {
      return
    }
    setDraftScheduledCapacityHours(rawValue)
  }

  function applyCapacityToThisWeek(): void {
    if (!selectedCapacityWeek) {
      setScheduleValidationMessage('Select a week before applying total capacity.')
      return
    }
    if (draftScheduledCapacityHours.trim() === '') {
      setScheduleValidationMessage('Enter total weekly capacity hours before saving.')
      return
    }
    const committed = parseCommittedHours(draftScheduledCapacityHours)
    onSetTotalWeekCapacityForWeek(selectedCapacityWeek, committed)
    setDraftScheduledCapacityHours(String(committed))
    setScheduleValidationMessage('')
  }

  function applyCapacityFromWeekForward(): void {
    if (!selectedCapacityWeek) {
      setScheduleValidationMessage('Select a week before applying total capacity.')
      return
    }
    if (draftScheduledCapacityHours.trim() === '') {
      setScheduleValidationMessage('Enter total weekly capacity hours before saving.')
      return
    }
    const committed = parseCommittedHours(draftScheduledCapacityHours)
    onSetTotalWeekCapacityFromWeekForward(selectedCapacityWeek, committed)
    setDraftScheduledCapacityHours(String(committed))
    setScheduleValidationMessage('')
  }

  function clearSelectedWeekSetting(): void {
    if (!selectedCapacityWeek) {
      return
    }
    onClearTotalWeekCapacityEntry(selectedCapacityWeek)
    setDraftScheduledCapacityHours('')
    setScheduleValidationMessage('')
  }

  function clearAllScheduleEntries(): void {
    if (scheduleRows.length === 0) {
      return
    }
    const confirmed = window.confirm('Clear all total weekly capacity schedule entries?')
    if (!confirmed) {
      return
    }
    onClearAllTotalWeekCapacityEntries()
    setDraftScheduledCapacityHours('')
    setScheduleValidationMessage('')
  }

  const selectedBaseTotal = selectedCapacityWeek ? (baseTotalCapacityByWeek[selectedCapacityWeek] ?? 0) : 0
  const selectedScheduledTotal = selectedCapacityWeek ? totalWeekCapacitySchedule[selectedCapacityWeek] : undefined
  const hasSelectedScheduleEntry = Number.isFinite(selectedScheduledTotal)
  const selectedEffectiveTotal = selectedCapacityWeek
    ? (effectiveTotalCapacityByWeek[selectedCapacityWeek] ?? selectedBaseTotal)
    : 0

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
              v
            </span>
            {isCollapsed ? 'Expand Table' : 'Collapse Table'}
          </button>
        </div>
      </div>

      {!isCollapsed && (
        <>
          <div className="total-weekly-capacity-panel">
            <div className="section-header">
              <h3>Total Weekly Capacity</h3>
              <p>
                Set a total capacity schedule for the full week. Explicit entries become breakpoints and stay active
                until the next scheduled change.
              </p>
            </div>

            {weekKeys.length === 0 ? (
              <div className="status">No weeks are available for total weekly capacity planning.</div>
            ) : (
              <>
                <div className="total-weekly-capacity-controls">
                  <label>
                    Week
                    <select value={selectedCapacityWeek} onChange={(event) => setSelectedCapacityWeek(event.target.value)}>
                      {weekKeys.map((weekKey) => (
                        <option key={weekKey} value={weekKey}>
                          {weekRangeLabel(weekKey)}
                        </option>
                      ))}
                    </select>
                  </label>

                  <label>
                    Total Capacity Hours
                    <input
                      type="text"
                      inputMode="decimal"
                      value={draftScheduledCapacityHours}
                      onChange={(event) => updateScheduledCapacityDraft(event.target.value)}
                      placeholder={hasSelectedScheduleEntry ? '' : `${selectedBaseTotal.toFixed(2)} base`}
                      aria-label="Total weekly capacity hours"
                    />
                  </label>
                </div>

                <div className="total-weekly-capacity-summary">
                  <span>
                    <strong>Base Total:</strong> {selectedBaseTotal.toFixed(2)} h
                  </span>
                  <span>
                    <strong>Scheduled Entry:</strong>{' '}
                    {hasSelectedScheduleEntry ? `${Number(selectedScheduledTotal).toFixed(2)} h` : 'None'}
                  </span>
                  <span>
                    <strong>Effective Total:</strong> {selectedEffectiveTotal.toFixed(2)} h
                  </span>
                  <span
                    className={
                      hasSelectedScheduleEntry
                        ? 'total-weekly-capacity-tag total-weekly-capacity-tag-scheduled'
                        : 'total-weekly-capacity-tag'
                    }
                  >
                    {hasSelectedScheduleEntry ? 'Scheduled Entry Active' : 'Using Base Total'}
                  </span>
                </div>

                <div className="total-weekly-capacity-actions">
                  <button type="button" onClick={applyCapacityToThisWeek} disabled={!selectedCapacityWeek}>
                    Apply to This Week
                  </button>
                  <button type="button" onClick={applyCapacityFromWeekForward} disabled={!selectedCapacityWeek}>
                    Apply From This Week Forward
                  </button>
                  <button
                    type="button"
                    className="ghost-btn"
                    onClick={clearSelectedWeekSetting}
                    disabled={!hasSelectedScheduleEntry}
                  >
                    Clear This Week Setting
                  </button>
                  <button
                    type="button"
                    className="ghost-btn"
                    onClick={clearAllScheduleEntries}
                    disabled={scheduleRows.length === 0}
                  >
                    Clear All Capacity Schedule Entries
                  </button>
                </div>

                {scheduleValidationMessage && <div className="error-text">{scheduleValidationMessage}</div>}

                {scheduleRows.length > 0 && (
                  <div className="table-wrap total-weekly-capacity-table-wrap">
                    <table className="total-weekly-capacity-table">
                      <thead>
                        <tr>
                          <th>Week</th>
                          <th>Scheduled Total</th>
                          <th>Effective Behavior</th>
                          <th>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {scheduleRows.map((row) => (
                          <tr key={row.weekIso}>
                            <td>{weekRangeLabel(row.weekIso)}</td>
                            <td>{row.hours.toFixed(2)}</td>
                            <td>
                              {row.nextWeekIso
                                ? `Active until next change on ${weekRangeLabel(row.nextWeekIso)}`
                                : 'Active through remaining visible weeks'}
                            </td>
                            <td>
                              <button
                                type="button"
                                className="ghost-btn"
                                onClick={() => onClearTotalWeekCapacityEntry(row.weekIso)}
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
        </>
      )}
    </section>
  )
}
