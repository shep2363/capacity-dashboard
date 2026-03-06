import { useState } from 'react'

interface ResourceCapacityTableProps {
  resources: string[]
  enabledResources: Record<string, boolean>
  weeklyCapacitiesByResource: Record<string, number>
  onWeeklyCapacityChange: (resource: string, weeklyCapacity: number) => void
  onToggleResource: (resource: string, enabled: boolean) => void
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
}: ResourceCapacityTableProps) {
  const [isCollapsed, setIsCollapsed] = useState(false)

  return (
    <section className="panel table-panel">
      <div className="section-header section-header-row">
        <div>
          <h2>Resource Capacity Input</h2>
          <p>Enter weekly capacity by resource. Monthly capacity is derived automatically.</p>
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
                <th>Monthly Capacity Hours</th>
              </tr>
            </thead>
            <tbody>
              {resources.map((resource) => {
                const weekly = weeklyCapacitiesByResource[resource] ?? 0
                const monthly = monthlyFromWeekly(weekly)
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
                        type="number"
                        min={0}
                        step={1}
                        value={Number.isFinite(weekly) ? weekly : 0}
                        onChange={(event) => onWeeklyCapacityChange(resource, Number(event.target.value))}
                        aria-label={`Weekly capacity for ${resource}`}
                      />
                    </td>
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
