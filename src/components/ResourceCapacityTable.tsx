interface ResourceCapacityTableProps {
  resources: string[]
  selectedResources: Set<string>
  weeklyCapacitiesByResource: Record<string, number>
  onWeeklyCapacityChange: (resource: string, weeklyCapacity: number) => void
}

const WEEKS_PER_MONTH = 52 / 12

function monthlyFromWeekly(weekly: number): number {
  return weekly * WEEKS_PER_MONTH
}

export function ResourceCapacityTable({
  resources,
  selectedResources,
  weeklyCapacitiesByResource,
  onWeeklyCapacityChange,
}: ResourceCapacityTableProps) {
  return (
    <section className="panel table-panel">
      <div className="section-header">
        <h2>Resource Capacity Input</h2>
        <p>Enter weekly capacity by resource. Monthly capacity is derived automatically.</p>
      </div>

      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Resource Name</th>
              <th>Weekly Capacity Hours</th>
              <th>Monthly Capacity Hours</th>
            </tr>
          </thead>
          <tbody>
            {resources.map((resource) => {
              const weekly = weeklyCapacitiesByResource[resource] ?? 0
              const monthly = monthlyFromWeekly(weekly)
              const isSelected = selectedResources.has(resource)

              return (
                <tr key={resource} className={isSelected ? 'capacity-selected-row' : ''}>
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
    </section>
  )
}
