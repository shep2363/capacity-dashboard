import { useMemo, useState } from 'react'
import clsx from 'clsx'
import type { WeeklyBucket } from '../types'

type SortKey = 'weekStartIso' | 'totalHours' | 'capacity' | 'variance'
type SortDirection = 'asc' | 'desc'

interface ForecastTableProps {
  weeklyBuckets: WeeklyBucket[]
  onWeekCapacityChange: (weekIso: string, capacity: number) => void
}

function nextDirection(currentKey: SortKey, clickedKey: SortKey, currentDirection: SortDirection): SortDirection {
  if (currentKey !== clickedKey) {
    return 'desc'
  }

  return currentDirection === 'desc' ? 'asc' : 'desc'
}

export function ForecastTable({ weeklyBuckets, onWeekCapacityChange }: ForecastTableProps) {
  const [sortKey, setSortKey] = useState<SortKey>('weekStartIso')
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc')
  const [isCollapsed, setIsCollapsed] = useState(false)

  const sortedRows = useMemo(() => {
    return [...weeklyBuckets].sort((a, b) => {
      const left = a[sortKey]
      const right = b[sortKey]

      if (left < right) {
        return sortDirection === 'asc' ? -1 : 1
      }

      if (left > right) {
        return sortDirection === 'asc' ? 1 : -1
      }

      return 0
    })
  }, [weeklyBuckets, sortDirection, sortKey])

  function handleSort(clickedKey: SortKey): void {
    setSortDirection((currentDirection) => nextDirection(sortKey, clickedKey, currentDirection))
    setSortKey(clickedKey)
  }

  return (
    <div className="panel table-panel">
      <div className="section-header section-header-row">
        <div>
          <h2>Weekly Forecast Table</h2>
          <p>Each row is one Monday-Friday bucket with editable capacity and variance.</p>
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
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('weekStartIso')}>
                    Week (Mon-Fri)
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('totalHours')}>
                    Forecast Hours
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('capacity')}>
                    Capacity Hours
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('variance')}>
                    Variance (Forecast - Capacity)
                  </button>
                </th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {sortedRows.map((bucket) => (
                <tr key={bucket.weekStartIso} className={clsx({ 'over-row': bucket.overCapacity })}>
                  <td>{bucket.weekLabel}</td>
                  <td>{bucket.totalHours.toFixed(2)}</td>
                  <td>
                    <input
                      className="capacity-cell"
                      type="number"
                      min={0}
                      step={1}
                      value={Number.isFinite(bucket.capacity) ? bucket.capacity : 0}
                      onChange={(event) => onWeekCapacityChange(bucket.weekStartIso, Number(event.target.value))}
                      aria-label={`Capacity for week ${bucket.weekLabel}`}
                    />
                  </td>
                  <td className={bucket.overCapacity ? 'negative' : 'positive'}>{bucket.variance.toFixed(2)}</td>
                  <td>{bucket.overCapacity ? 'Over capacity' : 'Within capacity'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
