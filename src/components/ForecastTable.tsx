import { useMemo, useState } from 'react'
import clsx from 'clsx'
import type { WeeklyBucket } from '../types'
import { getCapacityStatus } from '../utils/planner'

type SortKey = 'weekStartIso' | 'totalHours' | 'capacity' | 'variance'
type SortDirection = 'asc' | 'desc'

interface ForecastTableProps {
  weeklyBuckets: WeeklyBucket[]
}

function nextDirection(currentKey: SortKey, clickedKey: SortKey, currentDirection: SortDirection): SortDirection {
  if (currentKey !== clickedKey) {
    return 'desc'
  }

  return currentDirection === 'desc' ? 'asc' : 'desc'
}

export function ForecastTable({ weeklyBuckets }: ForecastTableProps) {
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
          <p>Each row compares plan weekly hours against total selected resource capacity.</p>
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
                    Plan Weekly Hours
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('capacity')}>
                    Total Capacity
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
              {sortedRows.map((bucket) => {
                const status = getCapacityStatus(bucket.totalHours, bucket.capacity)
                return (
                  <tr
                    key={bucket.weekStartIso}
                    className={clsx({
                      'weekly-over-row': status === 'Over Capacity',
                      'weekly-under-row': status === 'Under Capacity',
                      'weekly-within-row': status === 'Within Capacity',
                    })}
                  >
                    <td>{bucket.weekLabel}</td>
                    <td>{bucket.totalHours.toFixed(2)}</td>
                    <td>{bucket.capacity.toFixed(2)}</td>
                    <td className={bucket.variance !== 0 ? 'negative' : ''}>
                      {bucket.variance.toFixed(2)}
                    </td>
                    <td>{status}</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
