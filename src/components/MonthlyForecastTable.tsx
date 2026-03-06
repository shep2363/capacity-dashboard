import { useMemo, useState } from 'react'
import clsx from 'clsx'
import type { MonthlyBucket } from '../types'

type SortKey = 'monthKey' | 'plannedHours' | 'capacity' | 'variance'
type SortDirection = 'asc' | 'desc'

interface MonthlyForecastTableProps {
  monthlyBuckets: MonthlyBucket[]
}

function nextDirection(currentKey: SortKey, clickedKey: SortKey, currentDirection: SortDirection): SortDirection {
  if (currentKey !== clickedKey) {
    return 'asc'
  }

  return currentDirection === 'asc' ? 'desc' : 'asc'
}

export function MonthlyForecastTable({ monthlyBuckets }: MonthlyForecastTableProps) {
  const [sortKey, setSortKey] = useState<SortKey>('monthKey')
  const [sortDirection, setSortDirection] = useState<SortDirection>('asc')
  const [isCollapsed, setIsCollapsed] = useState(false)

  const sortedRows = useMemo(() => {
    return [...monthlyBuckets].sort((a, b) => {
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
  }, [monthlyBuckets, sortDirection, sortKey])

  function handleSort(clickedKey: SortKey): void {
    setSortDirection((currentDirection) => nextDirection(sortKey, clickedKey, currentDirection))
    setSortKey(clickedKey)
  }

  return (
    <div className="panel table-panel">
      <div className="section-header section-header-row">
        <div>
          <h2>Monthly Forecast Table</h2>
          <p>Monthly plan versus total selected resource capacity.</p>
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
                  <button type="button" className="sortable" onClick={() => handleSort('monthKey')}>
                    Month
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('plannedHours')}>
                    Planned Hours
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('capacity')}>
                    Total Capacity
                  </button>
                </th>
                <th>
                  <button type="button" className="sortable" onClick={() => handleSort('variance')}>
                    Variance (Planned - Capacity)
                  </button>
                </th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {sortedRows.map((bucket) => (
                <tr
                  key={bucket.monthKey}
                  className={clsx({
                    'monthly-over-row': bucket.overCapacity,
                    'monthly-under-row': bucket.underCapacity,
                  })}
                >
                  <td>{bucket.monthLabel}</td>
                  <td>{bucket.plannedHours.toFixed(2)}</td>
                  <td>{bucket.capacity.toFixed(2)}</td>
                  <td className={bucket.variance !== 0 ? 'negative' : ''}>
                    {bucket.variance.toFixed(2)}
                  </td>
                  <td>{bucket.status}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
