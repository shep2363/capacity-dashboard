import { useMemo, useState } from 'react'
import { format, isPast, isThisWeek, parseISO, startOfDay } from 'date-fns'
import type { AppFilters, TaskRow } from '../types'

interface DetailingPageProps {
  tasks: TaskRow[]
  filters: AppFilters
  selectedProjects: Set<string>
  projectColors: Record<string, string>
}

interface ReleaseRow {
  project: string
  sequence: string
  startDate: string
  releaseDate: string
  releaseDateObj: Date
  status: 'Released' | 'This Week' | 'Upcoming'
}

function getStatus(releaseDate: Date): ReleaseRow['status'] {
  if (isThisWeek(releaseDate, { weekStartsOn: 1 })) return 'This Week'
  if (isPast(releaseDate)) return 'Released'
  return 'Upcoming'
}

export function DetailingPage({ tasks, filters, selectedProjects, projectColors }: DetailingPageProps) {
  const [projectFilter, setProjectFilter] = useState<string>('all')
  const [sortField, setSortField] = useState<'project' | 'releaseDate'>('releaseDate')
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc')

  const releaseRows: ReleaseRow[] = useMemo(() => {
    const filterYear = filters.year ? Number(filters.year) : null
    const parsedFrom = filters.dateFrom ? parseISO(filters.dateFrom) : null
    const parsedTo = filters.dateTo ? parseISO(filters.dateTo) : null

    return tasks
      .filter((task) => {
        if (task.resourceName !== 'Detailing') return false
        if (selectedProjects.size > 0 && !selectedProjects.has(task.project)) return false
        if (filterYear && task.finish.getFullYear() !== filterYear) return false
        if (parsedFrom && task.finish < parsedFrom) return false
        if (parsedTo && task.finish > parsedTo) return false
        return true
      })
      .map((task) => {
        const releaseDateObj = startOfDay(task.finish)
        return {
          project: task.project,
          sequence: task.name,
          startDate: format(task.start, 'MMM d, yyyy'),
          releaseDate: format(releaseDateObj, 'MMM d, yyyy'),
          releaseDateObj,
          status: getStatus(releaseDateObj),
        }
      })
  }, [tasks, filters, selectedProjects])

  const allProjects = useMemo(
    () => Array.from(new Set(releaseRows.map((r) => r.project))).sort(),
    [releaseRows],
  )

  const filtered = useMemo(() => {
    const rows = projectFilter === 'all' ? releaseRows : releaseRows.filter((r) => r.project === projectFilter)
    return [...rows].sort((a, b) => {
      let cmp = 0
      if (sortField === 'releaseDate') {
        cmp = a.releaseDateObj.getTime() - b.releaseDateObj.getTime()
        if (cmp === 0) cmp = a.project.localeCompare(b.project)
      } else {
        cmp = a.project.localeCompare(b.project)
        if (cmp === 0) cmp = a.releaseDateObj.getTime() - b.releaseDateObj.getTime()
      }
      return sortDir === 'asc' ? cmp : -cmp
    })
  }, [releaseRows, projectFilter, sortField, sortDir])

  function handleSort(field: 'project' | 'releaseDate') {
    if (sortField === field) {
      setSortDir((d) => (d === 'asc' ? 'desc' : 'asc'))
    } else {
      setSortField(field)
      setSortDir('asc')
    }
  }

  const sortIcon = (field: 'project' | 'releaseDate') => {
    if (sortField !== field) return ' ↕'
    return sortDir === 'asc' ? ' ↑' : ' ↓'
  }

  const statusCounts = useMemo(
    () => ({
      released: filtered.filter((r) => r.status === 'Released').length,
      thisWeek: filtered.filter((r) => r.status === 'This Week').length,
      upcoming: filtered.filter((r) => r.status === 'Upcoming').length,
    }),
    [filtered],
  )

  return (
    <section className="panel dept-page">
      <div className="section-header">
        <div>
          <h2>Detailing — Release Schedule</h2>
          <p>Planned release dates for all Detailing tasks derived from the current workbook.</p>
        </div>
      </div>

      <div className="dept-filter-bar">
        <div className="dept-filter-group">
          <label htmlFor="detailing-project-filter">Project</label>
          <select
            id="detailing-project-filter"
            value={projectFilter}
            onChange={(e) => setProjectFilter(e.target.value)}
          >
            <option value="all">All Projects</option>
            {allProjects.map((p) => (
              <option key={p} value={p}>{p}</option>
            ))}
          </select>
        </div>

        <div className="dept-filter-stats">
          <span className="dept-stat">
            <span className="dept-stat-dot" style={{ background: '#22c55e' }} />
            Released: {statusCounts.released}
          </span>
          <span className="dept-stat">
            <span className="dept-stat-dot" style={{ background: '#f59e0b' }} />
            This Week: {statusCounts.thisWeek}
          </span>
          <span className="dept-stat">
            <span className="dept-stat-dot" style={{ background: '#3b82f6' }} />
            Upcoming: {statusCounts.upcoming}
          </span>
        </div>
      </div>

      {filtered.length === 0 ? (
        <div className="panel status">
          No Detailing tasks found. Ensure tasks in the workbook have a resource named &ldquo;Detailing&rdquo;.
        </div>
      ) : (
        <div className="dept-table-wrap">
          <table className="dept-table release-table">
            <thead>
              <tr className="dept-table-header">
                <th
                  className="dept-th-sortable"
                  onClick={() => handleSort('project')}
                >
                  Project{sortIcon('project')}
                </th>
                <th>Sequence</th>
                <th>Start Date</th>
                <th
                  className="dept-th-sortable"
                  onClick={() => handleSort('releaseDate')}
                >
                  Release Date{sortIcon('releaseDate')}
                </th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((row, i) => {
                const color = projectColors[row.project] ?? '#64748b'
                const statusClass =
                  row.status === 'Released'
                    ? 'dept-status-complete'
                    : row.status === 'This Week'
                    ? 'dept-status-warning'
                    : 'dept-status-scheduled'
                return (
                  <tr key={`${row.project}-${row.sequence}-${i}`} className="dept-table-row">
                    <td>
                      <span className="dept-project-chip" style={{ borderColor: color, color }}>
                        {row.project}
                      </span>
                    </td>
                    <td>{row.sequence}</td>
                    <td>{row.startDate}</td>
                    <td className="dept-release-date">{row.releaseDate}</td>
                    <td>
                      <span className={`dept-status-badge ${statusClass}`}>{row.status}</span>
                    </td>
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
