import { useMemo, useState } from 'react'
import { eachDayOfInterval, format, getDay, getYear, isAfter, isBefore, parseISO, startOfDay, startOfWeek } from 'date-fns'
import type { AppFilters, TaskRow } from '../types'
import { weekRangeLabel } from '../utils/planner'

const PROJECT_COLOR_PALETTE = [
  '#4f46e5',
  '#0f766e',
  '#f59e0b',
  '#0284c7',
  '#7c3aed',
  '#16a34a',
  '#9333ea',
  '#be123c',
  '#0ea5e9',
  '#64748b',
  '#ea580c',
  '#047857',
  '#3b82f6',
  '#10b981',
  '#e11d48',
  '#c084fc',
]

interface DepartmentPageProps {
  resource: string
  tasks: TaskRow[]
  filters: AppFilters
  selectedProjects: Set<string>
  selectedWeekendDates: Set<string>
  projectColors: Record<string, string>
  resourceEnabled: boolean
  filter: DepartmentFilters
  onFilterChange: (next: DepartmentFilters) => void
}

export interface DepartmentRow {
  project: string
  sequence: string
  weekStartIso: string
  weekLabel: string
  hours: number
  percentComplete: number
  finishDate: string
  status: string
}

export interface DepartmentFilters {
  projects: string[]
  sequences: string[]
  weeks: string[]
  statuses: string[]
}

function distributeTaskWork(task: TaskRow, workingWeekendDates: Set<string>): Array<{ date: Date; hours: number }> {
  const start = startOfDay(task.start)
  const finish = startOfDay(task.finish)
  const startDate = isAfter(start, finish) ? finish : start
  const endDate = isAfter(start, finish) ? start : finish

  const allDays = eachDayOfInterval({ start: startDate, end: endDate })
  const activeDays = allDays.filter((day) => {
    const dow = getDay(day)
    if (dow === 0 || dow === 6) {
      return workingWeekendDates.has(format(day, 'yyyy-MM-dd'))
    }
    return true
  })
  const days = activeDays.length > 0 ? activeDays : [startDate]
  const hoursPerDay = task.workHours / days.length

  return days.map((day) => ({ date: day, hours: hoursPerDay }))
}

export function buildDepartmentRows(params: {
  resource: string
  tasks: TaskRow[]
  filters: AppFilters
  selectedProjects: Set<string>
  selectedWeekendDates: Set<string>
  resourceEnabled: boolean
}): DepartmentRow[] {
  const { resource, tasks, filters, selectedProjects, selectedWeekendDates, resourceEnabled } = params
  if (!resourceEnabled) return []
  const filterYear = filters.year ? Number(filters.year) : null
  const parsedFrom = filters.dateFrom ? parseISO(filters.dateFrom) : null
  const parsedTo = filters.dateTo ? parseISO(filters.dateTo) : null
  const dateFrom = parsedFrom && !Number.isNaN(parsedFrom.getTime()) ? startOfDay(parsedFrom) : null
  const dateTo = parsedTo && !Number.isNaN(parsedTo.getTime()) ? startOfDay(parsedTo) : null
  const weekendSet = selectedWeekendDates
  const now = startOfDay(new Date())
  const result: DepartmentRow[] = []

  tasks.forEach((task) => {
    if (task.resourceName !== resource) return
    if (selectedProjects.size > 0 && !selectedProjects.has(task.project)) return

    const daily = distributeTaskWork(task, weekendSet)
    const workedHours = daily
      .filter(({ date }) => isBefore(date, now))
      .reduce((sum, { hours }) => sum + hours, 0)
    const percent = Math.min(100, Math.max(0, (workedHours / task.workHours) * 100))
    const status =
      percent >= 99.5
        ? 'Completed'
        : isAfter(now, task.finish)
        ? 'Overdue'
        : isBefore(now, task.start)
        ? 'Scheduled'
        : 'In Progress'

    const hoursByWeek = new Map<string, number>()
    daily.forEach(({ date, hours }) => {
      const weekStart = startOfWeek(date, { weekStartsOn: 1 })
      if (dateFrom && isBefore(weekStart, dateFrom)) return
      if (dateTo && isAfter(weekStart, dateTo)) return
      if (filterYear && getYear(weekStart) !== filterYear) return
      const weekIso = format(weekStart, 'yyyy-MM-dd')
      hoursByWeek.set(weekIso, (hoursByWeek.get(weekIso) ?? 0) + hours)
    })

    hoursByWeek.forEach((hours, weekStartIso) => {
      result.push({
        project: task.project,
        sequence: task.name,
        weekStartIso,
        weekLabel: weekRangeLabel(weekStartIso),
        hours,
        percentComplete: percent,
        finishDate: format(task.finish, 'yyyy-MM-dd'),
        status,
      })
    })
  })

  return result
}

function DepartmentPage({
  resource,
  tasks,
  filters,
  selectedProjects,
  selectedWeekendDates,
  projectColors,
  resourceEnabled,
  filter,
  onFilterChange,
}: DepartmentPageProps) {
  const rows = useMemo(
    () =>
      buildDepartmentRows({
        resource,
        tasks,
        filters,
        selectedProjects,
        selectedWeekendDates,
        resourceEnabled,
      }),
    [resource, tasks, filters, selectedProjects, selectedWeekendDates, resourceEnabled],
  )

  const projectFilteredRows = useMemo(() => {
    if (filter.projects.length === 0) return rows
    return rows.filter((row) => filter.projects.includes(row.project))
  }, [rows, filter.projects])

  const filteredRows = useMemo(() => {
    return projectFilteredRows.filter((row) => {
      if (filter.sequences.length > 0 && !filter.sequences.includes(row.sequence)) return false
      if (filter.weeks.length > 0 && !filter.weeks.includes(row.weekStartIso)) return false
       if (filter.statuses.length > 0 && !filter.statuses.includes(row.status)) return false
      return true
    })
  }, [projectFilteredRows, filter.sequences, filter.weeks, filter.statuses])

  const grouped = useMemo(() => {
    const map = new Map<
      string,
      { weekLabel: string; rows: DepartmentRow[]; totalHours: number; projectTotals: Record<string, number> }
    >()
    filteredRows.forEach((row) => {
      const entry = map.get(row.weekStartIso) ?? {
        weekLabel: weekRangeLabel(row.weekStartIso),
        rows: [],
        totalHours: 0,
        projectTotals: {},
      }
      entry.rows.push(row)
      entry.totalHours += row.hours
      entry.projectTotals[row.project] = (entry.projectTotals[row.project] ?? 0) + row.hours
      map.set(row.weekStartIso, entry)
    })
    return [...map.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([weekStartIso, data]) => ({ weekStartIso, ...data }))
  }, [filteredRows])

  const weekOptions = useMemo(
    () =>
      [...new Set(projectFilteredRows.map((r) => r.weekStartIso))]
        .sort((a, b) => a.localeCompare(b))
        .map((iso) => ({ value: iso, label: weekRangeLabel(iso) })),
    [projectFilteredRows],
  )
  const projectOptions = useMemo(
    () =>
      [...new Set(rows.map((r) => r.project))]
        .sort((a, b) => a.localeCompare(b))
        .map((project) => ({ value: project, label: project })),
    [rows],
  )
  const sequenceOptions = useMemo(
    () =>
      [...new Set(projectFilteredRows.map((r) => r.sequence))]
        .sort((a, b) => a.localeCompare(b))
        .map((sequence) => ({ value: sequence, label: sequence })),
    [projectFilteredRows],
  )
  const statusOptions = useMemo(
    () =>
      [...new Set(projectFilteredRows.map((r) => r.status))]
        .sort((a, b) => a.localeCompare(b))
        .map((status) => ({ value: status, label: status })),
    [projectFilteredRows],
  )

  const totalHours = filteredRows.reduce((sum, row) => sum + row.hours, 0)
  const totalSequences = filteredRows.length

  const getProjectColor = (project: string) => {
    if (projectColors[project]) return projectColors[project]
    const hash = project.split('').reduce((sum, char) => sum + char.charCodeAt(0), 0)
    return PROJECT_COLOR_PALETTE[hash % PROJECT_COLOR_PALETTE.length]
  }

  const handleMultiSelect = (event: React.ChangeEvent<HTMLSelectElement>, setter: (values: string[]) => void) => {
    const values = Array.from(event.target.selectedOptions).map((option) => option.value)
    setter(values)
  }

  const setProjects = (values: string[]) => onFilterChange({ ...filter, projects: values })
  const setSequences = (values: string[]) => onFilterChange({ ...filter, sequences: values })
  const setWeeks = (values: string[]) => onFilterChange({ ...filter, weeks: values })
  const setStatuses = (values: string[]) => onFilterChange({ ...filter, statuses: values })
  const resetFilters = () => onFilterChange({ projects: [], sequences: [], weeks: [], statuses: [] })
  const [filtersOpen, setFiltersOpen] = useState(true)

  if (!resourceEnabled) {
    return (
      <section className="panel department-page">
        <div className="section-header">
          <h2>{resource} — Weekly Plan</h2>
          <p>This resource is currently disabled in the Resource Capacity Input toggles.</p>
        </div>
      </section>
    )
  }

  return (
    <section className="panel department-page">
      <div className="section-header-row">
        <div>
          <h2>{resource} — Weekly Production Plan</h2>
          <p>Project and sequence view filtered to the currently loaded workbook and dashboard filters.</p>
        </div>
        <div className="dept-summary">
          <div>
            <span>Total Hours</span>
            <strong>{totalHours.toFixed(1)}</strong>
          </div>
          <div>
            <span>Sequences</span>
            <strong>{totalSequences}</strong>
          </div>
          <button
            type="button"
            className="ghost-btn collapse-toggle"
            onClick={() => setFiltersOpen((open) => !open)}
            aria-expanded={filtersOpen}
          >
            {filtersOpen ? 'Hide Filters' : 'Show Filters'}
          </button>
        </div>
      </div>

      <div className={`dept-filters ${filtersOpen ? '' : 'collapsed'}`}>
        <label>
          Project Filter
          <select multiple value={filter.projects} onChange={(event) => handleMultiSelect(event, setProjects)}>
            {projectOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label>
          Sequence Filter
          <select multiple value={filter.sequences} onChange={(event) => handleMultiSelect(event, setSequences)}>
            {sequenceOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label>
          Week Filter
          <select multiple value={filter.weeks} onChange={(event) => handleMultiSelect(event, setWeeks)}>
            {weekOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label>
          Status Filter
          <select multiple value={filter.statuses} onChange={(event) => handleMultiSelect(event, setStatuses)}>
            {statusOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <div className="dept-filter-actions">
          <button type="button" className="ghost-btn" onClick={resetFilters}>
            Reset Filters
          </button>
        </div>
      </div>

      {grouped.length === 0 ? (
        <div className="status">No sequences match the current filters.</div>
      ) : (
        grouped.map((week) => (
          <div key={week.weekStartIso} className="dept-week">
            <div className="dept-week-header">
              <div>
                <h3>{week.weekLabel}</h3>
                <p>{week.rows.length} sequence{week.rows.length === 1 ? '' : 's'}</p>
              </div>
              <div className="dept-week-totals">
                <span className="pill pill-ghost">Week Hours: {week.totalHours.toFixed(1)}</span>
                <div className="dept-project-totals">
                  {Object.entries(week.projectTotals)
                    .sort((a, b) => b[1] - a[1])
                    .map(([project, hours]) => (
                      <span
                        key={project}
                        className="project-chip"
                        style={{ backgroundColor: getProjectColor(project) }}
                        title={project}
                      >
                        {project}: {hours.toFixed(1)}h
                      </span>
                    ))}
                </div>
              </div>
            </div>

            <div className="dept-table">
              <div className="dept-table-header">
                <span>Project</span>
                <span>Sequence</span>
                <span>Finish Date</span>
                <span>Hours</span>
                <span>Progress</span>
                <span>Status</span>
              </div>
              {week.rows
                .sort((a, b) => a.project.localeCompare(b.project) || a.sequence.localeCompare(b.sequence))
                .map((row, idx) => (
                  <div key={`${row.project}-${row.sequence}-${row.weekStartIso}-${idx}`} className="dept-row">
                    <span className="project-cell">
                      <span className="project-dot" style={{ backgroundColor: getProjectColor(row.project) }} />
                      {row.project}
                    </span>
                    <span>{row.sequence}</span>
                    <span>{row.finishDate || 'Not Scheduled'}</span>
                    <span>{row.hours.toFixed(1)} h</span>
                    <span className="dept-progress">
                      <div className="dept-progress-bar" aria-label={`Progress ${row.percentComplete.toFixed(0)}%`}>
                        <div
                          className="dept-progress-fill"
                          style={{ width: `${row.percentComplete.toFixed(0)}%` }}
                        />
                      </div>
                      <span>{row.percentComplete.toFixed(0)}%</span>
                    </span>
                    <span className={`status-pill status-${row.status.replace(/\s+/g, '-').toLowerCase()}`}>
                      {row.status}
                    </span>
                  </div>
                ))}
            </div>
          </div>
        ))
      )}
    </section>
  )
}

export default DepartmentPage

