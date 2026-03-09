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
}

interface DepartmentRow {
  project: string
  sequence: string
  weekStartIso: string
  hours: number
  percentComplete: number
  remainingHours: number
  status: string
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

function DepartmentPage({
  resource,
  tasks,
  filters,
  selectedProjects,
  selectedWeekendDates,
  projectColors,
  resourceEnabled,
}: DepartmentPageProps) {
  const [projectFilter, setProjectFilter] = useState<string[]>([])
  const [sequenceFilter, setSequenceFilter] = useState<string[]>([])
  const [weekFilter, setWeekFilter] = useState<string[]>([])

  const filterYear = filters.year ? Number(filters.year) : null
  const parsedFrom = filters.dateFrom ? parseISO(filters.dateFrom) : null
  const parsedTo = filters.dateTo ? parseISO(filters.dateTo) : null
  const dateFrom = parsedFrom && !Number.isNaN(parsedFrom.getTime()) ? startOfDay(parsedFrom) : null
  const dateTo = parsedTo && !Number.isNaN(parsedTo.getTime()) ? startOfDay(parsedTo) : null

  const rows = useMemo<DepartmentRow[]>(() => {
    if (!resourceEnabled) return []
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
      const remaining = Math.max(0, task.workHours - workedHours)
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
          hours,
          percentComplete: percent,
          remainingHours: remaining,
          status,
        })
      })
    })

    return result
  }, [tasks, resource, selectedProjects, selectedWeekendDates, filterYear, dateFrom, dateTo, resourceEnabled])

  const filteredRows = useMemo(() => {
    return rows.filter((row) => {
      if (projectFilter.length > 0 && !projectFilter.includes(row.project)) return false
      if (sequenceFilter.length > 0 && !sequenceFilter.includes(row.sequence)) return false
      if (weekFilter.length > 0 && !weekFilter.includes(row.weekStartIso)) return false
      return true
    })
  }, [rows, projectFilter, sequenceFilter, weekFilter])

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
      [...new Set(rows.map((r) => r.weekStartIso))]
        .sort((a, b) => a.localeCompare(b))
        .map((iso) => ({ value: iso, label: weekRangeLabel(iso) })),
    [rows],
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
      [...new Set(rows.map((r) => r.sequence))]
        .sort((a, b) => a.localeCompare(b))
        .map((sequence) => ({ value: sequence, label: sequence })),
    [rows],
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
        </div>
      </div>

      <div className="dept-filters">
        <label>
          Project Filter
          <select multiple value={projectFilter} onChange={(event) => handleMultiSelect(event, setProjectFilter)}>
            {projectOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label>
          Sequence Filter
          <select multiple value={sequenceFilter} onChange={(event) => handleMultiSelect(event, setSequenceFilter)}>
            {sequenceOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
        <label>
          Week Filter
          <select multiple value={weekFilter} onChange={(event) => handleMultiSelect(event, setWeekFilter)}>
            {weekOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </label>
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
                <span>Hours</span>
                <span>Progress</span>
                <span>Remaining</span>
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
                    <span>{row.remainingHours.toFixed(1)} h</span>
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
