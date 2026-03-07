import {
  addDays,
  addWeeks,
  compareAsc,
  eachDayOfInterval,
  endOfDay,
  format,
  getDay,
  getYear,
  isAfter,
  isBefore,
  parseISO,
  startOfDay,
  startOfWeek,
} from 'date-fns'
import type {
  AppFilters,
  CapacityStatus,
  ChartGroupBy,
  LeafCell,
  MonthlyBucket,
  PivotRowGrouping,
  PivotTableModel,
  TaskRow,
  WeeklyBucket,
} from '../types'

const MANUAL_RESOURCE = '__Manual Resource'
const MANUAL_PROJECT = '__Manual Project'
const KEY_SEPARATOR = '\u0001'

export function getCapacityStatus(forecastHours: number, capacityHours: number): CapacityStatus {
  // Weekly capacity is the source of truth. Status tolerance is +/-10% of capacity.
  // For zero-capacity weeks/months, only zero forecast counts as within capacity.
  if (capacityHours === 0) {
    return forecastHours === 0 ? 'Within Capacity' : 'Over Capacity'
  }

  const lowerBound = capacityHours * 0.9
  const upperBound = capacityHours * 1.1
  if (forecastHours >= lowerBound && forecastHours <= upperBound) {
    return 'Within Capacity'
  }

  return forecastHours > upperBound ? 'Over Capacity' : 'Under Capacity'
}

export function isoFromDate(date: Date): string {
  return format(date, 'yyyy-MM-dd')
}

export function localDateFromIso(iso: string): Date {
  return parseISO(iso)
}

export function weekRangeLabel(weekStartIso: string): string {
  const monday = localDateFromIso(weekStartIso)
  const friday = addDays(monday, 4)
  return `${format(monday, 'EEE MMM d')} - ${format(friday, 'EEE MMM d, yyyy')}`
}

function inDateRange(weekStart: Date, dateFrom: string, dateTo: string): boolean {
  const from = dateFrom ? startOfDay(localDateFromIso(dateFrom)) : null
  const to = dateTo ? endOfDay(localDateFromIso(dateTo)) : null

  if (from && isBefore(weekStart, from)) {
    return false
  }

  if (to && isAfter(weekStart, to)) {
    return false
  }

  return true
}

function yearMatches(weekStart: Date, selectedYear: string): boolean {
  if (!selectedYear) {
    return true
  }
  return getYear(weekStart) === Number(selectedYear)
}

export function makeLeafKey(project: string, resource: string, weekStartIso: string): string {
  return [project, resource, weekStartIso].join(KEY_SEPARATOR)
}

export function parseLeafKey(leafKey: string): { project: string; resource: string; weekStartIso: string } {
  const [project = MANUAL_PROJECT, resource = MANUAL_RESOURCE, weekStartIso = ''] = leafKey.split(KEY_SEPARATOR)
  return { project, resource, weekStartIso }
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
      return workingWeekendDates.has(isoFromDate(day))
    }
    return true
  })
  const days = activeDays.length > 0 ? activeDays : [startDate]
  const hoursPerDay = task.workHours / days.length

  return days.map((day) => ({ date: day, hours: hoursPerDay }))
}

export function getAvailableYears(tasks: TaskRow[]): string[] {
  const years = new Set<number>()

  for (const task of tasks) {
    const startYear = getYear(task.start)
    const finishYear = getYear(task.finish)
    const min = Math.min(startYear, finishYear)
    const max = Math.max(startYear, finishYear)
    for (let year = min; year <= max; year += 1) {
      years.add(year)
    }
  }

  return [...years].sort((a, b) => a - b).map(String)
}

export function buildBaseLeafCells(
  tasks: TaskRow[],
  filters: AppFilters,
  workingWeekendDates: Set<string>,
  enabledResources: Set<string>,
): { leafCells: LeafCell[]; weekKeys: string[]; projects: string[]; resources: string[] } {
  const filteredTasks =
    enabledResources.size > 0
      ? tasks.filter((task) => {
          if (!enabledResources.has(task.resourceName)) {
            return false
          }
          if (filters.resources && filters.resources.length > 0 && !filters.resources.includes(task.resourceName)) {
            return false
          }
          return true
        })
      : []
  const leafMap = new Map<string, number>()

  for (const task of filteredTasks) {
    const daily = distributeTaskWork(task, workingWeekendDates)

    for (const allocation of daily) {
      const monday = startOfWeek(allocation.date, { weekStartsOn: 1 })
      if (!inDateRange(monday, filters.dateFrom, filters.dateTo)) {
        continue
      }
      if (!yearMatches(monday, filters.year)) {
        continue
      }

      const weekStartIso = isoFromDate(monday)
      const leafKey = makeLeafKey(task.project, task.resourceName, weekStartIso)
      leafMap.set(leafKey, (leafMap.get(leafKey) ?? 0) + allocation.hours)
    }
  }

  const leafCells: LeafCell[] = [...leafMap.entries()].map(([leafKey, hours]) => {
    const { project, resource, weekStartIso } = parseLeafKey(leafKey)
    return { project, resource, weekStartIso, hours }
  })

  const weekSet = new Set(leafCells.map((cell) => cell.weekStartIso))
  const sortedExistingWeeks = [...weekSet].sort((a, b) => compareAsc(localDateFromIso(a), localDateFromIso(b)))

  const weekKeys: string[] = []
  if (sortedExistingWeeks.length > 0) {
    const first = localDateFromIso(sortedExistingWeeks[0])
    const last = localDateFromIso(sortedExistingWeeks[sortedExistingWeeks.length - 1])
    for (let cursor = first; compareAsc(cursor, last) <= 0; cursor = addWeeks(cursor, 1)) {
      weekKeys.push(isoFromDate(cursor))
    }
  }

  const projects = [...new Set(leafCells.map((cell) => cell.project))].sort((a, b) => a.localeCompare(b))
  const resources = [...new Set(leafCells.map((cell) => cell.resource))].sort((a, b) => a.localeCompare(b))

  return { leafCells, weekKeys, projects, resources }
}

export function buildPivotModel(
  finalLeafByKey: Record<string, number>,
  baseLeafByKey: Record<string, number>,
  weekKeys: string[],
  rowGrouping: PivotRowGrouping,
): PivotTableModel {
  const rows = new Map<string, { valuesByWeek: Record<string, number>; totalHours: number }>()
  const columnTotals: Record<string, number> = {}
  const editedRowWeekKeys = new Set<string>()

  for (const week of weekKeys) {
    columnTotals[week] = 0
  }

  for (const [leafKey, finalHours] of Object.entries(finalLeafByKey)) {
    if (!Number.isFinite(finalHours)) {
      continue
    }

    const { project, resource, weekStartIso } = parseLeafKey(leafKey)
    const rowKey = rowGrouping === 'project' ? project : resource

    const row = rows.get(rowKey) ?? { valuesByWeek: {}, totalHours: 0 }
    row.valuesByWeek[weekStartIso] = (row.valuesByWeek[weekStartIso] ?? 0) + finalHours
    row.totalHours += finalHours
    rows.set(rowKey, row)

    columnTotals[weekStartIso] = (columnTotals[weekStartIso] ?? 0) + finalHours

    const baseHours = baseLeafByKey[leafKey] ?? 0
    if (Math.abs(baseHours - finalHours) > 0.0001) {
      editedRowWeekKeys.add(`${rowKey}${KEY_SEPARATOR}${weekStartIso}`)
    }
  }

  const pivotRows = [...rows.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([rowKey, row]) => ({
      rowKey,
      rowLabel: rowKey,
      valuesByWeek: row.valuesByWeek,
      totalHours: row.totalHours,
    }))

  const grandTotal = Object.values(columnTotals).reduce((sum, value) => sum + value, 0)

  return {
    rows: pivotRows,
    weekKeys,
    columnTotals,
    grandTotal,
    editedRowWeekKeys,
  }
}

export function buildWeeklyBucketsFromLeaf(
  finalLeafByKey: Record<string, number>,
  weekKeys: string[],
  weekCapacities: Record<string, number>,
  chartGroupBy: ChartGroupBy,
): WeeklyBucket[] {
  const perWeek = new Map<string, { total: number; groups: Record<string, number> }>()

  for (const week of weekKeys) {
    perWeek.set(week, { total: 0, groups: {} })
  }

  for (const [leafKey, hours] of Object.entries(finalLeafByKey)) {
    if (!Number.isFinite(hours)) {
      continue
    }

    const { project, resource, weekStartIso } = parseLeafKey(leafKey)
    const weekEntry = perWeek.get(weekStartIso)
    if (!weekEntry) {
      continue
    }

    const groupKey = chartGroupBy === 'project' ? project : resource
    weekEntry.total += hours
    weekEntry.groups[groupKey] = (weekEntry.groups[groupKey] ?? 0) + hours
  }

  return weekKeys.map((weekStartIso) => {
    const weekStart = localDateFromIso(weekStartIso)
    const weekEnd = addDays(weekStart, 4)
    const weekData = perWeek.get(weekStartIso) ?? { total: 0, groups: {} }
    // Capacity is counted only for active weeks (weeks with selected-project forecast > 0).
    // This keeps weekly and monthly capacity aligned to the active planning scope.
    const hasActiveForecast = weekData.total > 0
    const capacity = hasActiveForecast ? (weekCapacities[weekStartIso] ?? 0) : 0
    const variance = weekData.total - capacity
    const status = getCapacityStatus(weekData.total, capacity)

    return {
      weekStartIso,
      weekEndIso: isoFromDate(weekEnd),
      weekLabel: weekRangeLabel(weekStartIso),
      totalHours: weekData.total,
      capacity,
      variance,
      overCapacity: status === 'Over Capacity',
      status,
      groups: weekData.groups,
    }
  })
}

export function buildMonthlyBuckets(
  weeklyBuckets: WeeklyBucket[],
): MonthlyBucket[] {
  const byMonth = new Map<string, { planned: number; capacity: number }>()

  for (const bucket of weeklyBuckets) {
    const monday = localDateFromIso(bucket.weekStartIso)
    const weekdays = Array.from({ length: 5 }, (_, offset) => addDays(monday, offset))
    const dailyForecast = bucket.totalHours / weekdays.length
    const dailyCapacity = bucket.capacity / weekdays.length

    // Allocate each weekday's portion into its calendar month for accurate month totals.
    for (const day of weekdays) {
      const monthKey = format(day, 'yyyy-MM')
      const entry = byMonth.get(monthKey) ?? { planned: 0, capacity: 0 }
      entry.planned += dailyForecast
      entry.capacity += dailyCapacity
      byMonth.set(monthKey, entry)
    }
  }

  return [...byMonth.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([monthKey, data]) => {
      const monthDate = parseISO(`${monthKey}-01`)
      const planned = Math.round(data.planned)
      const capacity = Math.round(data.capacity)
      const variance = planned - capacity
      const status = getCapacityStatus(planned, capacity)
      const overCapacity = status === 'Over Capacity'
      const underCapacity = status === 'Under Capacity'

      return {
        monthKey,
        monthLabel: format(monthDate, 'MMM yyyy'),
        plannedHours: planned,
        capacity,
        variance,
        overCapacity,
        underCapacity,
        status,
      }
    })
}

export function computeCategoryKeys(weeklyBuckets: WeeklyBucket[]): string[] {
  const set = new Set<string>()
  for (const bucket of weeklyBuckets) {
    Object.keys(bucket.groups).forEach((key) => set.add(key))
  }
  return [...set].sort((a, b) => a.localeCompare(b))
}

export function buildLeafValueMap(
  baseLeafCells: LeafCell[],
  manualOverrides: Record<string, number>,
  selectedProjects: Set<string>,
): { baseByKey: Record<string, number>; finalByKey: Record<string, number> } {
  const baseByKey: Record<string, number> = {}

  for (const cell of baseLeafCells) {
    if (!selectedProjects.has(cell.project)) {
      continue
    }

    const key = makeLeafKey(cell.project, cell.resource, cell.weekStartIso)
    baseByKey[key] = cell.hours
  }

  const finalByKey: Record<string, number> = { ...baseByKey }

  for (const [leafKey, overrideHours] of Object.entries(manualOverrides)) {
    const { project } = parseLeafKey(leafKey)
    if (!selectedProjects.has(project)) {
      continue
    }
    finalByKey[leafKey] = overrideHours
  }

  return { baseByKey, finalByKey }
}

export function editableLeafKeysForRowWeek(
  rowKey: string,
  weekStartIso: string,
  rowGrouping: PivotRowGrouping,
  allLeafKeys: string[],
  selectedProjects: Set<string>,
): string[] {
  return allLeafKeys.filter((leafKey) => {
    const parsed = parseLeafKey(leafKey)
    if (parsed.weekStartIso !== weekStartIso) {
      return false
    }

    if (!selectedProjects.has(parsed.project)) {
      return false
    }

    const candidateRowKey = rowGrouping === 'project' ? parsed.project : parsed.resource
    return candidateRowKey === rowKey
  })
}

export function makeSyntheticLeafKey(
  rowKey: string,
  weekStartIso: string,
  rowGrouping: PivotRowGrouping,
  selectedProjects: Set<string>,
): string {
  if (rowGrouping === 'project') {
    return makeLeafKey(rowKey, MANUAL_RESOURCE, weekStartIso)
  }

  const firstProject = [...selectedProjects][0] ?? MANUAL_PROJECT
  return makeLeafKey(firstProject, rowKey, weekStartIso)
}

export function shortWeekLabel(weekStartIso: string): string {
  return format(localDateFromIso(weekStartIso), 'MMM d')
}

export function rowWeekEditedKey(rowKey: string, weekStartIso: string): string {
  return `${rowKey}${KEY_SEPARATOR}${weekStartIso}`
}
