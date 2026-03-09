import { useEffect, useMemo, useState } from 'react'
import { addDays, format, parseISO, startOfWeek } from 'date-fns'
import { PivotPlanningTable } from './components/PivotPlanningTable'
import { ReportWorkspace, type ReportTab } from './components/ReportWorkspace'
import { ResourceCapacityTable } from './components/ResourceCapacityTable'
import { type ExecutiveData } from './components/ExecutiveSummary'
import type { AppFilters, ChartGroupBy, PivotRowGrouping, TaskRow } from './types'
import { parseSalesSpreadsheet, parseSpreadsheet } from './utils/excel'
import { exportReportWorkbook, type SummaryMetric } from './utils/reportExport'
import { exportReportWorkbookWithChartApi } from './utils/reportExportApi'
import {
  buildBaseLeafCells,
  buildLeafValueMap,
  buildMonthlyBuckets,
  buildPivotModel,
  buildWeeklyBucketsFromLeaf,
  computeCategoryKeys,
  editableLeafKeysForRowWeek,
  getAvailableYears,
  getCapacityStatus,
  makeSyntheticLeafKey,
  parseLeafKey,
  weekRangeLabel,
} from './utils/planner'

const INITIAL_FILE_NAME = 'Hours_03-05-26.xlsx'
const APP_LOCK_PASSWORD = '2431'
const APP_UNLOCK_SESSION_KEY = 'capacity_dashboard_unlocked'
const SALES_STORAGE_KEY = 'sales_tasks_v1'
const SALES_META_KEY = 'sales_meta_v1'
const DEFAULT_RESOURCE_WEEKLY: Record<string, number> = {
  Fabrication: 1440,
  Assembly: 80,
  Processing: 280,
  Paint: 60,
  Shipping: 200,
}

function uniqueSorted(values: string[]): string[] {
  return [...new Set(values)].sort((a, b) => a.localeCompare(b))
}

function App() {
  const reportTabParam =
    typeof window !== 'undefined' ? new URLSearchParams(window.location.search).get('reportTab') : null
  const initialReportTab: ReportTab =
    reportTabParam === 'weekly' ||
    reportTabParam === 'monthly' ||
    reportTabParam === 'summary' ||
    reportTabParam === 'sales' ||
    reportTabParam === 'combined' ||
    reportTabParam === 'sales-monthly' ||
    reportTabParam === 'combined-monthly' ||
    reportTabParam === 'executive'
      ? reportTabParam
      : 'snapshot'
  const [tasks, setTasks] = useState<TaskRow[]>([])
  const [fileName, setFileName] = useState(INITIAL_FILE_NAME)
  const [salesTasks, setSalesTasks] = useState<TaskRow[]>([])
  const [salesFileName, setSalesFileName] = useState('2026 Sales Production Report.xlsx')
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState('')
  const [pivotRowGrouping, setPivotRowGrouping] = useState<PivotRowGrouping>('project')
  const [chartGroupBy, setChartGroupBy] = useState<ChartGroupBy>('project')
  const [selectedWeekendDates, setSelectedWeekendDates] = useState<Set<string>>(new Set())
  const [manualOverrides, setManualOverrides] = useState<Record<string, number>>({})
  const [salesManualOverrides, setSalesManualOverrides] = useState<Record<string, number>>({})
  const [isPivotCollapsed, setIsPivotCollapsed] = useState(true)
  const [isSalesPivotCollapsed, setIsSalesPivotCollapsed] = useState(true)
  const [selectedProjects, setSelectedProjects] = useState<Set<string>>(new Set())
  const [salesSelectedProjects, setSalesSelectedProjects] = useState<Set<string>>(new Set())
  const [resourceWeeklyCapacities, setResourceWeeklyCapacities] = useState<Record<string, number>>({})
  const [enabledResources, setEnabledResources] = useState<Record<string, boolean>>({})
  const [salesEnabledResources, setSalesEnabledResources] = useState<Record<string, boolean>>({})
  const [weekendExtraByResource, setWeekendExtraByResource] = useState<Record<string, number>>({})
  const [holidayDates, setHolidayDates] = useState<Array<{ iso: string; name: string }>>([])
  const [projectsInitialized, setProjectsInitialized] = useState(false)
  const [salesProjectsInitialized, setSalesProjectsInitialized] = useState(false)
  const [pivotWeekWindowSize, setPivotWeekWindowSize] = useState(12)
  const [pivotWeekStartIndex, setPivotWeekStartIndex] = useState(0)
  const [collapseResetToken, setCollapseResetToken] = useState(0)
  const [salesCollapseResetToken, setSalesCollapseResetToken] = useState(0)
  const [salesPivotWeekWindowSize, setSalesPivotWeekWindowSize] = useState(12)
  const [salesPivotWeekStartIndex, setSalesPivotWeekStartIndex] = useState(0)
  const [isUnlocked, setIsUnlocked] = useState<boolean>(() => {
    if (typeof window === 'undefined') {
      return false
    }
    return window.sessionStorage.getItem(APP_UNLOCK_SESSION_KEY) === 'true'
  })
  const [passwordInput, setPasswordInput] = useState('')
  const [passwordError, setPasswordError] = useState('')
  const [hoveredProject, setHoveredProject] = useState<string | null>(null)

  function persistSalesState(
    tasksToStore: TaskRow[],
    file: string,
    overrides: Record<string, number>,
    selected: Set<string>,
    enabled: Record<string, boolean>,
  ): void {
    if (typeof window === 'undefined') return
    const plainTasks = tasksToStore.map((t) => ({
      ...t,
      start: t.start.toISOString(),
      finish: t.finish.toISOString(),
    }))
    window.sessionStorage.setItem(SALES_STORAGE_KEY, JSON.stringify(plainTasks))
    window.sessionStorage.setItem(
      SALES_META_KEY,
      JSON.stringify({
        file,
        overrides,
        selected: [...selected],
        enabled,
      }),
    )
  }

  const [filters, setFilters] = useState<AppFilters>({
    dateFrom: '',
    dateTo: '',
    year: '',
    resources: [],
  })

  useEffect(() => {
    async function loadInitialWorkbook(): Promise<void> {
      setIsLoading(true)
      setError('')

      try {
        const response = await fetch(`/${INITIAL_FILE_NAME}`)
        if (!response.ok) {
          throw new Error(`Unable to fetch ${INITIAL_FILE_NAME}`)
        }

        const workbookData = await response.arrayBuffer()
        const parsedTasks = parseSpreadsheet(workbookData)
        setTasks(parsedTasks)
      } catch {
        setError('Could not load the default workbook. Upload a .xlsx file to continue.')
      } finally {
        setIsLoading(false)
      }
    }

    void loadInitialWorkbook()
  }, [])

  useEffect(() => {
    if (typeof window === 'undefined') return
    try {
      const savedTasksRaw = window.sessionStorage.getItem(SALES_STORAGE_KEY)
      const savedMetaRaw = window.sessionStorage.getItem(SALES_META_KEY)
      if (!savedTasksRaw || !savedMetaRaw) {
        return
      }
      const parsedTasks: TaskRow[] = JSON.parse(savedTasksRaw).map((t: any) => ({
        ...t,
        start: parseISO(t.start),
        finish: parseISO(t.finish),
      }))
      const meta = JSON.parse(savedMetaRaw)
      setSalesTasks(parsedTasks)
      setSalesFileName(meta.file ?? salesFileName)
      setSalesManualOverrides(meta.overrides ?? {})
      // Always start with all projects visible on refresh; do not restore hidden selections.
      setSalesSelectedProjects(new Set())
      setSalesEnabledResources(meta.enabled ?? {})
      setSalesProjectsInitialized(false)
    } catch {
      // Ignore session restore errors and start fresh
    }
  }, [])

  const resources = useMemo(() => uniqueSorted(tasks.map((task) => task.resourceName)), [tasks])
  const salesResources = useMemo(() => uniqueSorted(salesTasks.map((task) => task.resourceName)), [salesTasks])
  const years = useMemo(() => getAvailableYears(tasks), [tasks])

  useEffect(() => {
    if (years.length === 0) {
      return
    }

    setFilters((current) => {
      if (current.year) {
        return current
      }
      return { ...current, year: years[years.length - 1] }
    })
  }, [years])

  useEffect(() => {
    if (tasks.length === 0) {
      setSelectedProjects(new Set())
      setProjectsInitialized(false)
      return
    }
  }, [tasks])

  useEffect(() => {
    if (resources.length === 0) {
      setEnabledResources({})
      setResourceWeeklyCapacities({})
      return
    }

    setEnabledResources((current) => {
      const next: Record<string, boolean> = {}
      for (const resource of resources) {
        next[resource] = current[resource] !== false
      }
      return next
    })
    setWeekendExtraByResource((current) => {
      const next: Record<string, number> = {}
      for (const resource of resources) {
        const candidate = current[resource]
        next[resource] = Number.isFinite(candidate) ? candidate : 0
      }
      return next
    })
  }, [resources])

  useEffect(() => {
    if (salesResources.length === 0) {
      setSalesEnabledResources({})
      return
    }

    setSalesEnabledResources((current) => {
      const next: Record<string, boolean> = {}
      for (const resource of salesResources) {
        next[resource] = current[resource] !== false
      }
      return next
    })
  }, [salesResources])

  useEffect(() => {
    if (resources.length === 0) {
      setResourceWeeklyCapacities({})
      return
    }

    setResourceWeeklyCapacities((current) => {
      const next: Record<string, number> = {}
      for (const resource of resources) {
        const candidate = current[resource]
        if (Number.isFinite(candidate)) {
          next[resource] = candidate
          continue
        }
        const mapped = DEFAULT_RESOURCE_WEEKLY[resource]
        if (Number.isFinite(mapped)) {
          next[resource] = mapped
          continue
        }
        next[resource] = 0
      }
      return next
    })
  }, [resources])

  const enabledResourceList = useMemo(
    () => resources.filter((resource) => enabledResources[resource] !== false),
    [resources, enabledResources],
  )
  const enabledResourceSet = useMemo(() => new Set(enabledResourceList), [enabledResourceList])
  const salesEnabledResourceList = useMemo(
    () => salesResources.filter((resource) => salesEnabledResources[resource] !== false),
    [salesResources, salesEnabledResources],
  )
  const salesEnabledResourceSet = useMemo(() => new Set(salesEnabledResourceList), [salesEnabledResourceList])

  const baseLayer = useMemo(
    () => buildBaseLeafCells(tasks, filters, selectedWeekendDates, enabledResourceSet),
    [tasks, filters, selectedWeekendDates, enabledResourceSet],
  )
  const salesBaseLayer = useMemo(
    () => buildBaseLeafCells(salesTasks, filters, selectedWeekendDates, salesEnabledResourceSet),
    [salesTasks, filters, selectedWeekendDates, salesEnabledResourceSet],
  )
  const availableProjects = useMemo(() => {
    const totals = new Map<string, number>()
    baseLayer.leafCells.forEach((cell) => {
      totals.set(cell.project, (totals.get(cell.project) ?? 0) + cell.hours)
    })
    return [...totals.entries()]
      .filter(([, hours]) => hours > 0)
      .map(([project]) => project)
      .sort((a, b) => a.localeCompare(b))
  }, [baseLayer.leafCells])
  const salesAvailableProjects = useMemo(() => {
    const totals = new Map<string, number>()
    salesBaseLayer.leafCells.forEach((cell) => {
      totals.set(cell.project, (totals.get(cell.project) ?? 0) + cell.hours)
    })
    return [...totals.entries()]
      .filter(([, hours]) => hours > 0)
      .map(([project]) => project)
      .sort((a, b) => a.localeCompare(b))
  }, [salesBaseLayer.leafCells])
  const combinedProjects = useMemo(
    () => [...availableProjects, ...salesAvailableProjects.map((p) => `Sales - ${p}`)],
    [availableProjects, salesAvailableProjects],
  )

  useEffect(() => {
    if (availableProjects.length === 0) {
      setSelectedProjects(new Set())
      setProjectsInitialized(false)
      return
    }

    if (!projectsInitialized) {
      setSelectedProjects(new Set(availableProjects))
      setProjectsInitialized(true)
      return
    }

    setSelectedProjects((current) => new Set([...current].filter((project) => availableProjects.includes(project))))
  }, [availableProjects, projectsInitialized])

  useEffect(() => {
    if (salesAvailableProjects.length === 0) {
      setSalesSelectedProjects(new Set())
      setSalesProjectsInitialized(false)
      return
    }

    if (!salesProjectsInitialized) {
      setSalesSelectedProjects(new Set(salesAvailableProjects))
      setSalesProjectsInitialized(true)
      return
    }

    setSalesSelectedProjects((current) => new Set([...current].filter((project) => salesAvailableProjects.includes(project))))
  }, [salesAvailableProjects, salesProjectsInitialized])
  const selectedProjectsForCalc = useMemo(() => selectedProjects, [selectedProjects])
  const salesSelectedProjectsForCalc = useMemo(() => salesSelectedProjects, [salesSelectedProjects])
  const combinedSelectedProjects = useMemo(() => {
    const ops = [...selectedProjects].filter((p) => availableProjects.includes(p))
    const sales = [...salesSelectedProjects].map((p) => `Sales - ${p}`)
    return new Set<string>([...ops, ...sales])
  }, [selectedProjects, salesSelectedProjects, availableProjects])

  const { baseByKey, finalByKey } = useMemo(
    () => buildLeafValueMap(baseLayer.leafCells, manualOverrides, selectedProjectsForCalc),
    [baseLayer.leafCells, manualOverrides, selectedProjectsForCalc],
  )
  const { baseByKey: salesBaseByKey, finalByKey: salesFinalByKey } = useMemo(
    () => buildLeafValueMap(salesBaseLayer.leafCells, salesManualOverrides, salesSelectedProjectsForCalc),
    [salesBaseLayer.leafCells, salesManualOverrides, salesSelectedProjectsForCalc],
  )

  const weekendWeeks = useMemo(() => {
    const set = new Set<string>()
    selectedWeekendDates.forEach((iso) => {
      const d = parseISO(iso)
      const weekStart = startOfWeek(d, { weekStartsOn: 1 })
      set.add(format(weekStart, 'yyyy-MM-dd'))
    })
    return set
  }, [selectedWeekendDates])

  useEffect(() => {
    if (!filters.year) {
      setHolidayDates([])
      return
    }
    const year = Number(filters.year)
    const holidays: Array<{ iso: string; name: string }> = []
    const push = (iso: string, name: string) => holidays.push({ iso, name })
    const observedMonday = (iso: string, name: string) => {
      const d = parseISO(iso)
      const dow = d.getDay()
      if (dow === 6) {
        push(format(addDays(d, 2), 'yyyy-MM-dd'), `${name} (observed)`)
      } else if (dow === 0) {
        push(format(addDays(d, 1), 'yyyy-MM-dd'), `${name} (observed)`)
      } else {
        push(iso, name)
      }
    }

    push(`${year}-01-01`, "New Year's Day")
    // Family Day - third Monday of February
    {
      const first = parseISO(`${year}-02-01`)
      const dow = first.getDay()
      const offset = (dow === 0 ? 1 : 8 - dow) // first Monday
      const thirdMonday = offset + 14
      push(format(addDays(first, thirdMonday - 1), 'yyyy-MM-dd'), 'Family Day')
    }
    // Good Friday approximation using known 2026? use simple rule: skip accurate calc, fallback none
    // Skipped to avoid date calc complexity; add here if required.
    // Victoria Day: Monday preceding May 25
    {
      const may25 = parseISO(`${year}-05-25`)
      const dow = may25.getDay()
      const offset = dow === 1 ? 7 : dow === 0 ? 6 : dow - 1
      push(format(addDays(may25, -offset), 'yyyy-MM-dd'), 'Victoria Day')
    }
    // Canada Day observed
    observedMonday(`${year}-07-01`, 'Canada Day')
    // Saskatchewan Day first Monday August
    {
      const aug1 = parseISO(`${year}-08-01`)
      const dow = aug1.getDay()
      const offset = dow === 1 ? 0 : dow === 0 ? 1 : 8 - dow
      push(format(addDays(aug1, offset), 'yyyy-MM-dd'), 'Saskatchewan Day')
    }
    // Labour Day first Monday September
    {
      const sep1 = parseISO(`${year}-09-01`)
      const dow = sep1.getDay()
      const offset = dow === 1 ? 0 : dow === 0 ? 1 : 8 - dow
      push(format(addDays(sep1, offset), 'yyyy-MM-dd'), 'Labour Day')
    }
    // National Day for Truth and Reconciliation (federal) - include if weekday
    observedMonday(`${year}-09-30`, 'National Day for Truth and Reconciliation')
    // Thanksgiving second Monday October
    {
      const oct1 = parseISO(`${year}-10-01`)
      const dow = oct1.getDay()
      const firstMondayOffset = dow === 1 ? 0 : dow === 0 ? 1 : 8 - dow
      const secondMonday = firstMondayOffset + 7
      push(format(addDays(oct1, secondMonday), 'yyyy-MM-dd'), 'Thanksgiving')
    }
    // Remembrance Day
    observedMonday(`${year}-11-11`, 'Remembrance Day')
    // Christmas Day
    observedMonday(`${year}-12-25`, 'Christmas Day')
    // Boxing Day
    observedMonday(`${year}-12-26`, 'Boxing Day')

    setHolidayDates(holidays)
  }, [filters.year])

  const allWeekKeys = useMemo(() => {
    const set = new Set<string>()
    baseLayer.weekKeys.forEach((w) => set.add(w))
    salesBaseLayer.weekKeys.forEach((w) => set.add(w))
    return [...set].sort((a, b) => parseISO(a).getTime() - parseISO(b).getTime())
  }, [baseLayer.weekKeys, salesBaseLayer.weekKeys])

  const holidayWeeks = useMemo(() => {
    const map = new Map<string, Array<{ name: string; iso: string }>>()
    holidayDates.forEach(({ iso, name }) => {
      const d = parseISO(iso)
      const dow = d.getDay()
      if (dow === 0 || dow === 6) {
        if (!selectedWeekendDates.has(iso)) {
          return
        }
      }
      const weekStart = startOfWeek(d, { weekStartsOn: 1 })
      const key = format(weekStart, 'yyyy-MM-dd')
      const arr = map.get(key) ?? []
      arr.push({ iso, name })
      map.set(key, arr)
    })
    return map
  }, [holidayDates, selectedWeekendDates])

  const holidayDetailsByWeek = useMemo(() => {
    const rec: Record<string, Array<{ name: string; date: string }>> = {}
    holidayWeeks.forEach((arr, key) => {
      rec[key] = arr.map((entry) => ({ name: entry.name, date: entry.iso }))
    })
    return rec
  }, [holidayWeeks])

  const weekCapacities = useMemo(() => {
    const map: Record<string, number> = {}
    for (const weekIso of allWeekKeys) {
      let weeklyTotal = 0
      const holidays = holidayWeeks.get(weekIso) ?? []
      for (const resource of enabledResourceList) {
        const weekly = resourceWeeklyCapacities[resource] ?? 0
        const weekendExtra = weekendExtraByResource[resource] ?? 0
        const holidayReduction =
          holidays.length > 0 ? Math.min(weekly, (weekly / 5) * holidays.length) : 0
        const weekendApply = weekendWeeks.has(weekIso) ? weekendExtra : 0
        weeklyTotal += weekly + weekendApply - holidayReduction
      }
      map[weekIso] = weeklyTotal
    }
    return map
  }, [allWeekKeys, enabledResourceList, resourceWeeklyCapacities, weekendExtraByResource, weekendWeeks, holidayWeeks])

  const selectedWeeklyCapacity = useMemo(() => {
    if (baseLayer.weekKeys.length === 0) return 0
    return baseLayer.weekKeys.reduce((sum, week) => sum + (weekCapacities[week] ?? 0), 0) / baseLayer.weekKeys.length
  }, [baseLayer.weekKeys, weekCapacities])

  const weeklyBuckets = useMemo(
    () => buildWeeklyBucketsFromLeaf(finalByKey, baseLayer.weekKeys, weekCapacities, chartGroupBy, holidayDetailsByWeek),
    [finalByKey, baseLayer.weekKeys, weekCapacities, chartGroupBy, holidayDetailsByWeek],
  )
  const salesWeeklyBuckets = useMemo(
    () => buildWeeklyBucketsFromLeaf(salesFinalByKey, allWeekKeys, weekCapacities, chartGroupBy, holidayDetailsByWeek),
    [salesFinalByKey, allWeekKeys, weekCapacities, chartGroupBy, holidayDetailsByWeek],
  )

  const combinedWeeklyBuckets = useMemo(() => {
    const weekSet = new Set<string>()
    weeklyBuckets.forEach((b) => weekSet.add(b.weekStartIso))
    salesWeeklyBuckets.forEach((b) => weekSet.add(b.weekStartIso))
    const sortedWeeks = [...weekSet].sort((a, b) => parseISO(a).getTime() - parseISO(b).getTime())

    const opsMap = new Map(weeklyBuckets.map((b) => [b.weekStartIso, b]))
    const salesMap = new Map(salesWeeklyBuckets.map((b) => [b.weekStartIso, b]))

      return sortedWeeks.map((weekStartIso) => {
        const ops = opsMap.get(weekStartIso)
        const sales = salesMap.get(weekStartIso)
        const groups: Record<string, number> = {}

        if (ops) {
          Object.entries(ops.groups).forEach(([key, value]) => {
            groups[key] = value
          })
        }
        if (sales) {
          Object.entries(sales.groups).forEach(([key, value]) => {
            groups[`Sales - ${key}`] = value
          })
        }

        // Use unified capacity map so the capacity line remains continuous across all weeks.
        const capacity = weekCapacities[weekStartIso] ?? ops?.capacity ?? 0
        const totalHours = Object.values(groups).reduce((sum, value) => sum + value, 0)
        const variance = totalHours - capacity
        const status = getCapacityStatus(totalHours, capacity)
        const weekEndIso = ops?.weekEndIso ?? format(addDays(parseISO(weekStartIso), 4), 'yyyy-MM-dd')
        const mergedHolidayDetails: Array<{ name: string; date: string }> = []
        const seenHolidayKeys = new Set<string>()
        ;[...(ops?.holidayDetails ?? []), ...(sales?.holidayDetails ?? [])].forEach((entry) => {
          const key = `${entry.name}|${entry.date}`
          if (seenHolidayKeys.has(key)) {
            return
          }
          seenHolidayKeys.add(key)
          mergedHolidayDetails.push(entry)
        })

        return {
          weekStartIso,
          weekEndIso,
          weekLabel: weekRangeLabel(weekStartIso),
          totalHours,
          capacity,
          variance,
          overCapacity: status === 'Over Capacity',
          status,
          groups,
          holidayDetails: mergedHolidayDetails,
        }
      })
  }, [weeklyBuckets, salesWeeklyBuckets])
  const monthlyBuckets = useMemo(() => buildMonthlyBuckets(weeklyBuckets), [weeklyBuckets])
  const salesMonthlyBuckets = useMemo(() => buildMonthlyBuckets(salesWeeklyBuckets), [salesWeeklyBuckets])
  const combinedMonthlyBuckets = useMemo(() => buildMonthlyBuckets(combinedWeeklyBuckets), [combinedWeeklyBuckets])
  const monthlyCapacityTotal = useMemo(
    () => monthlyBuckets.reduce((sum, m) => sum + m.capacity, 0),
    [monthlyBuckets],
  )

  const categoryKeys = useMemo(() => computeCategoryKeys(weeklyBuckets), [weeklyBuckets])
  const salesCategoryKeys = useMemo(() => computeCategoryKeys(salesWeeklyBuckets), [salesWeeklyBuckets])
  const combinedCategoryKeys = useMemo(() => computeCategoryKeys(combinedWeeklyBuckets), [combinedWeeklyBuckets])

  const projectTotals = useMemo(() => {
    const totals = new Map<string, number>()
    const weekSet = new Set(baseLayer.weekKeys)
    Object.entries(finalByKey).forEach(([leafKey, hours]) => {
      if (!Number.isFinite(hours)) {
        return
      }
      const { project, weekStartIso } = parseLeafKey(leafKey)
      if (!weekSet.has(weekStartIso)) {
        return
      }
      totals.set(project, (totals.get(project) ?? 0) + hours)
    })
    return totals
  }, [finalByKey, baseLayer.weekKeys])

  const topProjects = useMemo(() => {
    const entries = [...projectTotals.entries()].sort((a, b) => (b[1] ?? 0) - (a[1] ?? 0))
    const grandTotal = entries.reduce((sum, [, hours]) => sum + hours, 0)
    return entries.slice(0, 5).map(([project, hours]) => ({
      project,
      hours,
      percent: grandTotal > 0 ? (hours / grandTotal) * 100 : 0,
    }))
  }, [projectTotals])

  const executiveData = useMemo<ExecutiveData>(() => {
    const bookedYtd = weeklyBuckets.reduce((sum, w) => sum + w.totalHours, 0)
    const capacityYtd = weeklyBuckets.reduce((sum, w) => sum + w.capacity, 0)
    const utilization = capacityYtd > 0 ? bookedYtd / capacityYtd : 0
    const remaining = capacityYtd - bookedYtd
    const activeProjects = projectTotals.size
    const kpiStatus = utilization > 1 ? 'red' : utilization >= 0.9 ? 'yellow' : 'green'

    const quarterlySummary = (() => {
      const map = new Map<string, { booked: number; capacity: number }>()
      monthlyBuckets.forEach((m) => {
        const monthDate = parseISO(`${m.monthKey}-01`)
        const quarter = `Q${Math.floor(monthDate.getMonth() / 3) + 1}`
        const entry = map.get(quarter) ?? { booked: 0, capacity: 0 }
        entry.booked += m.plannedHours
        entry.capacity += m.capacity
        map.set(quarter, entry)
      })
      return [...map.entries()]
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([quarter, values]) => ({
          quarter,
          booked: values.booked,
          capacity: values.capacity,
          utilization: values.capacity > 0 ? values.booked / values.capacity : 0,
        }))
    })()

    const annual = {
      booked: bookedYtd,
      capacity: capacityYtd,
      utilization,
      status: getCapacityStatus(bookedYtd, capacityYtd),
    }

    const riskMonths = monthlyBuckets
      .filter((m) => m.plannedHours > m.capacity)
      .map((m) => ({ monthLabel: m.monthLabel, variance: m.variance }))

    const utilizationTrend = monthlyBuckets.map((m) => ({
      monthLabel: m.monthLabel,
      utilization: m.capacity > 0 ? m.plannedHours / m.capacity : 0,
    }))

    return {
      kpis: { bookedYtd, capacityYtd, utilization, remaining, activeProjects, status: kpiStatus },
      monthlyBuckets,
      quarterlySummary,
      annual,
      riskMonths,
      topProjects,
      utilizationTrend,
    }
  }, [weeklyBuckets, monthlyBuckets, projectTotals, topProjects])

  useEffect(() => {
    persistSalesState(salesTasks, salesFileName, salesManualOverrides, salesSelectedProjects, salesEnabledResources)
  }, [salesTasks, salesFileName, salesManualOverrides, salesSelectedProjects, salesEnabledResources])

  const pivotModel = useMemo(
    () => buildPivotModel(finalByKey, baseByKey, baseLayer.weekKeys, pivotRowGrouping),
    [finalByKey, baseByKey, baseLayer.weekKeys, pivotRowGrouping],
  )
  const salesPivotModel = useMemo(
    () => buildPivotModel(salesFinalByKey, salesBaseByKey, salesBaseLayer.weekKeys, pivotRowGrouping),
    [salesFinalByKey, salesBaseByKey, salesBaseLayer.weekKeys, pivotRowGrouping],
  )

  useEffect(() => {
    setPivotWeekStartIndex(0)
  }, [baseLayer.weekKeys, pivotWeekWindowSize])

  const maxPivotStartIndex = Math.max(0, baseLayer.weekKeys.length - pivotWeekWindowSize)
  const safePivotStartIndex = Math.min(pivotWeekStartIndex, maxPivotStartIndex)
  const visiblePivotWeekKeys = baseLayer.weekKeys.slice(
    safePivotStartIndex,
    safePivotStartIndex + pivotWeekWindowSize,
  )
  const pivotWeekWindowLabel =
    visiblePivotWeekKeys.length > 0
      ? `${visiblePivotWeekKeys[0]} to ${visiblePivotWeekKeys[visiblePivotWeekKeys.length - 1]}`
      : 'No weeks'

  useEffect(() => {
    setSalesPivotWeekStartIndex(0)
  }, [salesBaseLayer.weekKeys, salesPivotWeekWindowSize])

  const salesMaxPivotStartIndex = Math.max(0, salesBaseLayer.weekKeys.length - salesPivotWeekWindowSize)
  const salesSafePivotStartIndex = Math.min(salesPivotWeekStartIndex, salesMaxPivotStartIndex)
  const salesVisiblePivotWeekKeys = salesBaseLayer.weekKeys.slice(
    salesSafePivotStartIndex,
    salesSafePivotStartIndex + salesPivotWeekWindowSize,
  )
  const salesPivotWeekWindowLabel =
    salesVisiblePivotWeekKeys.length > 0
      ? `${salesVisiblePivotWeekKeys[0]} to ${salesVisiblePivotWeekKeys[salesVisiblePivotWeekKeys.length - 1]}`
      : 'No weeks'

  const totals = useMemo(() => {
    return weeklyBuckets.reduce(
      (acc, bucket) => {
        acc.hours += bucket.totalHours
        acc.capacity += bucket.capacity
        acc.variance += bucket.variance
        acc.overCount += bucket.overCapacity ? 1 : 0
        return acc
      },
      { hours: 0, capacity: 0, variance: 0, overCount: 0 },
    )
  }, [weeklyBuckets])

  const summaryMetrics = useMemo<SummaryMetric[]>(() => {
    const reportTimestamp = format(new Date(), 'yyyy-MM-dd HH:mm')
    return [
      { metric: 'File Name', value: fileName },
      { metric: 'Selected Year', value: filters.year || 'All years' },
      { metric: 'Total Forecast Hours', value: totals.hours.toFixed(2) },
      { metric: 'Total Capacity Hours', value: totals.capacity.toFixed(2) },
      { metric: 'Selected Weekly Capacity', value: selectedWeeklyCapacity.toFixed(2) },
      { metric: 'Total Monthly Capacity', value: monthlyCapacityTotal.toFixed(2) },
      { metric: 'Variance (Forecast - Capacity)', value: totals.variance.toFixed(2) },
      { metric: 'Over-Capacity Weeks', value: totals.overCount },
      { metric: 'Manual Overrides Count', value: Object.keys(manualOverrides).length },
      { metric: 'Last Updated', value: reportTimestamp },
    ]
  }, [fileName, filters.year, totals, selectedWeeklyCapacity, monthlyCapacityTotal, manualOverrides])

  const reportContext = useMemo(
    () => [
      `Year: ${filters.year || 'All years'}`,
      `Projects Selected: ${selectedProjects.size}/${availableProjects.length}`,
      `Resources Enabled: ${enabledResourceList.length}`,
      `Weeks in View: ${baseLayer.weekKeys.length}`,
    ],
    [filters.year, selectedProjects, availableProjects.length, enabledResourceList.length, baseLayer.weekKeys.length],
  )

  useEffect(() => {
    if (import.meta.env.DEV) {
      const weeklyCapacitySum = weeklyBuckets.reduce((sum, bucket) => sum + bucket.capacity, 0)
      const monthlyCapacitySum = monthlyBuckets.reduce((sum, month) => sum + month.capacity, 0)
      // Debug helper: monthly capacity should reconcile to weekly capacity totals (after day split).
      console.debug('[capacity-debug] weekly capacity sum', weeklyCapacitySum.toFixed(2))
      console.debug('[capacity-debug] monthly capacity sum', monthlyCapacitySum.toFixed(2))
    }
  }, [weeklyBuckets, monthlyBuckets])

  const taskDateSpan = useMemo(() => {
    if (tasks.length === 0) {
      return { start: '', end: '' }
    }

    let minDate = tasks[0].start
    let maxDate = tasks[0].finish

    for (const task of tasks) {
      if (task.start < minDate) {
        minDate = task.start
      }
      if (task.finish > maxDate) {
        maxDate = task.finish
      }
    }

    return {
      start: format(minDate, 'yyyy-MM-dd'),
      end: format(maxDate, 'yyyy-MM-dd'),
    }
  }, [tasks])

  const allLeafKeys = useMemo(() => {
    const keys = new Set<string>()
    baseLayer.leafCells.forEach((cell) => {
      keys.add([cell.project, cell.resource, cell.weekStartIso].join('\u0001'))
    })
    Object.keys(manualOverrides).forEach((key) => keys.add(key))
    return [...keys]
  }, [baseLayer.leafCells, manualOverrides])
  const salesLeafKeys = useMemo(() => {
    const keys = new Set<string>()
    salesBaseLayer.leafCells.forEach((cell) => {
      keys.add([cell.project, cell.resource, cell.weekStartIso].join('\u0001'))
    })
    Object.keys(salesManualOverrides).forEach((key) => keys.add(key))
    return [...keys]
  }, [salesBaseLayer.leafCells, salesManualOverrides])

  async function handleUpload(event: React.ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    setIsLoading(true)
    setError('')

    try {
      const arrayBuffer = await file.arrayBuffer()
      const parsed = parseSpreadsheet(arrayBuffer)
      setTasks(parsed)
      setFileName(file.name)
      setManualOverrides({})
      setResourceWeeklyCapacities({})
      setEnabledResources({})
      setFilters({ dateFrom: '', dateTo: '', year: '', resources: [] })
      setSelectedWeekendDates(new Set())
      setWeekendExtraByResource({})
      setProjectsInitialized(false)
      // On new workbook upload, reset panel defaults (pivot/resource collapsed; forecast tables expanded by component default).
      setIsPivotCollapsed(true)
      setCollapseResetToken((current) => current + 1)
    } catch {
      setError('Failed to parse workbook. Please upload a valid .xlsx file with Work, Start, and Finish columns.')
    } finally {
      setIsLoading(false)
      event.target.value = ''
    }
  }

  async function handleSalesUpload(event: React.ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    setIsLoading(true)
    setError('')

    try {
      const arrayBuffer = await file.arrayBuffer()
      const parsed = parseSalesSpreadsheet(arrayBuffer)
      setSalesTasks(parsed)
      setSalesFileName(file.name)
      setSalesManualOverrides({})
      setSalesEnabledResources({})
      setSalesSelectedProjects(new Set())
      setSalesProjectsInitialized(false)
      setIsSalesPivotCollapsed(true)
      setSalesCollapseResetToken((current) => current + 1)
      persistSalesState(parsed, file.name, {}, new Set(), {})
    } catch {
      setError('Failed to parse sales workbook. Please upload a valid Sales Production Report with the expected columns.')
    } finally {
      setIsLoading(false)
      event.target.value = ''
    }
  }

  function handleToggleResource(resource: string, enabled: boolean): void {
    setEnabledResources((current) => ({ ...current, [resource]: enabled }))
  }

  function handleResourceWeeklyCapacityChange(resource: string, weeklyCapacity: number): void {
    if (!Number.isFinite(weeklyCapacity) || weeklyCapacity < 0) {
      return
    }

    setResourceWeeklyCapacities((current) => ({
      ...current,
      [resource]: weeklyCapacity,
    }))
  }

  function handlePivotCellEdit(rowKey: string, weekStartIso: string, newValue: number): void {
    if (!Number.isFinite(newValue) || newValue < 0) {
      return
    }

    const matchingLeafKeys = editableLeafKeysForRowWeek(
      rowKey,
      weekStartIso,
      pivotRowGrouping,
      allLeafKeys,
      selectedProjects,
    )

    const targetLeafKeys =
      matchingLeafKeys.length > 0
        ? matchingLeafKeys
        : [makeSyntheticLeafKey(rowKey, weekStartIso, pivotRowGrouping, selectedProjects)]

    setManualOverrides((current) => {
      const next = { ...current }

      const currentValues = targetLeafKeys.map((leafKey) => {
        if (leafKey in current) {
          return current[leafKey]
        }
        return baseByKey[leafKey] ?? 0
      })

      const currentSum = currentValues.reduce((sum, value) => sum + value, 0)

      if (targetLeafKeys.length === 1) {
        next[targetLeafKeys[0]] = newValue
        return next
      }

      if (currentSum <= 0) {
        targetLeafKeys.forEach((leafKey, index) => {
          next[leafKey] = index === 0 ? newValue : 0
        })
        return next
      }

      let assigned = 0
      targetLeafKeys.forEach((leafKey, index) => {
        if (index === targetLeafKeys.length - 1) {
          next[leafKey] = Math.max(0, newValue - assigned)
          return
        }

        const source = currentValues[index]
        const allocated = (source / currentSum) * newValue
        next[leafKey] = allocated
        assigned += allocated
      })

      return next
    })
  }

  function handleSalesPivotCellEdit(rowKey: string, weekStartIso: string, newValue: number): void {
    if (!Number.isFinite(newValue) || newValue < 0) {
      return
    }

    const matchingLeafKeys = editableLeafKeysForRowWeek(
      rowKey,
      weekStartIso,
      pivotRowGrouping,
      salesLeafKeys,
      salesSelectedProjects,
    )

    const targetLeafKeys =
      matchingLeafKeys.length > 0
        ? matchingLeafKeys
        : [makeSyntheticLeafKey(rowKey, weekStartIso, pivotRowGrouping, salesSelectedProjects)]

    setSalesManualOverrides((current) => {
      const next = { ...current }

      const currentValues = targetLeafKeys.map((leafKey) => {
        if (leafKey in current) {
          return current[leafKey]
        }
        return salesBaseByKey[leafKey] ?? 0
      })

      const currentSum = currentValues.reduce((sum, value) => sum + value, 0)

      if (targetLeafKeys.length === 1) {
        next[targetLeafKeys[0]] = newValue
        return next
      }

      if (currentSum <= 0) {
        targetLeafKeys.forEach((leafKey, index) => {
          next[leafKey] = index === 0 ? newValue : 0
        })
        return next
      }

      let assigned = 0
      targetLeafKeys.forEach((leafKey, index) => {
        if (index === targetLeafKeys.length - 1) {
          next[leafKey] = Math.max(0, newValue - assigned)
          return
        }

        const source = currentValues[index]
        const allocated = (source / currentSum) * newValue
        next[leafKey] = allocated
        assigned += allocated
      })

      return next
    })
  }

  async function exportReportExcel(): Promise<void> {
    const dateStamp = format(new Date(), 'yyyy-MM-dd')
    const fileName = `capacity-report-${dateStamp}.xlsx`

    const exportedWithChart = await exportReportWorkbookWithChartApi({
      weeklyBuckets,
      monthlyBuckets,
      chartCategoryKeys: categoryKeys,
      summary: summaryMetrics,
      fileName,
    })

    if (exportedWithChart) {
      return
    }

    window.alert(
      'Embedded chart export is unavailable. Start the local export API (python backend/export_api.py), then export again.',
    )

    // Optional fallback file (without embedded Excel chart object) so export still works.
    exportReportWorkbook({
      weeklyBuckets,
      monthlyBuckets,
      chartCategoryKeys: categoryKeys,
      summary: summaryMetrics,
      fileName,
    })
  }

  function resetFilters(): void {
    setFilters((current) => ({ ...current, dateFrom: '', dateTo: '', year: current.year, resources: [] }))
    setSelectedProjects(new Set(availableProjects))
    setSalesSelectedProjects(new Set(salesAvailableProjects))
    setSelectedWeekendDates(new Set())
    setEnabledResources(() => {
      const next: Record<string, boolean> = {}
      resources.forEach((resource) => {
        next[resource] = true
      })
      return next
    })
    setWeekendExtraByResource({})
  }

  function resetManualEdits(): void {
    setManualOverrides({})
  }
  function resetSalesManualEdits(): void {
    setSalesManualOverrides({})
  }

  function handleToggleProject(project: string): void {
    setSelectedProjects((current) => {
      const next = new Set(current)
      if (next.has(project)) {
        next.delete(project)
      } else {
        next.add(project)
      }
      return next
    })
  }
  function handleToggleSalesProject(project: string): void {
    setSalesSelectedProjects((current) => {
      const next = new Set(current)
      if (next.has(project)) {
        next.delete(project)
      } else {
        next.add(project)
      }
      return next
    })
  }
  function handleToggleCombinedProject(project: string): void {
    if (project.startsWith('Sales - ')) {
      const raw = project.replace(/^Sales - /, '')
      handleToggleSalesProject(raw)
      return
    }
    handleToggleProject(project)
  }

  const overCapacityWeeks = useMemo(
    () => new Set(weeklyBuckets.filter((bucket) => bucket.overCapacity).map((bucket) => bucket.weekStartIso)),
    [weeklyBuckets],
  )
  const salesOverCapacityWeeks = useMemo(
    () => new Set(salesWeeklyBuckets.filter((bucket) => bucket.overCapacity).map((bucket) => bucket.weekStartIso)),
    [salesWeeklyBuckets],
  )

  const allResourcesVisible = resources.length > 0

  function handleUnlock(event: React.FormEvent<HTMLFormElement>): void {
    event.preventDefault()
    if (passwordInput === APP_LOCK_PASSWORD) {
      setIsUnlocked(true)
      setPasswordError('')
      setPasswordInput('')
      window.sessionStorage.setItem(APP_UNLOCK_SESSION_KEY, 'true')
      return
    }
    setPasswordError('Incorrect password. Please try again.')
  }

  function handleLock(): void {
    setIsUnlocked(false)
    setPasswordInput('')
    setPasswordError('')
    window.sessionStorage.removeItem(APP_UNLOCK_SESSION_KEY)
  }

  if (!isUnlocked) {
    return (
      <div className="lock-screen">
        <section className="panel lock-card">
          <h1>Capacity Dashboard Locked</h1>
          <p>Enter the access password to open the planning dashboard.</p>
          <form className="lock-form" onSubmit={handleUnlock}>
            <label htmlFor="dashboard-password">Password</label>
            <input
              id="dashboard-password"
              type="password"
              value={passwordInput}
              onChange={(event) => setPasswordInput(event.target.value)}
              autoFocus
              autoComplete="current-password"
            />
            {passwordError && <p className="lock-error">{passwordError}</p>}
            <button type="submit">Unlock Dashboard</button>
          </form>
        </section>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <header className="panel control-panel">
        <div className="title-bar">
          <div className="title-stack">
            <h1>Production Capacity Planning Dashboard</h1>
            <p className="subtitle">
              Import forecast data, edit weekly hours in the planning pivot, and watch chart/table/capacity metrics update
              instantly.
            </p>
          </div>
          <div className="title-actions">
            <label className="upload-inline">
              Upload or Replace Workbook (.xlsx)
              <input type="file" accept=".xlsx" onChange={handleUpload} />
            </label>
            <label className="upload-inline">
              Upload Sales Workbook (.xlsx)
              <input type="file" accept=".xlsx" onChange={handleSalesUpload} />
            </label>
            <button type="button" className="ghost-btn lock-btn" onClick={handleLock}>
              Lock
            </button>
          </div>
        </div>

        <div className="controls-grid">

          <label>
            Planning Rows
            <select
              value={pivotRowGrouping}
              onChange={(event) => setPivotRowGrouping(event.target.value as PivotRowGrouping)}
            >
              <option value="project">Project</option>
              <option value="resource">Resource</option>
            </select>
          </label>

          <label>
            Chart Stacking
            <select value={chartGroupBy} onChange={(event) => setChartGroupBy(event.target.value as ChartGroupBy)}>
              <option value="project">Project</option>
              <option value="resource">Resource</option>
            </select>
          </label>

          <label>
            Year
            <select
              value={filters.year}
              onChange={(event) => setFilters((current) => ({ ...current, year: event.target.value }))}
            >
              <option value="">All years</option>
              {years.map((year) => (
                <option key={year} value={year}>
                  {year}
                </option>
              ))}
            </select>
          </label>

          <label>
            Working Weekends
            <input
              type="date"
              value=""
              onChange={(event) => {
                const iso = event.target.value
                if (!iso) return
                if (selectedWeekendDates.has(iso)) {
                  setSelectedWeekendDates((current) => {
                    const next = new Set(current)
                    next.delete(iso)
                    return next
                  })
                } else {
                  setSelectedWeekendDates((current) => {
                    const next = new Set(current)
                    next.add(iso)
                    return next
                  })
                }
                event.target.value = ''
              }}
              min={filters.year ? `${filters.year}-01-01` : undefined}
              max={filters.year ? `${filters.year}-12-31` : undefined}
            />
            <div className="weekend-pills">
              {[...selectedWeekendDates]
                .filter((iso) => !filters.year || iso.startsWith(filters.year))
                .sort()
                .map((iso) => (
                  <button
                    key={iso}
                    type="button"
                    className="chip-toggle chip-on"
                    onClick={() =>
                      setSelectedWeekendDates((current) => {
                        const next = new Set(current)
                        next.delete(iso)
                        return next
                      })
                    }
                  >
                    {iso}
                  </button>
                ))}
            </div>
          </label>

          <label>
            Week Range Start (Monday)
            <input
              type="date"
              value={filters.dateFrom}
              onChange={(event) => setFilters((current) => ({ ...current, dateFrom: event.target.value }))}
            />
          </label>

          <label>
            Week Range End (Monday)
            <input
              type="date"
              value={filters.dateTo}
              onChange={(event) => setFilters((current) => ({ ...current, dateTo: event.target.value }))}
            />
          </label>
        </div>

        <div className="meta-row">
          <div>
            <strong>File:</strong> {fileName}
          </div>
          <div>
            <strong>Sales File:</strong> {salesFileName}
          </div>
          <div>
            <strong>Parsed Rows:</strong> {tasks.length}
          </div>
          <div>
            <strong>Weeks in View:</strong> {baseLayer.weekKeys.length}
          </div>
          <div>
            <strong>Enabled Resources:</strong> {enabledResourceList.length}
          </div>
          <div>
            <strong>Data Date Span:</strong>{' '}
            {taskDateSpan.start && taskDateSpan.end ? `${taskDateSpan.start} to ${taskDateSpan.end}` : 'N/A'}
          </div>
          <button type="button" className="ghost-btn" onClick={resetFilters}>
            Reset Filters
          </button>
          <button type="button" onClick={() => void exportReportExcel()}>
            Export Report Excel
          </button>
        </div>
      </header>

      {!isLoading && !error && allResourcesVisible && (
        <ResourceCapacityTable
          key={`resource-capacity-${collapseResetToken}`}
          resources={resources}
          enabledResources={enabledResources}
          weeklyCapacitiesByResource={resourceWeeklyCapacities}
          onWeeklyCapacityChange={handleResourceWeeklyCapacityChange}
          onToggleResource={handleToggleResource}
          weekendExtraByResource={weekendExtraByResource}
          onWeekendExtraChange={(resource, hours) =>
            setWeekendExtraByResource((current) => ({
              ...current,
              [resource]: Number.isFinite(hours) && hours >= 0 ? hours : 0,
            }))
          }
        />
      )}

      {isLoading && <div className="panel status">Loading workbook...</div>}
      {!isLoading && error && <div className="panel status error">{error}</div>}
      {!isLoading && !error && weeklyBuckets.length === 0 && (
        <div className="panel status">No weekly forecast buckets match current filter and project toggle settings.</div>
      )}

      {!isLoading && !error && (weeklyBuckets.length > 0 || salesWeeklyBuckets.length > 0) && (
        <>
          {weeklyBuckets.length > 0 && (
            <PivotPlanningTable
              model={pivotModel}
              rowGrouping={pivotRowGrouping}
              overCapacityWeeks={overCapacityWeeks}
              visibleWeekKeys={visiblePivotWeekKeys}
              weekWindowLabel={pivotWeekWindowLabel}
              canPageBack={safePivotStartIndex > 0}
              canPageForward={safePivotStartIndex + pivotWeekWindowSize < baseLayer.weekKeys.length}
              onPageBack={() => setPivotWeekStartIndex((current) => Math.max(0, current - pivotWeekWindowSize))}
              onPageForward={() =>
                setPivotWeekStartIndex((current) => Math.min(maxPivotStartIndex, current + pivotWeekWindowSize))
              }
              weekWindowSize={pivotWeekWindowSize}
              onWeekWindowSizeChange={(size) => {
                if (!Number.isFinite(size) || size <= 0) {
                  return
                }
                setPivotWeekWindowSize(size)
              }}
              isCollapsed={isPivotCollapsed}
              onToggleCollapsed={() => setIsPivotCollapsed((current) => !current)}
              onEditCell={handlePivotCellEdit}
              onResetEdits={resetManualEdits}
            />
          )}
          {salesWeeklyBuckets.length > 0 && (
            <PivotPlanningTable
              key={`sales-pivot-${salesCollapseResetToken}`}
              model={salesPivotModel}
              rowGrouping={pivotRowGrouping}
              overCapacityWeeks={salesOverCapacityWeeks}
              visibleWeekKeys={salesVisiblePivotWeekKeys}
              weekWindowLabel={salesPivotWeekWindowLabel}
              canPageBack={salesSafePivotStartIndex > 0}
              canPageForward={salesSafePivotStartIndex + salesPivotWeekWindowSize < salesBaseLayer.weekKeys.length}
              onPageBack={() => setSalesPivotWeekStartIndex((current) => Math.max(0, current - salesPivotWeekWindowSize))}
              onPageForward={() =>
                setSalesPivotWeekStartIndex((current) =>
                  Math.min(salesMaxPivotStartIndex, current + salesPivotWeekWindowSize),
                )
              }
              weekWindowSize={salesPivotWeekWindowSize}
              onWeekWindowSizeChange={(size) => {
                if (!Number.isFinite(size) || size <= 0) {
                  return
                }
                setSalesPivotWeekWindowSize(size)
              }}
              isCollapsed={isSalesPivotCollapsed}
              onToggleCollapsed={() => setIsSalesPivotCollapsed((current) => !current)}
              onEditCell={handleSalesPivotCellEdit}
              onResetEdits={resetSalesManualEdits}
              title="Sales Pivot Planning"
              subtitle="Editable sales forecast planning grid using Sales Production Report data."
            />
          )}
          <section className="panel summary-panel">
            <div className="section-header">
              <h2>Summary</h2>
              <p>All metrics below are driven by the final adjusted planning dataset.</p>
            </div>

            <div className="summary-grid">
              <div>
                <span>Total Forecast Hours</span>
                <strong>{totals.hours.toFixed(2)}</strong>
              </div>
              <div>
                <span>Total Capacity Hours</span>
                <strong>{totals.capacity.toFixed(2)}</strong>
              </div>
              <div>
                <span>Selected Weekly Capacity</span>
                <strong>{selectedWeeklyCapacity.toFixed(2)}</strong>
              </div>
              <div>
                <span>Total Monthly Capacity (visible months)</span>
                <strong>{monthlyCapacityTotal.toLocaleString()}</strong>
              </div>
              <div>
                <span>Variance (Forecast - Capacity)</span>
                <strong className={totals.variance < 0 ? 'negative' : 'warning'}>
                  {totals.variance.toFixed(2)}
                </strong>
              </div>
              <div>
                <span>Over-Capacity Weeks</span>
                <strong>{totals.overCount}</strong>
              </div>
              <div>
                <span>Manual Overrides</span>
                <strong>{Object.keys(manualOverrides).length}</strong>
              </div>
            </div>
          </section>
          <ReportWorkspace
            key={`report-workspace-${collapseResetToken}-${salesCollapseResetToken}`}
            weeklyBuckets={weeklyBuckets}
            combinedWeeklyBuckets={combinedWeeklyBuckets}
            salesWeeklyBuckets={salesWeeklyBuckets}
            salesMonthlyBuckets={salesMonthlyBuckets}
            combinedMonthlyBuckets={combinedMonthlyBuckets}
            monthlyBuckets={monthlyBuckets}
            categoryKeys={categoryKeys}
            combinedCategoryKeys={combinedCategoryKeys}
            salesCategoryKeys={salesCategoryKeys}
            projects={availableProjects}
            combinedProjects={combinedProjects}
            salesProjects={salesAvailableProjects}
            selectedProjects={selectedProjects}
            selectedCombinedProjects={combinedSelectedProjects}
            selectedSalesProjects={salesSelectedProjects}
            onToggleProject={handleToggleProject}
            onToggleCombinedProject={handleToggleCombinedProject}
            onToggleSalesProject={handleToggleSalesProject}
            hoveredProject={hoveredProject}
            onHoverProject={setHoveredProject}
            summaryMetrics={summaryMetrics}
            reportContext={reportContext}
            initialTab={initialReportTab}
            executiveData={executiveData}
          />
        </>
      )}

      {!allResourcesVisible && !isLoading && (
        <div className="panel status">No resources available in the current data scope.</div>
      )}
    </div>
  )
}

export default App
