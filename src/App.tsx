import { useEffect, useMemo, useState } from 'react'
import { format, parseISO, startOfWeek } from 'date-fns'
import { PivotPlanningTable } from './components/PivotPlanningTable'
import { ReportSnapshot } from './components/ReportSnapshot'
import { ResourceCapacityTable } from './components/ResourceCapacityTable'
import type { AppFilters, ChartGroupBy, PivotRowGrouping, TaskRow } from './types'
import { parseSpreadsheet } from './utils/excel'
import { exportReportWorkbook } from './utils/reportExport'
import {
  buildBaseLeafCells,
  buildLeafValueMap,
  buildMonthlyBuckets,
  buildPivotModel,
  buildWeeklyBucketsFromLeaf,
  computeCategoryKeys,
  editableLeafKeysForRowWeek,
  getAvailableYears,
  makeSyntheticLeafKey,
} from './utils/planner'

const INITIAL_FILE_NAME = 'Hours_03-05-26.xlsx'
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
  const [tasks, setTasks] = useState<TaskRow[]>([])
  const [fileName, setFileName] = useState(INITIAL_FILE_NAME)
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState('')
  const [pivotRowGrouping, setPivotRowGrouping] = useState<PivotRowGrouping>('project')
  const [chartGroupBy, setChartGroupBy] = useState<ChartGroupBy>('project')
  const [selectedWeekendDates, setSelectedWeekendDates] = useState<Set<string>>(new Set())
  const [manualOverrides, setManualOverrides] = useState<Record<string, number>>({})
  const [isPivotCollapsed, setIsPivotCollapsed] = useState(false)
  const [selectedProjects, setSelectedProjects] = useState<Set<string>>(new Set())
  const [resourceWeeklyCapacities, setResourceWeeklyCapacities] = useState<Record<string, number>>({})
  const [enabledResources, setEnabledResources] = useState<Record<string, boolean>>({})
  const [weekendExtraByResource, setWeekendExtraByResource] = useState<Record<string, number>>({})
  const [projectsInitialized, setProjectsInitialized] = useState(false)
  const [pivotWeekWindowSize, setPivotWeekWindowSize] = useState(12)
  const [pivotWeekStartIndex, setPivotWeekStartIndex] = useState(0)

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

  const resources = useMemo(() => uniqueSorted(tasks.map((task) => task.resourceName)), [tasks])
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

  const baseLayer = useMemo(
    () => buildBaseLeafCells(tasks, filters, selectedWeekendDates, enabledResourceSet),
    [tasks, filters, selectedWeekendDates, enabledResourceSet],
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
  const selectedProjectsForCalc = useMemo(() => selectedProjects, [selectedProjects])

  const { baseByKey, finalByKey } = useMemo(
    () => buildLeafValueMap(baseLayer.leafCells, manualOverrides, selectedProjectsForCalc),
    [baseLayer.leafCells, manualOverrides, selectedProjectsForCalc],
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

  const weekCapacities = useMemo(() => {
    const map: Record<string, number> = {}
    for (const weekIso of baseLayer.weekKeys) {
      let weeklyTotal = 0
      let weekendTotal = 0
      for (const resource of enabledResourceList) {
        const weekly = resourceWeeklyCapacities[resource] ?? 0
        const weekendExtra = weekendExtraByResource[resource] ?? 0
        weeklyTotal += weekly
        weekendTotal += weekendExtra
      }
      map[weekIso] = weekendWeeks.has(weekIso) ? weeklyTotal + weekendTotal : weeklyTotal
    }
    return map
  }, [baseLayer.weekKeys, enabledResourceList, resourceWeeklyCapacities, weekendExtraByResource, weekendWeeks])

  const selectedWeeklyCapacity = useMemo(() => {
    if (baseLayer.weekKeys.length === 0) return 0
    return baseLayer.weekKeys.reduce((sum, week) => sum + (weekCapacities[week] ?? 0), 0) / baseLayer.weekKeys.length
  }, [baseLayer.weekKeys, weekCapacities])

  const weeklyBuckets = useMemo(
    () => buildWeeklyBucketsFromLeaf(finalByKey, baseLayer.weekKeys, weekCapacities, chartGroupBy),
    [finalByKey, baseLayer.weekKeys, weekCapacities, chartGroupBy],
  )

  const monthlyBuckets = useMemo(() => buildMonthlyBuckets(weeklyBuckets), [weeklyBuckets])
  const monthlyCapacityTotal = useMemo(
    () => monthlyBuckets.reduce((sum, m) => sum + m.capacity, 0),
    [monthlyBuckets],
  )

  const categoryKeys = useMemo(() => computeCategoryKeys(weeklyBuckets), [weeklyBuckets])

  const pivotModel = useMemo(
    () => buildPivotModel(finalByKey, baseByKey, baseLayer.weekKeys, pivotRowGrouping),
    [finalByKey, baseByKey, baseLayer.weekKeys, pivotRowGrouping],
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
    } catch {
      setError('Failed to parse workbook. Please upload a valid .xlsx file with Work, Start, and Finish columns.')
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

  function exportReportExcel(): void {
    const dateStamp = format(new Date(), 'yyyy-MM-dd')
    const timestamp = format(new Date(), 'yyyy-MM-dd HH:mm')
    exportReportWorkbook({
      weeklyBuckets,
      monthlyBuckets,
      summary: [
        { metric: 'File Name', value: fileName },
        { metric: 'Selected Year', value: filters.year || 'All years' },
        { metric: 'Total Forecast Hours', value: totals.hours },
        { metric: 'Total Capacity Hours', value: totals.capacity },
        { metric: 'Selected Weekly Capacity', value: selectedWeeklyCapacity },
        { metric: 'Total Monthly Capacity', value: monthlyCapacityTotal },
        { metric: 'Variance (Forecast - Capacity)', value: totals.variance },
        { metric: 'Over-Capacity Weeks', value: totals.overCount },
        { metric: 'Manual Overrides Count', value: Object.keys(manualOverrides).length },
        { metric: 'Export Timestamp', value: timestamp },
      ],
      fileName: `capacity-report-${dateStamp}.xlsx`,
    })
  }

  function resetFilters(): void {
    setFilters((current) => ({ ...current, dateFrom: '', dateTo: '', year: current.year, resources: [] }))
    setSelectedProjects(new Set(availableProjects))
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

  const overCapacityWeeks = useMemo(
    () => new Set(weeklyBuckets.filter((bucket) => bucket.overCapacity).map((bucket) => bucket.weekStartIso)),
    [weeklyBuckets],
  )

  const allResourcesVisible = resources.length > 0

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
          <label className="upload-inline">
            Upload or Replace Workbook (.xlsx)
            <input type="file" accept=".xlsx" onChange={handleUpload} />
          </label>
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
          <button type="button" onClick={exportReportExcel}>
            Export Report Excel
          </button>
        </div>
      </header>

      {!isLoading && !error && weeklyBuckets.length > 0 && (
        <ReportSnapshot
          weeklyBuckets={weeklyBuckets}
          monthlyBuckets={monthlyBuckets}
          categoryKeys={categoryKeys}
          projects={availableProjects}
          selectedProjects={selectedProjects}
          onToggleProject={(project) =>
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
            <strong className={totals.variance > 0 ? 'negative' : 'positive'}>{totals.variance.toFixed(2)}</strong>
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

      {!isLoading && !error && allResourcesVisible && (
        <ResourceCapacityTable
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

      {!isLoading && !error && weeklyBuckets.length > 0 && (
        <>
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
        </>
      )}

      {!allResourcesVisible && !isLoading && (
        <div className="panel status">No resources available in the current data scope.</div>
      )}
    </div>
  )
}

export default App
