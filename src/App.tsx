import { useEffect, useMemo, useState } from 'react'
import { format } from 'date-fns'
import { ForecastChart } from './components/ForecastChart'
import { ForecastTable } from './components/ForecastTable'
import { MonthlyForecastTable } from './components/MonthlyForecastTable'
import { MultiSelectProjects } from './components/MultiSelectProjects'
import { PivotPlanningTable } from './components/PivotPlanningTable'
import { ResourceCapacityTable } from './components/ResourceCapacityTable'
import type { AppFilters, ChartGroupBy, PivotRowGrouping, TaskRow } from './types'
import { downloadCsv, weeklyBucketsToCsv } from './utils/csv'
import { parseSpreadsheet } from './utils/excel'
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

const WEEKS_PER_MONTH = 52 / 12
const INITIAL_FILE_NAME = 'Hours_03-05-26.xlsx'
const DEFAULT_RESOURCE_WEEKLY: Record<string, number> = {
  Fabrication: 1400,
  Assembly: 40,
  Processing: 280,
  Paint: 60,
  Shipping: 200,
}

function uniqueSorted(values: string[]): string[] {
  return [...new Set(values)].sort((a, b) => a.localeCompare(b))
}

function monthlyFromWeekly(weekly: number): number {
  return weekly * WEEKS_PER_MONTH
}

function App() {
  const [tasks, setTasks] = useState<TaskRow[]>([])
  const [fileName, setFileName] = useState(INITIAL_FILE_NAME)
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState('')
  const [pivotRowGrouping, setPivotRowGrouping] = useState<PivotRowGrouping>('project')
  const [chartGroupBy, setChartGroupBy] = useState<ChartGroupBy>('project')
  const [includeWeekends, setIncludeWeekends] = useState(false)
  const [manualOverrides, setManualOverrides] = useState<Record<string, number>>({})
  const [isPivotCollapsed, setIsPivotCollapsed] = useState(false)
  const [selectedProjects, setSelectedProjects] = useState<Set<string>>(new Set())
  const [resourceWeeklyCapacities, setResourceWeeklyCapacities] = useState<Record<string, number>>({})
  const [enabledResources, setEnabledResources] = useState<Record<string, boolean>>({})
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

  const allProjects = useMemo(() => uniqueSorted(tasks.map((task) => task.project)), [tasks])
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
    if (allProjects.length === 0) {
      setSelectedProjects(new Set())
      setProjectsInitialized(false)
      return
    }

    if (!projectsInitialized) {
      setSelectedProjects(new Set(allProjects))
      setProjectsInitialized(true)
      return
    }

    setSelectedProjects((current) => new Set([...current].filter((project) => allProjects.includes(project))))
  }, [allProjects, projectsInitialized])

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
    () => buildBaseLeafCells(tasks, filters, includeWeekends, enabledResourceSet),
    [tasks, filters, includeWeekends, enabledResourceSet],
  )
  const selectedProjectsForCalc = useMemo(() => selectedProjects, [selectedProjects])

  const { baseByKey, finalByKey } = useMemo(
    () => buildLeafValueMap(baseLayer.leafCells, manualOverrides, selectedProjectsForCalc),
    [baseLayer.leafCells, manualOverrides, selectedProjectsForCalc],
  )

  const selectedWeeklyCapacity = useMemo(
    () =>
      enabledResourceList.reduce((sum, resource) => {
        const weekly = resourceWeeklyCapacities[resource] ?? 0
        return sum + weekly
      }, 0),
    [enabledResourceList, resourceWeeklyCapacities],
  )

  const selectedMonthlyCapacity = useMemo(
    () =>
      enabledResourceList.reduce((sum, resource) => {
        const weekly = resourceWeeklyCapacities[resource] ?? 0
        return sum + monthlyFromWeekly(weekly)
      }, 0),
    [enabledResourceList, resourceWeeklyCapacities],
  )

  const weeklyBuckets = useMemo(
    () => buildWeeklyBucketsFromLeaf(finalByKey, baseLayer.weekKeys, selectedWeeklyCapacity, chartGroupBy),
    [finalByKey, baseLayer.weekKeys, selectedWeeklyCapacity, chartGroupBy],
  )

  const monthlyBuckets = useMemo(
    () => buildMonthlyBuckets(weeklyBuckets, selectedMonthlyCapacity),
    [weeklyBuckets, selectedMonthlyCapacity],
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

  function exportTableCsv(): void {
    const csv = weeklyBucketsToCsv(weeklyBuckets)
    const dateStamp = format(new Date(), 'yyyy-MM-dd')
    downloadCsv(`weekly-forecast-${dateStamp}.csv`, csv)
  }

  function resetFilters(): void {
    setFilters((current) => ({ ...current, dateFrom: '', dateTo: '', year: current.year, resources: [] }))
    setSelectedProjects(new Set(allProjects))
    setEnabledResources(() => {
      const next: Record<string, boolean> = {}
      resources.forEach((resource) => {
        next[resource] = true
      })
      return next
    })
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
        <h1>Production Capacity Planning Dashboard</h1>
        <p className="subtitle">
          Import forecast data, edit weekly hours in the planning pivot, and watch chart/table/capacity metrics update
          instantly.
        </p>

        <div className="controls-grid">
          <label>
            Upload or Replace Workbook (.xlsx)
            <input type="file" accept=".xlsx" onChange={handleUpload} />
          </label>

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

          <MultiSelectProjects
            options={allProjects}
            selectedValues={allProjects.filter((project) => selectedProjects.has(project))}
            onChange={(nextSelected) => setSelectedProjects(new Set(nextSelected))}
            placeholder="Projects"
            entityPlural="Projects"
            searchPlaceholder="Search projects..."
            noMatchingText="No matching projects"
            ariaLabel="Projects"
          />

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

          <label className="checkbox-label">
            <input
              type="checkbox"
              checked={includeWeekends}
              onChange={(event) => setIncludeWeekends(event.target.checked)}
            />
            Include weekends in distribution
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
          <button type="button" onClick={exportTableCsv} disabled={weeklyBuckets.length === 0}>
            Export Weekly CSV
          </button>
        </div>
      </header>

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
            <span>Selected Monthly Capacity</span>
            <strong>{selectedMonthlyCapacity.toFixed(2)}</strong>
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
        />
      )}

      {isLoading && <div className="panel status">Loading workbook...</div>}
      {!isLoading && error && <div className="panel status error">{error}</div>}
      {!isLoading && !error && weeklyBuckets.length === 0 && (
        <div className="panel status">No weekly forecast buckets match current filter and project toggle settings.</div>
      )}

      {!isLoading && !error && weeklyBuckets.length > 0 && (
        <>
          <ForecastChart weeklyBuckets={weeklyBuckets} categoryKeys={categoryKeys} />
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
          <ForecastTable weeklyBuckets={weeklyBuckets} />
          <MonthlyForecastTable monthlyBuckets={monthlyBuckets} />
        </>
      )}

      {!allResourcesVisible && !isLoading && (
        <div className="panel status">No resources available in the current data scope.</div>
      )}
    </div>
  )
}

export default App
