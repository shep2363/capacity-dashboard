import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { addDays, format, parseISO, startOfWeek } from 'date-fns'
import DepartmentPage, { buildDepartmentRows, type DepartmentFilters, type DepartmentRow } from './components/DepartmentPage'
import { DetailingPage } from './components/DetailingPage'
import { HowToUsePage } from './components/HowToUsePage'
import { PivotPlanningTable } from './components/PivotPlanningTable'
import { ReportWorkspace, type ReportTab } from './components/ReportWorkspace'
import { RevenueWorkspace } from './components/RevenueWorkspace'
import { ResourceCapacityTable } from './components/ResourceCapacityTable'
import { ExecutiveSummary, type ExecutiveData, type KpiSet } from './components/ExecutiveSummary'
import type { AppFilters, ChartGroupBy, PivotRowGrouping, TaskRow } from './types'
import { parseSalesSpreadsheet, parseSpreadsheet } from './utils/excel'
import { exportReportWorkbook, type SummaryMetric } from './utils/reportExport'
import { exportReportWorkbookWithChartApi } from './utils/reportExportApi'
import { downloadActiveWorkbook, fetchActiveWorkbookStatus, uploadActiveWorkbook } from './utils/activeWorkbookApi'
import {
  PlanningStateApiError,
  fetchPlanningState,
  savePlanningState,
  type PlanningStatePayload,
  type WeekCapacitySchedule,
} from './utils/planningStateApi'
import {
  RevenueRatesApiError,
  fetchRevenueRates,
  saveRevenueRates,
  type RevenueRateMap,
  type RevenueRatesPayload,
} from './utils/revenueRatesApi'
import { fetchSmartsheetProgress, SmartsheetProgressApiError } from './utils/smartsheetProgressApi'
import { buildRevenueMetrics, buildRevenueRateRows, normalizeRateMap } from './utils/revenue'
import { buildDepartmentProgressMatcher, formatSmartsheetSyncLabel, type SmartsheetProgressEntry } from './utils/smartsheetProgress'
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

const INITIAL_FILE_NAME = 'Hours_04-22-26.xlsx'
const INITIAL_SALES_FILE_NAME = '2026 Sales Production Report.xlsx'
const APP_ADMIN_PASSWORD = '2431'
const APP_USER_PASSWORD = '1357'
const APP_FORECAST_PASSWORD = '9876'
const APP_UNLOCK_SESSION_KEY = 'capacity_dashboard_unlocked'
const APP_ROLE_SESSION_KEY = 'capacity_dashboard_role'
const DEFAULT_MAX_RATE_PER_HOUR = 1_000_000
const DEFAULT_RESOURCE_WEEKLY: Record<string, number> = {
  Fabrication: 1520,
  Assembly: 0,
  Processing: 280,
  Paint: 80,
  Shipping: 160,
  Detailer: 0,
}

const PROJECT_COLOR_PALETTE = [
  '#ef4444',  // Red
  '#3b82f6',  // Blue
  '#22c55e',  // Green
  '#f97316',  // Orange
  '#8b5cf6',  // Violet
  '#facc15',  // Yellow
  '#06b6d4',  // Cyan
  '#ec4899',  // Pink
  '#84cc16',  // Lime
  '#a855f7',  // Purple
  '#14b8a6',  // Teal
  '#f59e0b',  // Amber
  '#d946ef',  // Fuchsia
  '#6366f1',  // Indigo
  '#10b981',  // Emerald
  '#0ea5e9',  // Sky
]

type PageKey =
  | 'howToUse'
  | 'executive'
  | 'planning'
  | 'report'
  | 'processing'
  | 'fabrication'
  | 'assembly'
  | 'paint'
  | 'shipping'
  | 'detailing'
  | 'revenue'
type AccessRole = 'admin' | 'user' | 'forecast'
const PAGE_TAB_ORDER: PageKey[] = [
  'howToUse',
  'executive',
  'report',
  'processing',
  'fabrication',
  'assembly',
  'paint',
  'shipping',
  'detailing',
  'planning',
  'revenue',
]
const DEPARTMENT_RESOURCES: Array<PageKey> = ['processing', 'fabrication', 'assembly', 'paint', 'shipping', 'detailing']
const DEPT_RESOURCE_LABEL: Record<PageKey, string> = {
  howToUse: 'How to Use App',
  executive: 'Executive Summary',
  planning: 'Planning',
  report: 'Report Workspace',
  processing: 'Processing',
  fabrication: 'Fabrication',
  assembly: 'Assembly',
  paint: 'Paint',
  shipping: 'Shipping',
  detailing: 'Detailing',
  revenue: 'Revenue',
}
const DEFAULT_DEPT_FILTER: DepartmentFilters = { projects: [], sequences: [], weeks: [], statuses: [] }
type PlanningSaveStatus = 'idle' | 'saving' | 'saved' | 'error'
type RevenueSaveStatus = 'idle' | 'saving' | 'saved' | 'error'
type SmartsheetSyncStatus = 'idle' | 'loading' | 'loaded' | 'error'
type RevenueRateField = 'revenuePerHour' | 'grossProfitPerHour'

function toNumericOverrides(overrides: Record<string, number>): Record<string, number> {
  const normalized: Record<string, number> = {}
  Object.entries(overrides).forEach(([key, value]) => {
    const numeric = Number(value)
    if (Number.isFinite(numeric) && numeric >= 0) {
      normalized[key] = numeric
    }
  })
  return normalized
}

function toNumericWeekCapacitySchedule(schedule: WeekCapacitySchedule | undefined): WeekCapacitySchedule {
  const normalized: WeekCapacitySchedule = {}
  if (!schedule || typeof schedule !== 'object') {
    return normalized
  }

  Object.entries(schedule).forEach(([weekIso, value]) => {
    if (!weekIso) {
      return
    }
    const numeric = Number(value)
    if (Number.isFinite(numeric) && numeric >= 0) {
      normalized[weekIso] = numeric
    }
  })

  return normalized
}

function formatPlanningSaveLabel(status: PlanningSaveStatus, updatedAt: string | null, errorMessage: string): string {
  if (status === 'saving') {
    return 'Saving planning edits...'
  }
  if (status === 'saved') {
    if (updatedAt) {
      return `Saved ${new Date(updatedAt).toLocaleString()}`
    }
    return 'Saved'
  }
  if (status === 'error') {
    return errorMessage || 'Save failed'
  }
  return updatedAt ? `Loaded ${new Date(updatedAt).toLocaleString()}` : 'Not yet saved'
}

function formatRevenueSaveLabel(status: RevenueSaveStatus, updatedAt: string | null, errorMessage: string): string {
  if (status === 'saving') {
    return 'Saving rates...'
  }
  if (status === 'saved') {
    if (updatedAt) {
      return `Saved ${new Date(updatedAt).toLocaleString()}`
    }
    return 'Saved'
  }
  if (status === 'error') {
    return errorMessage || 'Save failed'
  }
  return updatedAt ? `Loaded ${new Date(updatedAt).toLocaleString()}` : 'Not yet saved'
}

// Meeus/Jones/Butcher computus for Gregorian calendar to find Easter Sunday.
const computeEasterSunday = (year: number): Date => {
  const a = year % 19
  const b = Math.floor(year / 100)
  const c = year % 100
  const d = Math.floor(b / 4)
  const e = b % 4
  const f = Math.floor((b + 8) / 25)
  const g = Math.floor((b - f + 1) / 3)
  const h = (19 * a + b - d - g + 15) % 30
  const i = Math.floor(c / 4)
  const k = c % 4
  const l = (32 + 2 * e + 2 * i - h - k) % 7
  const m = Math.floor((a + 11 * h + 22 * l) / 451)
  const month = Math.floor((h + l - 7 * m + 114) / 31) // 3 = March, 4 = April (1-based)
  const day = ((h + l - 7 * m + 114) % 31) + 1
  return new Date(year, month - 1, day)
}

function uniqueSorted(values: string[]): string[] {
  return [...new Set(values)].sort((a, b) => a.localeCompare(b))
}

function normalizeSalesProbability(value: number): number {
  const rounded = Math.round(value * 1000) / 1000
  return Object.is(rounded, -0) ? 0 : rounded
}

function createDefaultFilters(): AppFilters {
  const now = new Date()
  const currentWeekMonday = startOfWeek(now, { weekStartsOn: 1 })
  const currentYear = now.getFullYear()
  return {
    dateFrom: format(currentWeekMonday, 'yyyy-MM-dd'),
    dateTo: `${currentYear}-12-31`,
    year: String(currentYear),
    resources: [],
  }
}

function applyHolidayAdjustedWeeklyCapacity(
  baselineCapacity: number,
  weekendCapacity: number,
  holidayCount: number,
): number {
  const safeBaseline = Number.isFinite(baselineCapacity) && baselineCapacity > 0 ? baselineCapacity : 0
  const safeWeekend = Number.isFinite(weekendCapacity) && weekendCapacity > 0 ? weekendCapacity : 0
  const safeHolidayCount = Number.isFinite(holidayCount) && holidayCount > 0 ? holidayCount : 0
  const holidayReduction = safeHolidayCount > 0 ? Math.min(safeBaseline, (safeBaseline / 5) * safeHolidayCount) : 0
  return safeBaseline + safeWeekend - holidayReduction
}

function getInitialActivePage(): PageKey {
  if (typeof window === 'undefined') {
    return 'executive'
  }
  const role = window.sessionStorage.getItem(APP_ROLE_SESSION_KEY)
  if (role === 'user') return 'howToUse'
  if (role === 'forecast') return 'report'
  return 'executive'
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
  const [salesFileName, setSalesFileName] = useState(INITIAL_SALES_FILE_NAME)
  const [isLoading, setIsLoading] = useState(false)
  const [error, setError] = useState('')
  const [pivotRowGrouping, setPivotRowGrouping] = useState<PivotRowGrouping>('project')
  const [chartGroupBy, setChartGroupBy] = useState<ChartGroupBy>('project')
  const [selectedWeekendDates, setSelectedWeekendDates] = useState<Set<string>>(new Set())
  const [manualOverrides, setManualOverrides] = useState<Record<string, number>>({})
  const [salesManualOverrides, setSalesManualOverrides] = useState<Record<string, number>>({})
  const [mainPlanningSaveStatus, setMainPlanningSaveStatus] = useState<PlanningSaveStatus>('idle')
  const [salesPlanningSaveStatus, setSalesPlanningSaveStatus] = useState<PlanningSaveStatus>('idle')
  const [mainPlanningSaveError, setMainPlanningSaveError] = useState('')
  const [salesPlanningSaveError, setSalesPlanningSaveError] = useState('')
  const [mainPlanningUpdatedAt, setMainPlanningUpdatedAt] = useState<string | null>(null)
  const [salesPlanningUpdatedAt, setSalesPlanningUpdatedAt] = useState<string | null>(null)
  const [mainPlanningSyncReady, setMainPlanningSyncReady] = useState(false)
  const [salesPlanningSyncReady, setSalesPlanningSyncReady] = useState(false)
  const [mainRevenueRates, setMainRevenueRates] = useState<RevenueRateMap>({})
  const [salesRevenueRates, setSalesRevenueRates] = useState<RevenueRateMap>({})
  const [mainRevenueSaveStatus, setMainRevenueSaveStatus] = useState<RevenueSaveStatus>('idle')
  const [salesRevenueSaveStatus, setSalesRevenueSaveStatus] = useState<RevenueSaveStatus>('idle')
  const [mainRevenueSaveError, setMainRevenueSaveError] = useState('')
  const [salesRevenueSaveError, setSalesRevenueSaveError] = useState('')
  const [mainRevenueUpdatedAt, setMainRevenueUpdatedAt] = useState<string | null>(null)
  const [salesRevenueUpdatedAt, setSalesRevenueUpdatedAt] = useState<string | null>(null)
  const [mainRevenueSyncReady, setMainRevenueSyncReady] = useState(false)
  const [salesRevenueSyncReady, setSalesRevenueSyncReady] = useState(false)
  const [isPivotCollapsed, setIsPivotCollapsed] = useState(true)
  const [isSalesPivotCollapsed, setIsSalesPivotCollapsed] = useState(true)
  const [selectedProjects, setSelectedProjects] = useState<Set<string>>(new Set())
  const [salesSelectedProjects, setSalesSelectedProjects] = useState<Set<string>>(new Set())
  const [selectedSalesProbabilities, setSelectedSalesProbabilities] = useState<Set<number>>(new Set())
  const [resourceWeeklyCapacities, setResourceWeeklyCapacities] = useState<Record<string, number>>({})
  const [weekCapacitySchedule, setWeekCapacitySchedule] = useState<WeekCapacitySchedule>({})
  const [enabledResources, setEnabledResources] = useState<Record<string, boolean>>({})
  const [salesEnabledResources, setSalesEnabledResources] = useState<Record<string, boolean>>({})
  const [weekendExtraByResource, setWeekendExtraByResource] = useState<Record<string, number>>({})
  const [holidayDates, setHolidayDates] = useState<Array<{ iso: string; name: string }>>([])
  const [projectsInitialized, setProjectsInitialized] = useState(false)
  const [salesProjectsInitialized, setSalesProjectsInitialized] = useState(false)
  const [salesProbabilitiesInitialized, setSalesProbabilitiesInitialized] = useState(false)
  const [pivotWeekWindowSize, setPivotWeekWindowSize] = useState(12)
  const [pivotWeekStartIndex, setPivotWeekStartIndex] = useState(0)
  const [collapseResetToken, setCollapseResetToken] = useState(0)
  const [salesCollapseResetToken, setSalesCollapseResetToken] = useState(0)
  const [salesPivotWeekWindowSize, setSalesPivotWeekWindowSize] = useState(12)
  const [salesPivotWeekStartIndex, setSalesPivotWeekStartIndex] = useState(0)
  const [isHeaderCollapsed, setIsHeaderCollapsed] = useState(true)
  const [activePage, setActivePage] = useState<PageKey>(getInitialActivePage)
  const [accessRole, setAccessRole] = useState<AccessRole | null>(() => {
    if (typeof window === 'undefined') {
      return null
    }
    const persistedRole = window.sessionStorage.getItem(APP_ROLE_SESSION_KEY)
    if (persistedRole === 'admin' || persistedRole === 'user' || persistedRole === 'forecast') {
      return persistedRole
    }
    if (window.sessionStorage.getItem(APP_UNLOCK_SESSION_KEY) === 'true') {
      return 'admin'
    }
    return null
  })
  const [isUnlocked, setIsUnlocked] = useState<boolean>(() => {
    if (typeof window === 'undefined') {
      return false
    }
    const hasUnlockFlag = window.sessionStorage.getItem(APP_UNLOCK_SESSION_KEY) === 'true'
    const persistedRole = window.sessionStorage.getItem(APP_ROLE_SESSION_KEY)
    return hasUnlockFlag || persistedRole === 'admin' || persistedRole === 'user' || persistedRole === 'forecast'
  })
  const [passwordInput, setPasswordInput] = useState('')
  const [passwordError, setPasswordError] = useState('')
  const [hoveredProject, setHoveredProject] = useState<string | null>(null)
  const [deptFilters, setDeptFilters] = useState<Record<string, DepartmentFilters>>({})
  const [smartsheetProgressEntries, setSmartsheetProgressEntries] = useState<SmartsheetProgressEntry[]>([])
  const [smartsheetSyncStatus, setSmartsheetSyncStatus] = useState<SmartsheetSyncStatus>('idle')
  const [smartsheetSyncError, setSmartsheetSyncError] = useState('')
  const [smartsheetUpdatedAt, setSmartsheetUpdatedAt] = useState<string | null>(null)
  const mainPlanningVersionRef = useRef(0)
  const salesPlanningVersionRef = useRef(0)
  const mainPlanningSkipFirstPersistRef = useRef(true)
  const salesPlanningSkipFirstPersistRef = useRef(true)
  const mainPlanningRequestSeqRef = useRef(0)
  const salesPlanningRequestSeqRef = useRef(0)
  const mainRevenueVersionRef = useRef(0)
  const salesRevenueVersionRef = useRef(0)
  const mainRevenueSkipFirstPersistRef = useRef(true)
  const salesRevenueSkipFirstPersistRef = useRef(true)
  const mainRevenueRequestSeqRef = useRef(0)
  const salesRevenueRequestSeqRef = useRef(0)
  const pageTabs: Array<{ key: PageKey; label: string }> = PAGE_TAB_ORDER
    .filter((key) => accessRole === 'forecast' ? key === 'report' : true)
    .map((key) => ({
      key,
      label: DEPT_RESOURCE_LABEL[key],
    }))

  const [filters, setFilters] = useState<AppFilters>(createDefaultFilters())
  const maxRatePerHour = useMemo(() => {
    const raw = Number((import.meta.env as Record<string, string | undefined>).VITE_MAX_RATE_PER_HOUR ?? DEFAULT_MAX_RATE_PER_HOUR)
    if (!Number.isFinite(raw) || raw <= 0) {
      return DEFAULT_MAX_RATE_PER_HOUR
    }
    return raw
  }, [])

  async function refreshSmartsheetProgress(): Promise<void> {
    setSmartsheetSyncStatus('loading')
    setSmartsheetSyncError('')
    try {
      const payload = await fetchSmartsheetProgress()
      setSmartsheetProgressEntries(payload.rows)
      setSmartsheetUpdatedAt(payload.updatedAt ?? null)
      setSmartsheetSyncStatus('loaded')
    } catch (error) {
      const message =
        error instanceof SmartsheetProgressApiError || error instanceof Error
          ? error.message
          : 'Failed loading Smartsheet progress.'
      setSmartsheetSyncStatus('error')
      setSmartsheetSyncError(message)
    }
  }

  useEffect(() => {
    void refreshSmartsheetProgress()
  }, [])

  useEffect(() => {
    let cancelled = false

    async function loadBundledDefaultWorkbook(): Promise<void> {
      const response = await fetch(`/${INITIAL_FILE_NAME}`)
      if (!response.ok) {
        throw new Error(`Unable to fetch ${INITIAL_FILE_NAME}`)
      }

      const workbookData = await response.arrayBuffer()
      const parsedTasks = parseSpreadsheet(workbookData)
      if (cancelled) {
        return
      }

      setTasks(parsedTasks)
      setFileName(INITIAL_FILE_NAME)
      setManualOverrides({})
      setEnabledResources({})
      setResourceWeeklyCapacities({})
      setWeekCapacitySchedule({})
      setFilters(createDefaultFilters())
      setSelectedWeekendDates(new Set())
      setWeekendExtraByResource({})
      setProjectsInitialized(false)
    }

    function resetSalesPlanningState(file: string): void {
      setSalesFileName(file || INITIAL_SALES_FILE_NAME)
      setSalesManualOverrides({})
      setSalesEnabledResources({})
      setSalesSelectedProjects(new Set())
      setSalesProjectsInitialized(false)
    }

    async function loadInitialWorkbook(): Promise<void> {
      setMainPlanningSyncReady(false)
      setSalesPlanningSyncReady(false)
      setMainRevenueSyncReady(false)
      setSalesRevenueSyncReady(false)
      setMainPlanningSaveError('')
      setSalesPlanningSaveError('')
      setMainPlanningSaveStatus('idle')
      setSalesPlanningSaveStatus('idle')
      setMainRevenueSaveError('')
      setSalesRevenueSaveError('')
      setMainRevenueSaveStatus('idle')
      setSalesRevenueSaveStatus('idle')
      mainPlanningSkipFirstPersistRef.current = true
      salesPlanningSkipFirstPersistRef.current = true
      mainRevenueSkipFirstPersistRef.current = true
      salesRevenueSkipFirstPersistRef.current = true
      setIsLoading(true)
      setError('')

      try {
        let mainLoadedFromApi = false
        try {
          const activeStatus = await fetchActiveWorkbookStatus('main')
          if (activeStatus.hasWorkbook) {
            const workbookData = await downloadActiveWorkbook('main')
            const parsedTasks = parseSpreadsheet(workbookData)
            if (!cancelled) {
              setTasks(parsedTasks)
              setFileName(activeStatus.fileName || INITIAL_FILE_NAME)
              setManualOverrides({})
              setEnabledResources({})
              setResourceWeeklyCapacities({})
              setWeekCapacitySchedule({})
              setFilters(createDefaultFilters())
              setSelectedWeekendDates(new Set())
              setWeekendExtraByResource({})
              setProjectsInitialized(false)
            }
            mainLoadedFromApi = true
          }
        } catch (sharedLoadError) {
          console.warn('[capacity-dashboard] active workbook API load failed', sharedLoadError)
        }

        if (!mainLoadedFromApi) {
          await loadBundledDefaultWorkbook()
        }

        try {
          const mainPlanningState = await fetchPlanningState('main')
          if (!cancelled) {
            const normalized = toNumericOverrides(mainPlanningState.overrides ?? {})
            const normalizedWeekCapacitySchedule = toNumericWeekCapacitySchedule(mainPlanningState.weekCapacitySchedule)
            setManualOverrides(normalized)
            setWeekCapacitySchedule(normalizedWeekCapacitySchedule)
            mainPlanningVersionRef.current = mainPlanningState.version ?? 0
            setMainPlanningUpdatedAt(mainPlanningState.updatedAt ?? null)
            setMainPlanningSaveError('')
            setMainPlanningSaveStatus('idle')
          }
        } catch (planningLoadError) {
          console.warn('[capacity-dashboard] shared main planning state load failed', planningLoadError)
          if (!cancelled) {
            setManualOverrides({})
            setWeekCapacitySchedule({})
            mainPlanningVersionRef.current = 0
            setMainPlanningUpdatedAt(null)
            setMainPlanningSaveStatus('error')
            setMainPlanningSaveError('Failed to load shared planning edits.')
          }
        } finally {
          if (!cancelled) {
            setMainPlanningSyncReady(true)
          }
        }

        try {
          const salesStatus = await fetchActiveWorkbookStatus('sales')
          if (salesStatus.hasWorkbook) {
            const salesWorkbookData = await downloadActiveWorkbook('sales')
            const parsedSalesTasks = parseSalesSpreadsheet(salesWorkbookData)
            if (!cancelled) {
              setSalesTasks(parsedSalesTasks)
              resetSalesPlanningState(salesStatus.fileName || INITIAL_SALES_FILE_NAME)
            }
          } else if (!cancelled) {
            setSalesTasks([])
            resetSalesPlanningState(INITIAL_SALES_FILE_NAME)
          }
        } catch (salesLoadError) {
          console.warn('[capacity-dashboard] sales workbook API load failed', salesLoadError)
          if (!cancelled) {
            setSalesTasks([])
            resetSalesPlanningState(INITIAL_SALES_FILE_NAME)
          }
        }

        try {
          const salesPlanningState = await fetchPlanningState('sales')
          if (!cancelled) {
            const normalized = toNumericOverrides(salesPlanningState.overrides ?? {})
            setSalesManualOverrides(normalized)
            salesPlanningVersionRef.current = salesPlanningState.version ?? 0
            setSalesPlanningUpdatedAt(salesPlanningState.updatedAt ?? null)
            setSalesPlanningSaveError('')
            setSalesPlanningSaveStatus('idle')
          }
        } catch (salesPlanningLoadError) {
          console.warn('[capacity-dashboard] shared sales planning state load failed', salesPlanningLoadError)
          if (!cancelled) {
            setSalesManualOverrides({})
            salesPlanningVersionRef.current = 0
            setSalesPlanningUpdatedAt(null)
            setSalesPlanningSaveStatus('error')
            setSalesPlanningSaveError('Failed to load shared sales planning edits.')
          }
        } finally {
          if (!cancelled) {
            setSalesPlanningSyncReady(true)
          }
        }

        try {
          const mainRevenueState = await fetchRevenueRates('main')
          if (!cancelled) {
            const normalized = normalizeRateMap(mainRevenueState.rates ?? {})
            setMainRevenueRates(normalized)
            mainRevenueVersionRef.current = mainRevenueState.version ?? 0
            setMainRevenueUpdatedAt(mainRevenueState.updatedAt ?? null)
            setMainRevenueSaveError('')
            setMainRevenueSaveStatus('idle')
          }
        } catch (mainRevenueLoadError) {
          console.warn('[capacity-dashboard] shared main revenue rates load failed', mainRevenueLoadError)
          if (!cancelled) {
            setMainRevenueRates({})
            mainRevenueVersionRef.current = 0
            setMainRevenueUpdatedAt(null)
            setMainRevenueSaveStatus('error')
            setMainRevenueSaveError('Failed to load shared shop revenue rates.')
          }
        } finally {
          if (!cancelled) {
            setMainRevenueSyncReady(true)
          }
        }

        try {
          const salesRevenueState = await fetchRevenueRates('sales')
          if (!cancelled) {
            const normalized = normalizeRateMap(salesRevenueState.rates ?? {})
            setSalesRevenueRates(normalized)
            salesRevenueVersionRef.current = salesRevenueState.version ?? 0
            setSalesRevenueUpdatedAt(salesRevenueState.updatedAt ?? null)
            setSalesRevenueSaveError('')
            setSalesRevenueSaveStatus('idle')
          }
        } catch (salesRevenueLoadError) {
          console.warn('[capacity-dashboard] shared sales revenue rates load failed', salesRevenueLoadError)
          if (!cancelled) {
            setSalesRevenueRates({})
            salesRevenueVersionRef.current = 0
            setSalesRevenueUpdatedAt(null)
            setSalesRevenueSaveStatus('error')
            setSalesRevenueSaveError('Failed to load shared sales revenue rates.')
          }
        } finally {
          if (!cancelled) {
            setSalesRevenueSyncReady(true)
          }
        }
      } catch {
        if (!cancelled) {
          setError('Could not load the active workbook. Upload a .xlsx file to continue.')
          setMainPlanningSyncReady(true)
          setSalesPlanningSyncReady(true)
          setMainRevenueSyncReady(true)
          setSalesRevenueSyncReady(true)
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false)
        }
      }
    }

    void loadInitialWorkbook()

    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    if (!mainPlanningSyncReady) {
      return
    }
    if (mainPlanningSkipFirstPersistRef.current) {
      mainPlanningSkipFirstPersistRef.current = false
      return
    }

    const normalizedOverrides = toNumericOverrides(manualOverrides)
    const normalizedWeekCapacitySchedule = toNumericWeekCapacitySchedule(weekCapacitySchedule)
    const requestSeq = mainPlanningRequestSeqRef.current + 1
    mainPlanningRequestSeqRef.current = requestSeq
    setMainPlanningSaveStatus('saving')
    setMainPlanningSaveError('')

    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          let saved: PlanningStatePayload
          try {
            saved = await savePlanningState('main', normalizedOverrides, {
              baseVersion: mainPlanningVersionRef.current,
              source: 'planning-ui',
              weekCapacitySchedule: normalizedWeekCapacitySchedule,
            })
          } catch (error) {
            if (error instanceof PlanningStateApiError && error.status === 409) {
              const latest = await fetchPlanningState('main')
              mainPlanningVersionRef.current = latest.version
              saved = await savePlanningState('main', normalizedOverrides, {
                baseVersion: mainPlanningVersionRef.current,
                source: 'planning-ui-retry',
                weekCapacitySchedule: normalizedWeekCapacitySchedule,
              })
            } else {
              throw error
            }
          }
          if (requestSeq !== mainPlanningRequestSeqRef.current) {
            return
          }
          mainPlanningVersionRef.current = saved.version
          setMainPlanningUpdatedAt(saved.updatedAt)
          setMainPlanningSaveStatus('saved')
          setMainPlanningSaveError('')
        } catch (error) {
          if (requestSeq !== mainPlanningRequestSeqRef.current) {
            return
          }
          const message = error instanceof Error ? error.message : 'Failed to save shared planning edits.'
          setMainPlanningSaveStatus('error')
          setMainPlanningSaveError(message)
        }
      })()
    }, 350)

    return () => window.clearTimeout(timer)
  }, [mainPlanningSyncReady, manualOverrides, weekCapacitySchedule])

  useEffect(() => {
    if (!salesPlanningSyncReady) {
      return
    }
    if (salesPlanningSkipFirstPersistRef.current) {
      salesPlanningSkipFirstPersistRef.current = false
      return
    }

    const normalizedOverrides = toNumericOverrides(salesManualOverrides)
    const requestSeq = salesPlanningRequestSeqRef.current + 1
    salesPlanningRequestSeqRef.current = requestSeq
    setSalesPlanningSaveStatus('saving')
    setSalesPlanningSaveError('')

    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          let saved: PlanningStatePayload
          try {
            saved = await savePlanningState('sales', normalizedOverrides, {
              baseVersion: salesPlanningVersionRef.current,
              source: 'sales-planning-ui',
            })
          } catch (error) {
            if (error instanceof PlanningStateApiError && error.status === 409) {
              const latest = await fetchPlanningState('sales')
              salesPlanningVersionRef.current = latest.version
              saved = await savePlanningState('sales', normalizedOverrides, {
                baseVersion: salesPlanningVersionRef.current,
                source: 'sales-planning-ui-retry',
              })
            } else {
              throw error
            }
          }
          if (requestSeq !== salesPlanningRequestSeqRef.current) {
            return
          }
          salesPlanningVersionRef.current = saved.version
          setSalesPlanningUpdatedAt(saved.updatedAt)
          setSalesPlanningSaveStatus('saved')
          setSalesPlanningSaveError('')
        } catch (error) {
          if (requestSeq !== salesPlanningRequestSeqRef.current) {
            return
          }
          const message = error instanceof Error ? error.message : 'Failed to save shared sales planning edits.'
          setSalesPlanningSaveStatus('error')
          setSalesPlanningSaveError(message)
        }
      })()
    }, 350)

    return () => window.clearTimeout(timer)
  }, [salesManualOverrides, salesPlanningSyncReady])

  useEffect(() => {
    if (!mainRevenueSyncReady) {
      return
    }
    if (mainRevenueSkipFirstPersistRef.current) {
      mainRevenueSkipFirstPersistRef.current = false
      return
    }

    const normalizedRates = normalizeRateMap(mainRevenueRates)
    const requestSeq = mainRevenueRequestSeqRef.current + 1
    mainRevenueRequestSeqRef.current = requestSeq
    setMainRevenueSaveStatus('saving')
    setMainRevenueSaveError('')

    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          let saved: RevenueRatesPayload
          try {
            saved = await saveRevenueRates('main', normalizedRates, {
              baseVersion: mainRevenueVersionRef.current,
              source: 'revenue-ui',
            })
          } catch (error) {
            if (error instanceof RevenueRatesApiError && error.status === 409) {
              const latest = await fetchRevenueRates('main')
              mainRevenueVersionRef.current = latest.version
              saved = await saveRevenueRates('main', normalizedRates, {
                baseVersion: mainRevenueVersionRef.current,
                source: 'revenue-ui-retry',
              })
            } else {
              throw error
            }
          }
          if (requestSeq !== mainRevenueRequestSeqRef.current) {
            return
          }
          mainRevenueVersionRef.current = saved.version
          setMainRevenueUpdatedAt(saved.updatedAt)
          setMainRevenueSaveStatus('saved')
          setMainRevenueSaveError('')
        } catch (error) {
          if (requestSeq !== mainRevenueRequestSeqRef.current) {
            return
          }
          const message = error instanceof Error ? error.message : 'Failed to save shared shop revenue rates.'
          setMainRevenueSaveStatus('error')
          setMainRevenueSaveError(message)
        }
      })()
    }, 350)

    return () => window.clearTimeout(timer)
  }, [mainRevenueRates, mainRevenueSyncReady])

  useEffect(() => {
    if (!salesRevenueSyncReady) {
      return
    }
    if (salesRevenueSkipFirstPersistRef.current) {
      salesRevenueSkipFirstPersistRef.current = false
      return
    }

    const normalizedRates = normalizeRateMap(salesRevenueRates)
    const requestSeq = salesRevenueRequestSeqRef.current + 1
    salesRevenueRequestSeqRef.current = requestSeq
    setSalesRevenueSaveStatus('saving')
    setSalesRevenueSaveError('')

    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          let saved: RevenueRatesPayload
          try {
            saved = await saveRevenueRates('sales', normalizedRates, {
              baseVersion: salesRevenueVersionRef.current,
              source: 'sales-revenue-ui',
            })
          } catch (error) {
            if (error instanceof RevenueRatesApiError && error.status === 409) {
              const latest = await fetchRevenueRates('sales')
              salesRevenueVersionRef.current = latest.version
              saved = await saveRevenueRates('sales', normalizedRates, {
                baseVersion: salesRevenueVersionRef.current,
                source: 'sales-revenue-ui-retry',
              })
            } else {
              throw error
            }
          }
          if (requestSeq !== salesRevenueRequestSeqRef.current) {
            return
          }
          salesRevenueVersionRef.current = saved.version
          setSalesRevenueUpdatedAt(saved.updatedAt)
          setSalesRevenueSaveStatus('saved')
          setSalesRevenueSaveError('')
        } catch (error) {
          if (requestSeq !== salesRevenueRequestSeqRef.current) {
            return
          }
          const message = error instanceof Error ? error.message : 'Failed to save shared sales revenue rates.'
          setSalesRevenueSaveStatus('error')
          setSalesRevenueSaveError(message)
        }
      })()
    }, 350)

    return () => window.clearTimeout(timer)
  }, [salesRevenueRates, salesRevenueSyncReady])

  const resources = useMemo(() => uniqueSorted(tasks.map((task) => task.resourceName)), [tasks])
  const salesResources = useMemo(() => uniqueSorted(salesTasks.map((task) => task.resourceName)), [salesTasks])
  const salesProbabilityOptions = useMemo(
    () =>
      [...new Set(
        salesTasks.flatMap((task) => {
          if (typeof task.salesProbability !== 'number' || !Number.isFinite(task.salesProbability)) {
            return []
          }
          return [normalizeSalesProbability(task.salesProbability)]
        }),
      )].sort((a, b) => b - a),
    [salesTasks],
  )
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
      setWeekCapacitySchedule({})
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
  const filteredSalesTasksForReport = useMemo(
    () =>
      salesTasks.filter((task) => {
        if (typeof task.salesProbability !== 'number' || !Number.isFinite(task.salesProbability)) {
          return false
        }
        return selectedSalesProbabilities.has(normalizeSalesProbability(task.salesProbability))
      }),
    [salesTasks, selectedSalesProbabilities],
  )
  const reportSalesBaseLayer = useMemo(
    () => buildBaseLeafCells(filteredSalesTasksForReport, filters, selectedWeekendDates, salesEnabledResourceSet),
    [filteredSalesTasksForReport, filters, selectedWeekendDates, salesEnabledResourceSet],
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
  const reportSalesAvailableProjects = useMemo(() => {
    const totals = new Map<string, number>()
    reportSalesBaseLayer.leafCells.forEach((cell) => {
      totals.set(cell.project, (totals.get(cell.project) ?? 0) + cell.hours)
    })
    return [...totals.entries()]
      .filter(([, hours]) => hours > 0)
      .map(([project]) => project)
      .sort((a, b) => a.localeCompare(b))
  }, [reportSalesBaseLayer.leafCells])
  const reportCombinedProjects = useMemo(
    () => [...availableProjects, ...reportSalesAvailableProjects.map((p) => `Sales - ${p}`)],
    [availableProjects, reportSalesAvailableProjects],
  )
  const projectColors = useMemo(() => {
    const map: Record<string, string> = {}
    availableProjects.forEach((project, index) => {
      map[project] = PROJECT_COLOR_PALETTE[index % PROJECT_COLOR_PALETTE.length]
    })
    return map
  }, [availableProjects])
  const smartsheetProgressMatcher = useMemo(
    () => buildDepartmentProgressMatcher(smartsheetProgressEntries),
    [smartsheetProgressEntries],
  )
  const getDeptFilter = (resource: string): DepartmentFilters => deptFilters[resource] ?? DEFAULT_DEPT_FILTER
  const setDeptFilter = (resource: string, next: DepartmentFilters) =>
    setDeptFilters((current) => ({ ...current, [resource]: next }))

  const applyDeptFilters = (rows: DepartmentRow[], filter: DepartmentFilters): DepartmentRow[] =>
    rows.filter((row) => {
      if (filter.projects.length > 0 && !filter.projects.includes(row.project)) return false
      if (filter.sequences.length > 0 && !filter.sequences.includes(row.sequence)) return false
      if (filter.weeks.length > 0 && !filter.weeks.includes(row.weekStartIso)) return false
      if (filter.statuses.length > 0 && !filter.statuses.includes(row.status)) return false
      return true
    })

  const departmentRowsByResource = useMemo(() => {
    const map: Record<string, DepartmentRow[]> = {}
    const resourceEnabledMap = enabledResources
    DEPARTMENT_RESOURCES.forEach((pageKey) => {
      const resource = DEPT_RESOURCE_LABEL[pageKey] ?? ''
      const baseRows = buildDepartmentRows({
        resource,
        tasks,
        filters,
        selectedProjects,
        selectedWeekendDates,
        resourceEnabled: resourceEnabledMap[resource] !== false,
        progressMatcher: smartsheetProgressMatcher,
      })
      const filter = getDeptFilter(resource)
      const filtered = applyDeptFilters(baseRows, filter).sort(
        (a, b) =>
          a.weekStartIso.localeCompare(b.weekStartIso) ||
          a.project.localeCompare(b.project) ||
          a.sequence.localeCompare(b.sequence),
      )
      map[resource] = filtered
    })
    return map
  }, [tasks, filters, selectedProjects, selectedWeekendDates, enabledResources, deptFilters, smartsheetProgressMatcher])

  const matchedSmartsheetRowCount = useMemo(() => {
    const matchedRowIds = new Set<string>()
    Object.values(departmentRowsByResource).forEach((rows) => {
      rows.forEach((row) => {
        if (row.progressSource === 'smartsheet' && row.smartsheetRowId) {
          matchedRowIds.add(row.smartsheetRowId)
        }
      })
    })
    return matchedRowIds.size
  }, [departmentRowsByResource])

  const smartsheetSyncLabel = useMemo(
    () =>
      formatSmartsheetSyncLabel(
        smartsheetSyncStatus,
        smartsheetUpdatedAt,
        smartsheetSyncError,
        matchedSmartsheetRowCount,
        smartsheetProgressEntries.length,
      ),
    [smartsheetSyncStatus, smartsheetUpdatedAt, smartsheetSyncError, matchedSmartsheetRowCount, smartsheetProgressEntries.length],
  )

  useEffect(() => {
    if (smartsheetProgressEntries.length === 0) {
      return
    }
    const unmatched = smartsheetProgressEntries.length - matchedSmartsheetRowCount
    if (unmatched > 0) {
      console.warn(`[capacity-dashboard] ${unmatched} Smartsheet progress row(s) did not match department rows.`)
    }
  }, [smartsheetProgressEntries.length, matchedSmartsheetRowCount])

  useEffect(() => {
    if (smartsheetSyncStatus !== 'loaded' || smartsheetProgressEntries.length === 0) {
      return
    }

    const unmatchedByDepartment = Object.entries(departmentRowsByResource)
      .map(([resource, rows]) => ({
        resource,
        unmatched: rows.filter((row) => row.progressSource !== 'smartsheet').length,
        total: rows.length,
      }))
      .filter((entry) => entry.total > 0 && entry.unmatched > 0)

    if (unmatchedByDepartment.length > 0) {
      console.info(
        `[capacity-dashboard] Department rows without Smartsheet match: ${unmatchedByDepartment
          .map((entry) => `${entry.resource} ${entry.unmatched}/${entry.total}`)
          .join(', ')}`,
      )
    }
  }, [departmentRowsByResource, smartsheetProgressEntries.length, smartsheetSyncStatus])

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

  useEffect(() => {
    setSalesProbabilitiesInitialized(false)
  }, [salesTasks])

  useEffect(() => {
    if (salesProbabilityOptions.length === 0) {
      setSelectedSalesProbabilities(new Set())
      setSalesProbabilitiesInitialized(false)
      return
    }

    if (!salesProbabilitiesInitialized) {
      setSelectedSalesProbabilities(new Set(salesProbabilityOptions))
      setSalesProbabilitiesInitialized(true)
      return
    }

    setSelectedSalesProbabilities(
      (current) =>
        new Set(
          [...current].filter((probability) =>
            salesProbabilityOptions.some((option) => option === probability),
          ),
        ),
    )
  }, [salesProbabilityOptions, salesProbabilitiesInitialized])
  const selectedProjectsForCalc = useMemo(() => selectedProjects, [selectedProjects])
  const salesSelectedProjectsForCalc = useMemo(() => salesSelectedProjects, [salesSelectedProjects])
  const reportSalesSelectedProjectsForCalc = useMemo(
    () => new Set([...salesSelectedProjects].filter((project) => reportSalesAvailableProjects.includes(project))),
    [salesSelectedProjects, reportSalesAvailableProjects],
  )
  const reportCombinedSelectedProjects = useMemo(() => {
    const ops = [...selectedProjects].filter((p) => availableProjects.includes(p))
    const sales = [...reportSalesSelectedProjectsForCalc].map((p) => `Sales - ${p}`)
    return new Set<string>([...ops, ...sales])
  }, [selectedProjects, availableProjects, reportSalesSelectedProjectsForCalc])

  const { baseByKey, finalByKey } = useMemo(
    () => buildLeafValueMap(baseLayer.leafCells, manualOverrides, selectedProjectsForCalc),
    [baseLayer.leafCells, manualOverrides, selectedProjectsForCalc],
  )
  const { baseByKey: salesBaseByKey, finalByKey: salesFinalByKey } = useMemo(
    () => buildLeafValueMap(salesBaseLayer.leafCells, salesManualOverrides, salesSelectedProjectsForCalc),
    [salesBaseLayer.leafCells, salesManualOverrides, salesSelectedProjectsForCalc],
  )
  const { finalByKey: reportSalesFinalByKey } = useMemo(
    () =>
      buildLeafValueMap(reportSalesBaseLayer.leafCells, salesManualOverrides, reportSalesSelectedProjectsForCalc, {
        limitOverridesToBase: true,
      }),
    [reportSalesBaseLayer.leafCells, salesManualOverrides, reportSalesSelectedProjectsForCalc],
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
    // Good Friday (Friday before Easter Sunday)
    {
      const easter = computeEasterSunday(year)
      const goodFriday = format(addDays(easter, -2), 'yyyy-MM-dd')
      push(goodFriday, 'Good Friday')
    }
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

  useEffect(() => {
    if (resources.length === 0 || allWeekKeys.length === 0) {
      return
    }

    const weekSet = new Set(allWeekKeys)
    setWeekCapacitySchedule((current) => {
      let changed = false
      const next: WeekCapacitySchedule = {}

      Object.entries(current).forEach(([weekIso, value]) => {
        if (!weekSet.has(weekIso)) {
          changed = true
          return
        }
        next[weekIso] = value
      })

      return changed ? next : current
    })
  }, [allWeekKeys, resources.length])

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

  const sortedAllWeekKeys = useMemo(() => [...allWeekKeys].sort((a, b) => a.localeCompare(b)), [allWeekKeys])

  const baseWeeklyResourceCapacityTotal = useMemo(
    () => enabledResourceList.reduce((sum, resource) => sum + (resourceWeeklyCapacities[resource] ?? 0), 0),
    [enabledResourceList, resourceWeeklyCapacities],
  )

  const baseWeekendCapacityByWeek = useMemo(() => {
    const map: Record<string, number> = {}
    for (const weekIso of sortedAllWeekKeys) {
      let weekendTotal = 0
      if (weekendWeeks.has(weekIso)) {
        for (const resource of enabledResourceList) {
          const weekendExtra = weekendExtraByResource[resource] ?? 0
          weekendTotal += weekendExtra
        }
      }
      map[weekIso] = weekendTotal
    }
    return map
  }, [sortedAllWeekKeys, enabledResourceList, weekendExtraByResource, weekendWeeks])

  const baseWeekCapacities = useMemo(() => {
    const map: Record<string, number> = {}
    for (const weekIso of sortedAllWeekKeys) {
      const holidayCount = (holidayWeeks.get(weekIso) ?? []).length
      const weekendTotal = baseWeekendCapacityByWeek[weekIso] ?? 0
      map[weekIso] = applyHolidayAdjustedWeeklyCapacity(baseWeeklyResourceCapacityTotal, weekendTotal, holidayCount)
    }
    return map
  }, [sortedAllWeekKeys, holidayWeeks, baseWeekendCapacityByWeek, baseWeeklyResourceCapacityTotal])

  const weekCapacities = useMemo(() => {
    const map: Record<string, number> = {}
    let scheduledCapacity: number | null = null

    for (const weekIso of sortedAllWeekKeys) {
      const scheduledEntry = weekCapacitySchedule[weekIso]
      if (Number.isFinite(scheduledEntry)) {
        scheduledCapacity = scheduledEntry
      }
      const baselineCapacity = scheduledCapacity ?? baseWeeklyResourceCapacityTotal
      const holidayCount = (holidayWeeks.get(weekIso) ?? []).length
      const weekendTotal = baseWeekendCapacityByWeek[weekIso] ?? 0
      // Scheduled total capacity remains the weekly baseline, then weekend/holiday adjustments are applied.
      map[weekIso] = applyHolidayAdjustedWeeklyCapacity(baselineCapacity, weekendTotal, holidayCount)
    }

    return map
  }, [
    sortedAllWeekKeys,
    weekCapacitySchedule,
    baseWeeklyResourceCapacityTotal,
    holidayWeeks,
    baseWeekendCapacityByWeek,
  ])

  const baseWeekCapacityBeforeAdjustments = useMemo(() => {
    const map: Record<string, number> = {}
    for (const weekIso of sortedAllWeekKeys) {
      map[weekIso] = baseWeeklyResourceCapacityTotal
    }
    return map
  }, [sortedAllWeekKeys, baseWeeklyResourceCapacityTotal])

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
  const reportSalesWeeklyBuckets = useMemo(
    () => buildWeeklyBucketsFromLeaf(reportSalesFinalByKey, allWeekKeys, weekCapacities, chartGroupBy, holidayDetailsByWeek),
    [reportSalesFinalByKey, allWeekKeys, weekCapacities, chartGroupBy, holidayDetailsByWeek],
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
  const reportCombinedWeeklyBuckets = useMemo(() => {
    const weekSet = new Set<string>()
    weeklyBuckets.forEach((bucket) => weekSet.add(bucket.weekStartIso))
    reportSalesWeeklyBuckets.forEach((bucket) => weekSet.add(bucket.weekStartIso))
    const sortedWeeks = [...weekSet].sort((a, b) => parseISO(a).getTime() - parseISO(b).getTime())

    const opsMap = new Map(weeklyBuckets.map((bucket) => [bucket.weekStartIso, bucket]))
    const salesMap = new Map(reportSalesWeeklyBuckets.map((bucket) => [bucket.weekStartIso, bucket]))

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
  }, [weeklyBuckets, reportSalesWeeklyBuckets, weekCapacities])
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
  const reportSalesCategoryKeys = useMemo(() => computeCategoryKeys(reportSalesWeeklyBuckets), [reportSalesWeeklyBuckets])
  const reportCombinedCategoryKeys = useMemo(
    () => computeCategoryKeys(reportCombinedWeeklyBuckets),
    [reportCombinedWeeklyBuckets],
  )

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

  const salesProjectTotals = useMemo(() => {
    const totals = new Map<string, number>()
    const weekSet = new Set(salesBaseLayer.weekKeys)
    Object.entries(salesFinalByKey).forEach(([leafKey, hours]) => {
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
  }, [salesFinalByKey, salesBaseLayer.weekKeys])

  const revenueRateRows = useMemo(
    () =>
      buildRevenueRateRows(
        availableProjects,
        salesAvailableProjects,
        projectTotals,
        salesProjectTotals,
        mainRevenueRates,
        salesRevenueRates,
      ),
    [
      availableProjects,
      salesAvailableProjects,
      projectTotals,
      salesProjectTotals,
      mainRevenueRates,
      salesRevenueRates,
    ],
  )

  const {
    weeklyRevenueRows,
    weeklyProjectKeys,
    weeklyGrossProfitRows,
    weeklyGrossProfitProjectKeys,
    monthlyRevenueRows,
    monthlyGrossProfitRows,
    monthlyProjectKeys,
  } = useMemo(
    () =>
      buildRevenueMetrics({
        mainFinalByKey: finalByKey,
        salesFinalByKey: salesFinalByKey,
        mainWeekKeys: baseLayer.weekKeys,
        salesWeekKeys: salesBaseLayer.weekKeys,
        mainRates: mainRevenueRates,
        salesRates: salesRevenueRates,
      }),
    [finalByKey, salesFinalByKey, baseLayer.weekKeys, salesBaseLayer.weekKeys, mainRevenueRates, salesRevenueRates],
  )

  const topProjects = useMemo(() => {
    const combinedEntries: Array<{ project: string; hours: number }> = []
    projectTotals.forEach((hours, project) => combinedEntries.push({ project, hours }))
    salesProjectTotals.forEach((hours, project) => combinedEntries.push({ project: `Sales - ${project}`, hours }))
    const sorted = combinedEntries.sort((a, b) => b.hours - a.hours)
    const grandTotal = sorted.reduce((sum, item) => sum + item.hours, 0)
    return sorted.slice(0, 5).map((item) => ({
      project: item.project,
      hours: item.hours,
      percent: grandTotal > 0 ? (item.hours / grandTotal) * 100 : 0,
    }))
  }, [projectTotals, salesProjectTotals])

  const executiveData = useMemo<ExecutiveData>(() => {
    function buildKpis(buckets: typeof weeklyBuckets, activeProjectsCount: number, label: string): KpiSet {
      const bookedYtd = buckets.reduce((sum, w) => sum + w.totalHours, 0)
      const capacityYtd = buckets.reduce((sum, w) => sum + w.capacity, 0)
      const utilization = capacityYtd > 0 ? bookedYtd / capacityYtd : 0
      const remaining = capacityYtd - bookedYtd
      const status: KpiSet['status'] = utilization > 1 ? 'red' : utilization >= 0.9 ? 'yellow' : 'green'
      return { bookedYtd, capacityYtd, utilization, remaining, activeProjects: activeProjectsCount, status, label }
    }

    const combinedKpis = buildKpis(combinedWeeklyBuckets, combinedCategoryKeys.length, 'Combined')
    const opsKpis = buildKpis(weeklyBuckets, categoryKeys.length, 'Shop / Ops')
    const salesKpis = buildKpis(salesWeeklyBuckets, salesCategoryKeys.length, 'Sales')

    const monthlyComparison = combinedMonthlyBuckets.map((m, idx) => {
      const ops = monthlyBuckets[idx]
      const sales = salesMonthlyBuckets[idx]
      return {
        month: m.monthLabel,
        opsBooked: ops?.plannedHours ?? 0,
        salesBooked: sales?.plannedHours ?? 0,
        totalBooked: m.plannedHours,
        capacity: m.capacity,
      }
    })

    const quarterlySummary = (() => {
      const map = new Map<string, { booked: number; capacity: number }>()
      combinedMonthlyBuckets.forEach((m) => {
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

    const bookedYtdCombined = combinedWeeklyBuckets.reduce((sum, w) => sum + w.totalHours, 0)
    const capacityYtdCombined = combinedWeeklyBuckets.reduce((sum, w) => sum + w.capacity, 0)
    const utilizationCombined = capacityYtdCombined > 0 ? bookedYtdCombined / capacityYtdCombined : 0

    const annual = {
      booked: bookedYtdCombined,
      capacity: capacityYtdCombined,
      utilization: utilizationCombined,
      status: getCapacityStatus(bookedYtdCombined, capacityYtdCombined),
    }

    const riskMonths = combinedMonthlyBuckets
      .filter((m) => m.plannedHours > m.capacity)
      .map((m) => ({ monthLabel: m.monthLabel, variance: m.variance }))

    const utilizationTrend = combinedMonthlyBuckets.map((m) => ({
      monthLabel: m.monthLabel,
      utilization: m.capacity > 0 ? m.plannedHours / m.capacity : 0,
    }))

    return {
      combinedKpis,
      opsKpis,
      salesKpis,
      monthlyComparison,
      quarterlySummary,
      annual,
      riskMonths,
      topProjects,
      utilizationTrend,
    }
  }, [
    combinedWeeklyBuckets,
    weeklyBuckets,
    salesWeeklyBuckets,
    combinedMonthlyBuckets,
    monthlyBuckets,
    salesMonthlyBuckets,
    combinedCategoryKeys.length,
    categoryKeys.length,
    salesCategoryKeys.length,
    topProjects,
  ])

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

  const showingAllPivotWeeks = pivotWeekWindowSize === -1 || pivotWeekWindowSize >= baseLayer.weekKeys.length
  const effectivePivotWeekWindowSize = showingAllPivotWeeks ? Math.max(baseLayer.weekKeys.length, 1) : pivotWeekWindowSize
  const maxPivotStartIndex = Math.max(0, baseLayer.weekKeys.length - effectivePivotWeekWindowSize)
  const safePivotStartIndex = showingAllPivotWeeks ? 0 : Math.min(pivotWeekStartIndex, maxPivotStartIndex)
  const visiblePivotWeekKeys = showingAllPivotWeeks
    ? baseLayer.weekKeys
    : baseLayer.weekKeys.slice(
        safePivotStartIndex,
        safePivotStartIndex + effectivePivotWeekWindowSize,
      )
  const pivotWeekWindowLabel =
    visiblePivotWeekKeys.length > 0
      ? `${visiblePivotWeekKeys[0]} to ${visiblePivotWeekKeys[visiblePivotWeekKeys.length - 1]}${showingAllPivotWeeks ? ` (All ${visiblePivotWeekKeys.length} weeks)` : ''}`
      : 'No weeks'

  useEffect(() => {
    setSalesPivotWeekStartIndex(0)
  }, [salesBaseLayer.weekKeys, salesPivotWeekWindowSize])

  const showingAllSalesPivotWeeks =
    salesPivotWeekWindowSize === -1 || salesPivotWeekWindowSize >= salesBaseLayer.weekKeys.length
  const effectiveSalesPivotWeekWindowSize = showingAllSalesPivotWeeks
    ? Math.max(salesBaseLayer.weekKeys.length, 1)
    : salesPivotWeekWindowSize
  const salesMaxPivotStartIndex = Math.max(0, salesBaseLayer.weekKeys.length - effectiveSalesPivotWeekWindowSize)
  const salesSafePivotStartIndex = showingAllSalesPivotWeeks
    ? 0
    : Math.min(salesPivotWeekStartIndex, salesMaxPivotStartIndex)
  const salesVisiblePivotWeekKeys = showingAllSalesPivotWeeks
    ? salesBaseLayer.weekKeys
    : salesBaseLayer.weekKeys.slice(
        salesSafePivotStartIndex,
        salesSafePivotStartIndex + effectiveSalesPivotWeekWindowSize,
      )
  const salesPivotWeekWindowLabel =
    salesVisiblePivotWeekKeys.length > 0
      ? `${salesVisiblePivotWeekKeys[0]} to ${salesVisiblePivotWeekKeys[salesVisiblePivotWeekKeys.length - 1]}${showingAllSalesPivotWeeks ? ` (All ${salesVisiblePivotWeekKeys.length} weeks)` : ''}`
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

  const visibleSummaryMetrics = useMemo(
    () => summaryMetrics.filter((item) => item.metric !== 'Selected Weekly Capacity'),
    [summaryMetrics],
  )

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
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setError('Invalid file type. Please upload a .xlsx workbook.')
      event.target.value = ''
      return
    }

    setIsLoading(true)
    setError('')

    try {
      const arrayBuffer = await file.arrayBuffer()
      const parsed = parseSpreadsheet(arrayBuffer)
      const activeWorkbook = await uploadActiveWorkbook('main', file)
      setTasks(parsed)
      setFileName(activeWorkbook.fileName || file.name)
      setManualOverrides({})
      setFilters(createDefaultFilters())
      setSelectedWeekendDates(new Set())
      setProjectsInitialized(false)
      // Preserve resource capacity inputs across workbook replacement while resetting upload-specific view state.
      setIsPivotCollapsed(true)
      setCollapseResetToken((current) => current + 1)
    } catch (uploadError) {
      if (uploadError instanceof Error && uploadError.message) {
        setError(uploadError.message)
      } else {
        setError('Failed to upload workbook. Please try again.')
      }
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
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setError('Invalid file type. Please upload a .xlsx sales workbook.')
      event.target.value = ''
      return
    }

    setIsLoading(true)
    setError('')

    try {
      const arrayBuffer = await file.arrayBuffer()
      const parsed = parseSalesSpreadsheet(arrayBuffer)
      const activeWorkbook = await uploadActiveWorkbook('sales', file)
      setSalesTasks(parsed)
      setSalesFileName(activeWorkbook.fileName || file.name)
      setSalesManualOverrides({})
      setSalesEnabledResources({})
      setSalesSelectedProjects(new Set())
      setSalesProjectsInitialized(false)
      setIsSalesPivotCollapsed(true)
      setSalesCollapseResetToken((current) => current + 1)
    } catch (uploadError) {
      if (uploadError instanceof Error && uploadError.message) {
        setError(uploadError.message)
      } else {
        setError('Failed to upload sales workbook. Please try again.')
      }
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

  function areWeekCapacitySchedulesEqual(left: WeekCapacitySchedule, right: WeekCapacitySchedule): boolean {
    const leftKeys = Object.keys(left)
    const rightKeys = Object.keys(right)
    if (leftKeys.length !== rightKeys.length) {
      return false
    }
    return leftKeys.every((key) => right[key] === left[key])
  }

  function getEffectiveScheduledWeekCapacity(weekStartIso: string, schedule: WeekCapacitySchedule): number {
    let scheduledCapacity: number | null = null
    for (const weekIso of sortedAllWeekKeys) {
      if (weekIso > weekStartIso) {
        break
      }
      const explicitValue = schedule[weekIso]
      if (Number.isFinite(explicitValue)) {
        scheduledCapacity = explicitValue
      }
    }
    return scheduledCapacity ?? baseWeekCapacityBeforeAdjustments[weekStartIso] ?? 0
  }

  function normalizeWeekCapacityScheduleState(schedule: WeekCapacitySchedule): WeekCapacitySchedule {
    const normalized: WeekCapacitySchedule = {}

    Object.entries(schedule)
      .filter(([weekIso]) => !sortedAllWeekKeys.includes(weekIso))
      .sort(([left], [right]) => left.localeCompare(right))
      .forEach(([weekIso, value]) => {
        if (Number.isFinite(value) && value >= 0) {
          normalized[weekIso] = value
        }
      })

    let scheduledCapacity: number | null = null
    sortedAllWeekKeys.forEach((weekIso) => {
      if (!Object.prototype.hasOwnProperty.call(schedule, weekIso)) {
        return
      }
      const value = schedule[weekIso]
      if (!Number.isFinite(value) || value < 0) {
        return
      }
      const fallbackCapacity = scheduledCapacity ?? baseWeekCapacityBeforeAdjustments[weekIso] ?? 0
      if (value === fallbackCapacity) {
        return
      }
      normalized[weekIso] = value
      scheduledCapacity = value
    })

    return normalized
  }

  function handleSetWeekCapacityForWeek(weekStartIso: string, capacityHours: number): void {
    if (!sortedAllWeekKeys.includes(weekStartIso)) {
      return
    }
    if (!Number.isFinite(capacityHours) || capacityHours < 0) {
      return
    }

    setWeekCapacitySchedule((current) => {
      const weekIndex = sortedAllWeekKeys.indexOf(weekStartIso)
      if (weekIndex < 0) {
        return current
      }

      const nextSchedule: WeekCapacitySchedule = {
        ...current,
        [weekStartIso]: capacityHours,
      }

      const nextWeekIso = sortedAllWeekKeys[weekIndex + 1]
      if (nextWeekIso && !Object.prototype.hasOwnProperty.call(current, nextWeekIso)) {
        nextSchedule[nextWeekIso] = getEffectiveScheduledWeekCapacity(nextWeekIso, current)
      }

      const normalizedSchedule = normalizeWeekCapacityScheduleState(nextSchedule)
      return areWeekCapacitySchedulesEqual(current, normalizedSchedule) ? current : normalizedSchedule
    })
  }

  function handleSetWeekCapacityFromWeekForward(weekStartIso: string, capacityHours: number): void {
    if (!sortedAllWeekKeys.includes(weekStartIso)) {
      return
    }
    if (!Number.isFinite(capacityHours) || capacityHours < 0) {
      return
    }

    setWeekCapacitySchedule((current) => {
      const normalizedSchedule = normalizeWeekCapacityScheduleState({
        ...current,
        [weekStartIso]: capacityHours,
      })
      return areWeekCapacitySchedulesEqual(current, normalizedSchedule) ? current : normalizedSchedule
    })
  }

  function handleClearWeekCapacityEntry(weekStartIso: string): void {
    if (!sortedAllWeekKeys.includes(weekStartIso)) {
      return
    }

    setWeekCapacitySchedule((current) => {
      if (!Object.prototype.hasOwnProperty.call(current, weekStartIso)) {
        return current
      }
      const nextSchedule = { ...current }
      delete nextSchedule[weekStartIso]
      const normalizedSchedule = normalizeWeekCapacityScheduleState(nextSchedule)
      return areWeekCapacitySchedulesEqual(current, normalizedSchedule) ? current : normalizedSchedule
    })
  }

  function handleClearAllWeekCapacityScheduleEntries(): void {
    setWeekCapacitySchedule((current) => (Object.keys(current).length === 0 ? current : {}))
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
    setFilters(createDefaultFilters())
    setSelectedProjects(new Set(availableProjects))
    setSalesSelectedProjects(new Set(salesAvailableProjects))
    setSelectedSalesProbabilities(new Set(salesProbabilityOptions))
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
  function handleToggleSalesProbability(probability: number): void {
    setSelectedSalesProbabilities((current) => {
      const next = new Set(current)
      if (next.has(probability)) {
        next.delete(probability)
      } else {
        next.add(probability)
      }
      return next
    })
  }
  function handleSelectAllSalesProbabilities(): void {
    setSelectedSalesProbabilities(new Set(salesProbabilityOptions))
  }
  function handleClearSalesProbabilities(): void {
    setSelectedSalesProbabilities(new Set())
  }

  function handleRevenueRateChange(
    dataset: 'main' | 'sales',
    project: string,
    field: RevenueRateField,
    rawValue: string,
  ): void {
    const normalizedProject = project.trim()
    if (!normalizedProject) {
      return
    }
    const trimmedValue = rawValue.trim()
    const parsed = trimmedValue === '' ? 0 : Number(trimmedValue)
    if (!Number.isFinite(parsed) || parsed < 0) {
      return
    }
    const nextValue = Math.min(parsed, maxRatePerHour)
    const update = (current: RevenueRateMap): RevenueRateMap => {
      const currentEntry = current[normalizedProject] ?? { revenuePerHour: 0, grossProfitPerHour: 0 }
      const nextEntry = {
        ...currentEntry,
        [field]: nextValue,
      }
      const next = { ...current, [normalizedProject]: nextEntry }
      if (nextEntry.revenuePerHour === 0 && nextEntry.grossProfitPerHour === 0) {
        delete next[normalizedProject]
      }
      return next
    }
    if (dataset === 'main') {
      setMainRevenueRates(update)
      return
    }
    setSalesRevenueRates(update)
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

  const exportDepartmentWorkbook = (): void => {
    const workbook = XLSX.utils.book_new()
    const headers = ['Project', 'Sequence', 'Week', 'Assigned Hours', 'Percent Complete', 'Finish Date', 'Status']

    const autoWidth = (rows: Array<Array<string | number>>): Array<{ wch: number }> => {
      if (rows.length === 0) return []
      const widths: number[] = []
      rows.forEach((row) => {
        row.forEach((cell, idx) => {
          const len = String(cell ?? '').length
          widths[idx] = Math.max(widths[idx] ?? 0, Math.min(60, len + 2))
        })
      })
      return widths.map((wch) => ({ wch }))
    }

    const setHeaderRowFeatures = (sheet: XLSX.WorkSheet) => {
      const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
      sheet['!freeze'] = { xSplit: 0, ySplit: 1 }
      sheet['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: range.s.c },
          e: { r: 0, c: range.e.c },
        }),
      }
    }

    const resourcesForExport = ['Processing', 'Fabrication', 'Assembly', 'Paint', 'Shipping']
    resourcesForExport.forEach((resource) => {
      const rows = departmentRowsByResource[resource] ?? []
      const aoa: Array<Array<string | number>> = [
        headers,
        ...rows.map((row) => [
          row.project,
          row.sequence,
          row.weekLabel,
          Number(row.hours.toFixed(2)),
          Number(row.percentComplete.toFixed(1)),
          row.finishDate,
          row.status,
        ]),
      ]
      const sheet = XLSX.utils.aoa_to_sheet(aoa)
      sheet['!cols'] = autoWidth(aoa)
      setHeaderRowFeatures(sheet)
      XLSX.utils.book_append_sheet(workbook, sheet, resource)
    })

    const fileName = `Department_Export_${filters.year || 'all'}.xlsx`
    XLSX.writeFile(workbook, fileName)
  }

  function handleUnlock(event: React.FormEvent<HTMLFormElement>): void {
    event.preventDefault()
    if (passwordInput === APP_ADMIN_PASSWORD) {
      setIsUnlocked(true)
      setAccessRole('admin')
      setActivePage('executive')
      setPasswordError('')
      setPasswordInput('')
      window.sessionStorage.setItem(APP_UNLOCK_SESSION_KEY, 'true')
      window.sessionStorage.setItem(APP_ROLE_SESSION_KEY, 'admin')
      return
    }
    if (passwordInput === APP_USER_PASSWORD) {
      setIsUnlocked(true)
      setAccessRole('user')
      setActivePage('howToUse')
      setPasswordError('')
      setPasswordInput('')
      window.sessionStorage.setItem(APP_UNLOCK_SESSION_KEY, 'true')
      window.sessionStorage.setItem(APP_ROLE_SESSION_KEY, 'user')
      return
    }
    if (passwordInput === APP_FORECAST_PASSWORD) {
      setIsUnlocked(true)
      setAccessRole('forecast')
      setActivePage('report')
      setPasswordError('')
      setPasswordInput('')
      window.sessionStorage.setItem(APP_UNLOCK_SESSION_KEY, 'true')
      window.sessionStorage.setItem(APP_ROLE_SESSION_KEY, 'forecast')
      return
    }
    setPasswordError('Incorrect password. Please try again.')
  }

  function handleLock(): void {
    setIsUnlocked(false)
    setAccessRole(null)
    setActivePage('executive')
    setPasswordInput('')
    setPasswordError('')
    window.sessionStorage.removeItem(APP_UNLOCK_SESSION_KEY)
    window.sessionStorage.removeItem(APP_ROLE_SESSION_KEY)
  }

  const mainPlanningSaveLabel = formatPlanningSaveLabel(
    mainPlanningSaveStatus,
    mainPlanningUpdatedAt,
    mainPlanningSaveError,
  )
  const salesPlanningSaveLabel = formatPlanningSaveLabel(
    salesPlanningSaveStatus,
    salesPlanningUpdatedAt,
    salesPlanningSaveError,
  )
  const mainRevenueSaveLabel = formatRevenueSaveLabel(
    mainRevenueSaveStatus,
    mainRevenueUpdatedAt,
    mainRevenueSaveError,
  )
  const salesRevenueSaveLabel = formatRevenueSaveLabel(
    salesRevenueSaveStatus,
    salesRevenueUpdatedAt,
    salesRevenueSaveError,
  )
  const isUserMode = accessRole === 'user'
  const isForecastMode = accessRole === 'forecast'
  const canViewAdminPlanningControls = accessRole === 'admin'

  if (!isUnlocked) {
    return (
      <div className="lock-screen">
        <section className="panel lock-card">
          <div className="lock-logo-wrap">
            <img
              src="/brand/inframod-logo.svg"
              alt="InfraMOD"
              className="lock-logo"
              loading="eager"
              decoding="async"
            />
          </div>
          <h1>Capacity Dashboard</h1>
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
      <div className="brand-grid" aria-hidden />

      <header className="app-header">
        <div className="app-header-brand">
          <img
            src="/brand/inframod-logo.svg"
            alt="InfraMOD"
            className="app-header-logo"
            loading="eager"
            decoding="async"
          />
          <span className="app-header-wordmark">Capacity Dashboard</span>
        </div>

        <nav className="app-header-nav" aria-label="Main navigation">
          {pageTabs.map((tab) => (
            <button
              key={tab.key}
              type="button"
              className={activePage === tab.key ? 'page-tab page-tab-active' : 'page-tab'}
              onClick={() => setActivePage(tab.key)}
            >
              {tab.label}
            </button>
          ))}
        </nav>

        <div className="app-header-actions">
          {isUserMode && <span className="user-mode-badge">User Mode</span>}
          {isForecastMode && <span className="user-mode-badge">Forecast View</span>}
          <button type="button" className="ghost-btn lock-btn" onClick={handleLock}>
            Lock
          </button>
        </div>
      </header>

      {activePage === 'howToUse' && (
        <HowToUsePage
          isUserMode={isUserMode}
          onOpenReport={() => setActivePage('report')}
          onOpenProcessing={() => setActivePage('processing')}
          onOpenRevenue={() => setActivePage('revenue')}
          onOpenPlanning={() => setActivePage('planning')}
          onLock={handleLock}
        />
      )}

      {activePage === 'planning' && (
        <>
          <header className={`panel control-panel ${isHeaderCollapsed ? 'collapsed' : ''}`}>
            <div className="title-bar">
              <div className="brand-lockup">
                <img
                  src="/brand/inframod-logo.svg"
                  alt="InfraMOD"
                  className="brand-logo"
                  loading="lazy"
                  decoding="async"
                />
              </div>
              <div className="title-stack">
                <h1>Production Capacity Planning Dashboard</h1>
                <p className="subtitle">
                  Import forecast data, edit weekly hours in the planning pivot, and watch chart/table/capacity metrics
                  update instantly.
                </p>
              </div>
              <div className="title-actions">
                <button type="button" onClick={exportDepartmentWorkbook}>
                  Export Dept Workbook
                </button>
                {isUserMode && (
                  <span
                    style={{
                      padding: '6px 10px',
                      borderRadius: 999,
                      border: '1px solid #38bdf8',
                      color: '#e0f2fe',
                      fontSize: 12,
                      fontWeight: 600,
                    }}
                  >
                    User Mode
                  </span>
                )}
                <button type="button" className="ghost-btn lock-btn" onClick={handleLock}>
                  Lock
                </button>
                <button
                  type="button"
                  className="ghost-btn collapse-toggle"
                  onClick={() => setIsHeaderCollapsed((current) => !current)}
                  aria-expanded={!isHeaderCollapsed}
                >
                  <span className={`chevron ${isHeaderCollapsed ? 'chevron-closed' : 'chevron-open'}`} aria-hidden="true">
                    â¾
                  </span>
                  {isHeaderCollapsed ? 'Show Filters' : 'Hide Filters'}
                </button>
              </div>
            </div>

            {!isHeaderCollapsed && (
              <>
                <div className="controls-grid">
                  {canViewAdminPlanningControls && (
                    <>
                      <label>
                        Upload or Replace Workbook (.xlsx)
                        <input type="file" accept=".xlsx" onChange={handleUpload} />
                      </label>
                      <label>
                        Upload Sales Workbook (.xlsx)
                        <input type="file" accept=".xlsx" onChange={handleSalesUpload} />
                      </label>
                    </>
                  )}

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
                        const day = parseISO(iso).getDay()
                        if (day !== 0 && day !== 6) {
                          event.target.value = ''
                          return
                        }
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
                  <div className="meta-card meta-card-wide">
                    <strong>File</strong>
                    <span>{fileName}</span>
                  </div>
                  <div className="meta-card meta-card-wide">
                    <strong>Sales File</strong>
                    <span>{salesFileName}</span>
                  </div>
                  <div className="meta-card">
                    <strong>Planning Sync</strong>
                    <span
                      style={{
                        color:
                          mainPlanningSaveStatus === 'error'
                            ? '#fca5a5'
                            : mainPlanningSaveStatus === 'saved'
                              ? '#86efac'
                              : '#cbd5e1',
                      }}
                    >
                      {mainPlanningSaveLabel}
                    </span>
                  </div>
                  <div className="meta-card">
                    <strong>Sales Planning Sync</strong>
                    <span
                      style={{
                        color:
                          salesPlanningSaveStatus === 'error'
                            ? '#fca5a5'
                            : salesPlanningSaveStatus === 'saved'
                              ? '#86efac'
                              : '#cbd5e1',
                      }}
                    >
                      {salesPlanningSaveLabel}
                    </span>
                  </div>
                  <div className="meta-card">
                    <strong>Parsed Rows</strong>
                    <span>{tasks.length}</span>
                  </div>
                  <div className="meta-card">
                    <strong>Weeks in View</strong>
                    <span>{baseLayer.weekKeys.length}</span>
                  </div>
                  <div className="meta-card">
                    <strong>Enabled Resources</strong>
                    <span>{enabledResourceList.length}</span>
                  </div>
                  <div className="meta-card meta-card-wide">
                    <strong>Data Date Span</strong>
                    <span>{taskDateSpan.start && taskDateSpan.end ? `${taskDateSpan.start} to ${taskDateSpan.end}` : 'N/A'}</span>
                  </div>
                  <div className="meta-card meta-card-actions">
                    <button type="button" className="ghost-btn" onClick={resetFilters}>
                      Reset Filters
                    </button>
                    <button type="button" onClick={() => void exportReportExcel()}>
                      Export Report Excel
                    </button>
                  </div>
                </div>
              </>
            )}
          </header>

          {!isLoading && !error && allResourcesVisible && canViewAdminPlanningControls && (
            <ResourceCapacityTable
              key={`resource-capacity-${collapseResetToken}`}
              resources={resources}
              enabledResources={enabledResources}
              weeklyCapacitiesByResource={resourceWeeklyCapacities}
              weekKeys={allWeekKeys}
              totalWeekCapacitySchedule={weekCapacitySchedule}
              baseTotalCapacityByWeek={baseWeekCapacities}
              effectiveTotalCapacityByWeek={weekCapacities}
              onWeeklyCapacityChange={handleResourceWeeklyCapacityChange}
              onSetTotalWeekCapacityForWeek={handleSetWeekCapacityForWeek}
              onSetTotalWeekCapacityFromWeekForward={handleSetWeekCapacityFromWeekForward}
              onClearTotalWeekCapacityEntry={handleClearWeekCapacityEntry}
              onClearAllTotalWeekCapacityEntries={handleClearAllWeekCapacityScheduleEntries}
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

          {!isLoading && !error && canViewAdminPlanningControls && (weeklyBuckets.length > 0 || salesWeeklyBuckets.length > 0) && (
            <>
              {weeklyBuckets.length > 0 && (
                <PivotPlanningTable
                  model={pivotModel}
                  rowGrouping={pivotRowGrouping}
                  overCapacityWeeks={overCapacityWeeks}
                  visibleWeekKeys={visiblePivotWeekKeys}
                  weekWindowLabel={pivotWeekWindowLabel}
                  canPageBack={!showingAllPivotWeeks && safePivotStartIndex > 0}
                  canPageForward={
                    !showingAllPivotWeeks && safePivotStartIndex + effectivePivotWeekWindowSize < baseLayer.weekKeys.length
                  }
                  onPageBack={() =>
                    setPivotWeekStartIndex((current) => Math.max(0, current - effectivePivotWeekWindowSize))
                  }
                  onPageForward={() =>
                    setPivotWeekStartIndex((current) => Math.min(maxPivotStartIndex, current + effectivePivotWeekWindowSize))
                  }
                  weekWindowSize={pivotWeekWindowSize}
                  onWeekWindowSizeChange={(size) => {
                    if (!Number.isFinite(size) || (size !== -1 && size <= 0)) {
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
                    canPageBack={!showingAllSalesPivotWeeks && salesSafePivotStartIndex > 0}
                    canPageForward={
                      !showingAllSalesPivotWeeks &&
                      salesSafePivotStartIndex + effectiveSalesPivotWeekWindowSize < salesBaseLayer.weekKeys.length
                    }
                    onPageBack={() =>
                      setSalesPivotWeekStartIndex((current) =>
                        Math.max(0, current - effectiveSalesPivotWeekWindowSize),
                      )
                    }
                    onPageForward={() =>
                      setSalesPivotWeekStartIndex((current) =>
                        Math.min(salesMaxPivotStartIndex, current + effectiveSalesPivotWeekWindowSize),
                      )
                    }
                    weekWindowSize={salesPivotWeekWindowSize}
                    onWeekWindowSizeChange={(size) => {
                      if (!Number.isFinite(size) || (size !== -1 && size <= 0)) {
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
            </>
          )}

          {!isLoading && !error && isUserMode && (
            <div className="panel status">User mode is active. Planning edit sections are hidden.</div>
          )}

          {!allResourcesVisible && !isLoading && (
            <div className="panel status">No resources available in the current data scope.</div>
          )}
        </>
      )}

      {activePage === 'processing' && (
        <DepartmentPage
          resource="Processing"
          tasks={tasks}
          filters={filters}
          selectedProjects={selectedProjects}
          selectedWeekendDates={selectedWeekendDates}
          projectColors={projectColors}
          resourceEnabled={enabledResources['Processing'] !== false}
          filter={getDeptFilter('Processing')}
          onFilterChange={(next) => setDeptFilter('Processing', next)}
          progressMatcher={smartsheetProgressMatcher}
          smartsheetSyncLabel={smartsheetSyncLabel}
          onRefreshSmartsheet={() => void refreshSmartsheetProgress()}
          isSmartsheetSyncLoading={smartsheetSyncStatus === 'loading'}
        />
      )}
      {activePage === 'fabrication' && (
        <DepartmentPage
          resource="Fabrication"
          tasks={tasks}
          filters={filters}
          selectedProjects={selectedProjects}
          selectedWeekendDates={selectedWeekendDates}
          projectColors={projectColors}
          resourceEnabled={enabledResources['Fabrication'] !== false}
          filter={getDeptFilter('Fabrication')}
          onFilterChange={(next) => setDeptFilter('Fabrication', next)}
          progressMatcher={smartsheetProgressMatcher}
          smartsheetSyncLabel={smartsheetSyncLabel}
          onRefreshSmartsheet={() => void refreshSmartsheetProgress()}
          isSmartsheetSyncLoading={smartsheetSyncStatus === 'loading'}
        />
      )}
      {activePage === 'assembly' && (
        <DepartmentPage
          resource="Assembly"
          tasks={tasks}
          filters={filters}
          selectedProjects={selectedProjects}
          selectedWeekendDates={selectedWeekendDates}
          projectColors={projectColors}
          resourceEnabled={enabledResources['Assembly'] !== false}
          filter={getDeptFilter('Assembly')}
          onFilterChange={(next) => setDeptFilter('Assembly', next)}
          progressMatcher={smartsheetProgressMatcher}
          smartsheetSyncLabel={smartsheetSyncLabel}
          onRefreshSmartsheet={() => void refreshSmartsheetProgress()}
          isSmartsheetSyncLoading={smartsheetSyncStatus === 'loading'}
        />
      )}
      {activePage === 'paint' && (
        <DepartmentPage
          resource="Paint"
          tasks={tasks}
          filters={filters}
          selectedProjects={selectedProjects}
          selectedWeekendDates={selectedWeekendDates}
          projectColors={projectColors}
          resourceEnabled={enabledResources['Paint'] !== false}
          filter={getDeptFilter('Paint')}
          onFilterChange={(next) => setDeptFilter('Paint', next)}
          progressMatcher={smartsheetProgressMatcher}
          smartsheetSyncLabel={smartsheetSyncLabel}
          onRefreshSmartsheet={() => void refreshSmartsheetProgress()}
          isSmartsheetSyncLoading={smartsheetSyncStatus === 'loading'}
        />
      )}
      {activePage === 'shipping' && (
        <DepartmentPage
          resource="Shipping"
          tasks={tasks}
          filters={filters}
          selectedProjects={selectedProjects}
          selectedWeekendDates={selectedWeekendDates}
          projectColors={projectColors}
          resourceEnabled={enabledResources['Shipping'] !== false}
          filter={getDeptFilter('Shipping')}
          onFilterChange={(next) => setDeptFilter('Shipping', next)}
          progressMatcher={smartsheetProgressMatcher}
          smartsheetSyncLabel={smartsheetSyncLabel}
          onRefreshSmartsheet={() => void refreshSmartsheetProgress()}
          isSmartsheetSyncLoading={smartsheetSyncStatus === 'loading'}
        />
      )}
      {activePage === 'detailing' && (
        <DetailingPage
          tasks={tasks}
          filters={filters}
          selectedProjects={selectedProjects}
          projectColors={projectColors}
        />
      )}

      {activePage === 'revenue' && (
        <>
          {isLoading && <div className="panel status">Loading workbook...</div>}
          {!isLoading && error && <div className="panel status error">{error}</div>}
          {!isLoading && !error && (
            <RevenueWorkspace
              rateRows={revenueRateRows}
              weeklyRevenueRows={weeklyRevenueRows}
              weeklyProjectKeys={weeklyProjectKeys}
              weeklyGrossProfitRows={weeklyGrossProfitRows}
              weeklyGrossProfitProjectKeys={weeklyGrossProfitProjectKeys}
              monthlyRevenueRows={monthlyRevenueRows}
              monthlyGrossProfitRows={monthlyGrossProfitRows}
              monthlyProjectKeys={monthlyProjectKeys}
              maxRatePerHour={maxRatePerHour}
              onRateChange={handleRevenueRateChange}
              mainSaveLabel={mainRevenueSaveLabel}
              salesSaveLabel={salesRevenueSaveLabel}
              mainSaveStatus={mainRevenueSaveStatus}
              salesSaveStatus={salesRevenueSaveStatus}
            />
          )}
        </>
      )}

      {activePage === 'executive' && (
        <>
          {isLoading && <div className="panel status">Loading workbook...</div>}
          {!isLoading && error && <div className="panel status error">{error}</div>}
          {!isLoading && !error && (
            <ExecutiveSummary data={executiveData} />
          )}
        </>
      )}

            {activePage === 'report' && (
        <>
          {isLoading && <div className="panel status">Loading workbook...</div>}
          {!isLoading && error && <div className="panel status error">{error}</div>}
          {!isLoading && !error && allResourcesVisible && (
            <ReportWorkspace
              key={`report-workspace-${collapseResetToken}-${salesCollapseResetToken}`}
              weeklyBuckets={weeklyBuckets}
              combinedWeeklyBuckets={reportCombinedWeeklyBuckets}
              salesWeeklyBuckets={reportSalesWeeklyBuckets}
              salesMonthlyBuckets={salesMonthlyBuckets}
              combinedMonthlyBuckets={combinedMonthlyBuckets}
              monthlyBuckets={monthlyBuckets}
              categoryKeys={categoryKeys}
              combinedCategoryKeys={reportCombinedCategoryKeys}
              salesCategoryKeys={reportSalesCategoryKeys}
              projects={availableProjects}
              combinedProjects={reportCombinedProjects}
              salesProjects={reportSalesAvailableProjects}
              selectedProjects={selectedProjects}
              selectedCombinedProjects={reportCombinedSelectedProjects}
              selectedSalesProjects={reportSalesSelectedProjectsForCalc}
              onToggleProject={handleToggleProject}
              onToggleCombinedProject={handleToggleCombinedProject}
              onToggleSalesProject={handleToggleSalesProject}
              salesProbabilityOptions={salesProbabilityOptions}
              selectedSalesProbabilities={selectedSalesProbabilities}
              onToggleSalesProbability={handleToggleSalesProbability}
              onSelectAllSalesProbabilities={handleSelectAllSalesProbabilities}
              onClearSalesProbabilities={handleClearSalesProbabilities}
              hoveredProject={hoveredProject}
              onHoverProject={setHoveredProject}
              summaryMetrics={visibleSummaryMetrics}
              reportContext={reportContext}
              initialTab={initialReportTab}
              executiveData={executiveData}
              allowedTabs={isForecastMode ? ['snapshot', 'sales'] : undefined}
            />
          )}
          {!isLoading && !error && allResourcesVisible && (
            <section className="panel summary-panel">
              <div className="section-header">
                <h2>Summary</h2>
                <p>Key capacity metrics derived from the current planning dataset.</p>
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
          )}
          {!isLoading && !error && !allResourcesVisible && (
            <div className="panel status">No resources available in the current data scope.</div>
          )}
        </>
      )}
    </div>
  )
}
export default App
