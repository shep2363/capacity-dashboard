import { useEffect, useMemo, useRef, useState, type Dispatch, type SetStateAction } from 'react'
import type { MonthlyBucket, WeeklyBucket } from '../types'
import type { SummaryMetric } from '../utils/reportExport'
import { exportReportElementToPdf } from '../utils/exportReportPdf'
import { ForecastChart } from './ForecastChart'
import { ForecastTable } from './ForecastTable'
import { MonthlyForecastTable } from './MonthlyForecastTable'
import { ExecutiveSummary, type ExecutiveData } from './ExecutiveSummary'
import {
  fetchDealFieldKeys,
  fetchPipedriveDeals,
  fetchPipedriveStages,
  type PipedriveDeal,
} from '../utils/pipedrive'

export type ReportTab =
  | 'snapshot'
  | 'weekly'
  | 'monthly'
  | 'summary'
  | 'sales'
  | 'combined'
  | 'sales-monthly'
  | 'combined-monthly'
  | 'executive'

interface ReportWorkspaceProps {
  weeklyBuckets: WeeklyBucket[]
  salesWeeklyBuckets: WeeklyBucket[]
  combinedWeeklyBuckets: WeeklyBucket[]
  monthlyBuckets: MonthlyBucket[]
  salesMonthlyBuckets: MonthlyBucket[]
  combinedMonthlyBuckets: MonthlyBucket[]
  categoryKeys: string[]
  salesCategoryKeys: string[]
  combinedCategoryKeys: string[]
  projects: string[]
  salesProjects: string[]
  combinedProjects: string[]
  selectedProjects: Set<string>
  selectedSalesProjects: Set<string>
  selectedCombinedProjects: Set<string>
  onToggleProject: (project: string) => void
  onToggleSalesProject: (project: string) => void
  onToggleCombinedProject: (project: string) => void
  salesProbabilityOptions: number[]
  selectedSalesProbabilities: Set<number>
  onToggleSalesProbability: (probability: number) => void
  onSelectAllSalesProbabilities: () => void
  onClearSalesProbabilities: () => void
  hoveredProject: string | null
  onHoverProject: (project: string | null) => void
  summaryMetrics: SummaryMetric[]
  reportContext: string[]
  initialTab?: ReportTab
  executiveData: ExecutiveData
}

interface SelectedWeekSummary {
  count: number
  weekLabels: string[]
  projectHours: number
  capacity: number
  overCapacity: number
  underCapacity: number
}

function buildSelectedWeekSummary(
  weeklyBuckets: WeeklyBucket[],
  selectedWeekIds: Set<string>,
): SelectedWeekSummary | null {
  const selectedBuckets = weeklyBuckets.filter((bucket) => selectedWeekIds.has(bucket.weekStartIso))
  if (selectedBuckets.length === 0) {
    return null
  }

  const totals = selectedBuckets.reduce(
    (acc, bucket) => {
      const overCapacity = Math.max(bucket.totalHours - bucket.capacity, 0)
      const underCapacity = Math.max(bucket.capacity - bucket.totalHours, 0)
      acc.projectHours += bucket.totalHours
      acc.capacity += bucket.capacity
      acc.overCapacity += overCapacity
      acc.underCapacity += underCapacity
      return acc
    },
    { projectHours: 0, capacity: 0, overCapacity: 0, underCapacity: 0 },
  )

  return {
    count: selectedBuckets.length,
    weekLabels: selectedBuckets.map((bucket) => bucket.weekLabel),
    ...totals,
  }
}

function formatProbabilityLabel(probability: number): string {
  return `${Number.isInteger(probability) ? probability.toFixed(0) : probability.toFixed(1)}%`
}

function buildWeekSelectionRange(
  bucketList: WeeklyBucket[],
  anchorWeekStartIso: string | null,
  targetWeekStartIso: string,
): string[] {
  const orderedWeeks = bucketList.map((bucket) => bucket.weekStartIso)
  const targetIndex = orderedWeeks.indexOf(targetWeekStartIso)
  if (targetIndex < 0) {
    return []
  }

  const anchorIndex = anchorWeekStartIso ? orderedWeeks.indexOf(anchorWeekStartIso) : -1
  if (anchorIndex < 0) {
    return [targetWeekStartIso]
  }

  const startIndex = Math.min(anchorIndex, targetIndex)
  const endIndex = Math.max(anchorIndex, targetIndex)
  return orderedWeeks.slice(startIndex, endIndex + 1)
}

export function ReportWorkspace({
  weeklyBuckets,
  salesWeeklyBuckets,
  combinedWeeklyBuckets,
  monthlyBuckets,
  salesMonthlyBuckets,
  combinedMonthlyBuckets,
  categoryKeys,
  salesCategoryKeys,
  combinedCategoryKeys,
  projects,
  salesProjects,
  combinedProjects,
  selectedProjects,
  selectedSalesProjects,
  selectedCombinedProjects,
  onToggleProject,
  onToggleSalesProject,
  onToggleCombinedProject,
  salesProbabilityOptions,
  selectedSalesProbabilities,
  onToggleSalesProbability,
  onSelectAllSalesProbabilities,
  onClearSalesProbabilities,
  hoveredProject,
  onHoverProject,
  summaryMetrics,
  reportContext,
  initialTab = 'snapshot',
  executiveData,
}: ReportWorkspaceProps) {
  const reportRef = useRef<HTMLElement>(null)
  const [activeReportTab, setActiveReportTab] = useState<ReportTab>(initialTab)
  const [isExportingPdf, setIsExportingPdf] = useState(false)
  const [pdfError, setPdfError] = useState('')
  const [deals, setDeals] = useState<PipedriveDeal[]>([])
  const [dealStage, setDealStage] = useState('all')
  const [dealProbBucket, setDealProbBucket] = useState('all')
  const [dealsLoading, setDealsLoading] = useState(false)
  const [dealsError, setDealsError] = useState('')
  const [selectedShopWeekIds, setSelectedShopWeekIds] = useState<Set<string>>(new Set())
  const [selectedSalesWeekIds, setSelectedSalesWeekIds] = useState<Set<string>>(new Set())
  const [selectedCombinedWeekIds, setSelectedCombinedWeekIds] = useState<Set<string>>(new Set())
  const selectedShopWeekAnchorRef = useRef<string | null>(null)
  const selectedSalesWeekAnchorRef = useRef<string | null>(null)
  const selectedCombinedWeekAnchorRef = useRef<string | null>(null)
  const [stageMap, setStageMap] = useState<Record<number, string>>({})
  const envToken =
    (import.meta.env as Record<string, string | undefined>).VITE_PROJECT_47_API_TOKEN ??
    (import.meta.env as Record<string, string | undefined>).VITE_PIPEDRIVE_API_TOKEN ??
    ''
  const hoursFieldKeys: Record<string, string | undefined> = {
    fab:
      (import.meta.env as Record<string, string | undefined>).VITE_PIPEDRIVE_FAB_HOURS_KEY ??
      (import.meta.env as Record<string, string | undefined>).VITE_PROJECT_47_FAB_HOURS_KEY,
    blast:
      (import.meta.env as Record<string, string | undefined>).VITE_PIPEDRIVE_BLAST_HOURS_KEY ??
      (import.meta.env as Record<string, string | undefined>).VITE_PROJECT_47_BLAST_HOURS_KEY,
    paint:
      (import.meta.env as Record<string, string | undefined>).VITE_PIPEDRIVE_PAINT_HOURS_KEY ??
      (import.meta.env as Record<string, string | undefined>).VITE_PROJECT_47_PAINT_HOURS_KEY,
    ship:
      (import.meta.env as Record<string, string | undefined>).VITE_PIPEDRIVE_SHIP_HOURS_KEY ??
      (import.meta.env as Record<string, string | undefined>).VITE_PROJECT_47_SHIP_HOURS_KEY,
  }
  const [pipedriveToken, setPipedriveToken] = useState<string>(() => {
    if (envToken) return envToken
    if (typeof window !== 'undefined') {
      return window.localStorage.getItem('pipedrive_api_token') ?? ''
    }
    return ''
  })
  const [tokenInput, setTokenInput] = useState<string>(() => pipedriveToken)

  async function handleExportPdf(): Promise<void> {
    if (!reportRef.current) {
      return
    }

    setIsExportingPdf(true)
    setPdfError('')
    try {
      const dateStamp = new Date().toISOString().slice(0, 10)
      await exportReportElementToPdf(reportRef.current, {
        fileName: `capacity-report-${dateStamp}.pdf`,
      })
    } catch {
      setPdfError('Unable to generate PDF. Please try again.')
    } finally {
      setIsExportingPdf(false)
    }
  }

  useEffect(() => {
    if (!pipedriveToken) {
      setDealsError('Add a Pipedrive API token to load deals.')
      return
    }

    const controller = new AbortController()
    setDealsLoading(true)
    setDealsError('')

    const resolveHoursKeys = async (): Promise<Record<string, string | undefined>> => {
      const provided = hoursFieldKeys
      const missingKeys = Object.values(provided).every((v) => !v)
      if (missingKeys) {
        try {
          const auto = await fetchDealFieldKeys(pipedriveToken, controller.signal)
          return { fab: auto.fab, blast: auto.blast, paint: auto.paint, ship: auto.ship }
        } catch {
          return provided
        }
      }
      return provided
    }

    Promise.all([
      fetchPipedriveStages(pipedriveToken, controller.signal).catch(() => ({} as Record<number, string>)),
      resolveHoursKeys().then((resolvedHours) =>
        fetchPipedriveDeals(pipedriveToken, { signal: controller.signal, hoursFieldKeys: resolvedHours }),
      ),
    ])
      .then(([stages, data]) => {
        const stageLookup = stages as Record<number, string>
        setStageMap(stageLookup)
        const normalized = data.map((deal) => {
          const stageName = deal.stage_name || (deal.stage_id != null ? stageLookup[deal.stage_id] : undefined)
          return stageName ? { ...deal, stage_name: stageName } : deal
        })
        setDeals(normalized)
      })
      .catch((error) => {
        if (controller.signal.aborted) return
        const message = error instanceof Error ? error.message : 'Failed to load Pipedrive deals.'
        setDealsError(message)
      })
      .finally(() => {
        if (!controller.signal.aborted) {
          setDealsLoading(false)
        }
      })

    return () => controller.abort()
  }, [pipedriveToken])

  function handleSaveToken(): void {
    const trimmed = tokenInput.trim()
    if (!trimmed) {
      setDealsError('Please paste a valid Pipedrive API token.')
      return
    }
    if (typeof window !== 'undefined') {
      window.localStorage.setItem('pipedrive_api_token', trimmed)
    }
    setDealsError('')
    setPipedriveToken(trimmed)
  }

  function pruneSelectedWeeks(current: Set<string>, bucketList: WeeklyBucket[]): Set<string> {
    const validWeeks = new Set(bucketList.map((bucket) => bucket.weekStartIso))
    const next = new Set([...current].filter((weekId) => validWeeks.has(weekId)))
    if (next.size === current.size && [...next].every((weekId) => current.has(weekId))) {
      return current
    }
    return next
  }

  function toggleSelectedWeeks(
    setter: Dispatch<SetStateAction<Set<string>>>,
    anchorRef: { current: string | null },
    bucketList: WeeklyBucket[],
    weekStartIso: string,
    selection: { multiSelect: boolean; rangeSelect: boolean },
  ): void {
    const { multiSelect, rangeSelect } = selection
    setter((current) => {
      if (rangeSelect) {
        const rangeWeekIds = buildWeekSelectionRange(bucketList, anchorRef.current, weekStartIso)
        anchorRef.current = weekStartIso

        if (!multiSelect) {
          const next = new Set(rangeWeekIds)
          if (next.size === current.size && [...next].every((weekId) => current.has(weekId))) {
            return current
          }
          return next
        }

        const next = new Set(current)
        rangeWeekIds.forEach((weekId) => next.add(weekId))
        if (next.size === current.size && [...next].every((weekId) => current.has(weekId))) {
          return current
        }
        return next
      }

      anchorRef.current = weekStartIso
      if (!multiSelect) {
        if (current.size === 1 && current.has(weekStartIso)) {
          return current
        }
        return new Set([weekStartIso])
      }

      const next = new Set(current)
      if (next.has(weekStartIso)) {
        next.delete(weekStartIso)
      } else {
        next.add(weekStartIso)
      }
      return next
    })
  }

  function renderWeekSelectionSummary(
    summary: SelectedWeekSummary | null,
    onClear: () => void,
    emptyHint: string,
  ) {
    return (
      <div className="forecast-selection-summary">
        {summary ? (
          <>
            <div className="forecast-selection-summary-header">
              <div>
                <strong>{summary.count} selected week{summary.count === 1 ? '' : 's'}</strong>
                <p>{summary.weekLabels.join(', ')}</p>
              </div>
              <button type="button" className="ghost-btn" onClick={onClear}>
                Clear Selection
              </button>
            </div>
            <div className="forecast-selection-summary-grid">
              <div>
                <span>Total Project Hours</span>
                <strong>{summary.projectHours.toFixed(2)}</strong>
              </div>
              <div>
                <span>Total Capacity</span>
                <strong>{summary.capacity.toFixed(2)}</strong>
              </div>
              <div>
                <span>Total Over Capacity</span>
                <strong className={summary.overCapacity > 0 ? 'warning' : ''}>{summary.overCapacity.toFixed(2)}</strong>
              </div>
              <div>
                <span>Total Under Capacity</span>
                <strong className={summary.underCapacity > 0 ? 'negative' : ''}>
                  {summary.underCapacity.toFixed(2)}
                </strong>
              </div>
            </div>
          </>
        ) : (
          <p className="forecast-selection-hint">{emptyHint}</p>
        )}
      </div>
    )
  }

  function renderSalesProbabilityFilter(scopeLabel: string) {
    return (
      <div className="forecast-filter-panel">
        <div className="forecast-filter-header">
          <div>
            <strong>Sales Probability Filter</strong>
            <p>
              Choose which Sales Production Report probabilities to include in the {scopeLabel} chart.
              {selectedSalesProbabilities.size === 0 ? ' No probabilities selected = no sales data shown.' : ''}
            </p>
          </div>
          <div className="forecast-filter-actions">
            <button type="button" className="ghost-btn" onClick={onSelectAllSalesProbabilities}>
              Select All
            </button>
            <button type="button" className="ghost-btn" onClick={onClearSalesProbabilities}>
              Clear All
            </button>
          </div>
        </div>
        {salesProbabilityOptions.length > 0 ? (
          <>
            <div className="toggle-chips">
              {salesProbabilityOptions.map((probability) => {
                const isSelected = selectedSalesProbabilities.has(probability)
                return (
                  <button
                    key={probability}
                    type="button"
                    className={`chip-toggle ${isSelected ? 'chip-on' : 'chip-off'}`}
                    onClick={() => onToggleSalesProbability(probability)}
                    aria-pressed={isSelected}
                  >
                    {formatProbabilityLabel(probability)}
                  </button>
                )
              })}
            </div>
            <p className="forecast-filter-note">
              {selectedSalesProbabilities.size} of {salesProbabilityOptions.length} probabilities selected
            </p>
          </>
        ) : (
          <p className="forecast-filter-note">No probability values were found in the current Sales workbook.</p>
        )}
      </div>
    )
  }

  useEffect(() => {
    setSelectedShopWeekIds((current) => pruneSelectedWeeks(current, weeklyBuckets))
  }, [weeklyBuckets])

  useEffect(() => {
    setSelectedSalesWeekIds((current) => pruneSelectedWeeks(current, salesWeeklyBuckets))
  }, [salesWeeklyBuckets])

  useEffect(() => {
    setSelectedCombinedWeekIds((current) => pruneSelectedWeeks(current, combinedWeeklyBuckets))
  }, [combinedWeeklyBuckets])

  function clearSelectedWeeks(
    setter: Dispatch<SetStateAction<Set<string>>>,
    anchorRef: { current: string | null },
  ): void {
    anchorRef.current = null
    setter(new Set())
  }

  function handleSelectShopWeek(
    weekStartIso: string,
    selection: { multiSelect: boolean; rangeSelect: boolean },
  ): void {
    toggleSelectedWeeks(setSelectedShopWeekIds, selectedShopWeekAnchorRef, weeklyBuckets, weekStartIso, selection)
  }

  function handleSelectSalesWeek(
    weekStartIso: string,
    selection: { multiSelect: boolean; rangeSelect: boolean },
  ): void {
    toggleSelectedWeeks(setSelectedSalesWeekIds, selectedSalesWeekAnchorRef, salesWeeklyBuckets, weekStartIso, selection)
  }

  function handleSelectCombinedWeek(
    weekStartIso: string,
    selection: { multiSelect: boolean; rangeSelect: boolean },
  ): void {
    toggleSelectedWeeks(
      setSelectedCombinedWeekIds,
      selectedCombinedWeekAnchorRef,
      combinedWeeklyBuckets,
      weekStartIso,
      selection,
    )
  }

  const selectedShopWeekSummary = useMemo(
    () => buildSelectedWeekSummary(weeklyBuckets, selectedShopWeekIds),
    [weeklyBuckets, selectedShopWeekIds],
  )

  const selectedSalesWeekSummary = useMemo(
    () => buildSelectedWeekSummary(salesWeeklyBuckets, selectedSalesWeekIds),
    [salesWeeklyBuckets, selectedSalesWeekIds],
  )

  const selectedCombinedWeekSummary = useMemo(
    () => buildSelectedWeekSummary(combinedWeeklyBuckets, selectedCombinedWeekIds),
    [combinedWeeklyBuckets, selectedCombinedWeekIds],
  )

  const probabilityLabels: Record<string, string> = {
    all: 'All probabilities',
    '100': '100%',
    '70-100': '70-100%',
    '75-99': '75-99%',
    '50-74': '50-74%',
    '25-49': '25-49%',
    '1-24': '1-24%',
    '0': '0%',
  }

  function matchesProbability(bucket: string, probability?: number | null): boolean {
    if (bucket === 'all') return true
    if (probability == null) return false
    switch (bucket) {
      case '100':
        return probability === 100
      case '70-100':
        return probability >= 70 && probability <= 100
      case '75-99':
        return probability >= 75 && probability <= 99
      case '50-74':
        return probability >= 50 && probability <= 74
      case '25-49':
        return probability >= 25 && probability <= 49
      case '1-24':
        return probability >= 1 && probability <= 24
      case '0':
        return probability === 0
      default:
        return true
    }
  }

  const stageOptions = [
    'all',
    ...new Set(
      deals
        .map((deal) => deal.stage_name || (deal.stage_id != null ? stageMap[deal.stage_id] : ''))
        .filter(Boolean) as string[],
    ),
  ]
  const filteredDeals = deals
    .filter((deal) => (dealStage === 'all' ? true : deal.stage_name === dealStage))
    .filter((deal) => matchesProbability(dealProbBucket, deal.probability))
    .sort((a, b) => {
      const probA = a.probability ?? -1
      const probB = b.probability ?? -1
      if (probA !== probB) return probB - probA
      const valA = a.value ?? 0
      const valB = b.value ?? 0
      return valB - valA
    })

  function renderDealsPanel(title: string) {
    return (
      <div className="panel table-panel">
        <div className="section-header">
          <div>
            <h3>{title}</h3>
            <p>Live Pipedrive deals filtered by Stage and Probability.</p>
          </div>
          <div className="deals-filter-row">
            <label className="control-inline">
              Stage
              <select value={dealStage} onChange={(event) => setDealStage(event.target.value)}>
                {stageOptions.map((stage) => (
                  <option key={stage} value={stage}>
                    {stage === 'all' ? 'All stages' : stage}
                  </option>
                ))}
              </select>
            </label>
            <label className="control-inline">
              Probability
              <select value={dealProbBucket} onChange={(event) => setDealProbBucket(event.target.value)}>
                {Object.entries(probabilityLabels).map(([key, label]) => (
                  <option key={key} value={key}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
            {dealsLoading && <span className="pill">Loading deals...</span>}
            {dealsError && <span className="error-text">{dealsError}</span>}
            {!pipedriveToken && (
              <div className="token-inline">
                <input
                  type="password"
                  placeholder="Paste Pipedrive API token"
                  value={tokenInput}
                  onChange={(event) => setTokenInput(event.target.value)}
                />
                <button type="button" onClick={handleSaveToken}>
                  Use Token
                </button>
                <small>Saved to this browser only.</small>
              </div>
            )}
          </div>
        </div>

        <div className="table-wrap">
          <table className="report-summary-table">
            <thead>
              <tr>
                <th>Title</th>
                <th>Organization</th>
                <th>Stage</th>
                <th>Value</th>
                <th>Fab Hours</th>
                <th>Blast Hours</th>
                <th>Paint Hours</th>
                <th>Ship/Handling Hours</th>
                <th>Probability</th>
              </tr>
            </thead>
            <tbody>
              {filteredDeals.length === 0 ? (
                <tr>
                  <td colSpan={7}>{dealsLoading ? 'Loading deals...' : 'No deals match the current filters.'}</td>
                </tr>
              ) : (
                filteredDeals.slice(0, 100).map((deal) => (
                  <tr key={deal.id}>
                    <td>{deal.title}</td>
                    <td>{deal.org_name || '—'}</td>
                    <td>{deal.stage_name || '—'}</td>
                    <td>{deal.value != null ? deal.value.toLocaleString() : '—'}</td>
                    {(['fab', 'blast', 'paint', 'ship'] as const).map((key) => (
                      <td key={`${deal.id}-${key}`}>
                        {deal.hours && deal.hours[key] != null && !Number.isNaN(deal.hours[key] as number)
                          ? (deal.hours[key] as number).toLocaleString(undefined, { maximumFractionDigits: 2 })
                          : '—'}
                      </td>
                    ))}
                    <td>{deal.probability != null ? `${deal.probability}%` : '—'}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    )
  }

  return (
    <section ref={reportRef} className="panel report-workspace">
      <div className="section-header section-header-row report-header">
        <div className="report-header-main">
          <div className="report-header-copy">
            <h2>Report Workspace</h2>
            <p>Review charts and forecast tables using the current planning filters, selected projects, and live edits.</p>
          </div>
          <div className="report-context">
            {reportContext.map((line) => (
              <span key={line}>{line}</span>
            ))}
          </div>
        </div>
        <div className="section-actions">
          <button type="button" className="ghost-btn" disabled={isExportingPdf} onClick={() => void handleExportPdf()}>
            {isExportingPdf ? 'Generating PDF...' : 'Export Report PDF'}
          </button>
          <button type="button" className="ghost-btn" onClick={() => window.print()}>
            Print Report
          </button>
        </div>
      </div>

      <div className="report-tabs" aria-label="Report Views">
        <button
          type="button"
          className={activeReportTab === 'snapshot' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('snapshot')}
        >
          Shop Forecast
        </button>
        <button
          type="button"
          className={activeReportTab === 'weekly' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('weekly')}
        >
          Weekly Forecast
        </button>
        <button
          type="button"
          className={activeReportTab === 'monthly' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('monthly')}
        >
          Monthly Forecast
        </button>
        <button
          type="button"
          className={activeReportTab === 'summary' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('summary')}
        >
          Summary
        </button>
        <button
          type="button"
          className={activeReportTab === 'executive' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('executive')}
        >
          Executive Summary
        </button>
        <button
          type="button"
          className={activeReportTab === 'sales' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('sales')}
        >
          Sales Forecast
        </button>
        <button
          type="button"
          className={activeReportTab === 'combined' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('combined')}
        >
          Shop and Sales Forecast
        </button>
        <button
          type="button"
          className={activeReportTab === 'sales-monthly' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('sales-monthly')}
        >
          Sales Monthly Forecast
        </button>
        <button
          type="button"
          className={activeReportTab === 'combined-monthly' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveReportTab('combined-monthly')}
        >
          Shop and Sales Monthly Forecast
        </button>
      </div>

      {activeReportTab === 'executive' && (
        <div className="report-tab-panel">
          <ExecutiveSummary data={executiveData} />
        </div>
      )}

      {activeReportTab === 'snapshot' && (
        <div className="report-tab-panel">
          <div className="section-header">
            <h2>Shop Forecast</h2>
            <p>Stacked weekly hours with capacity line using your current filters, capacities, and manual edits.</p>
          </div>
          {renderWeekSelectionSummary(
            selectedShopWeekSummary,
            () => clearSelectedWeeks(setSelectedShopWeekIds, selectedShopWeekAnchorRef),
            'Ctrl + click week bars or labels to toggle weeks. Ctrl + Shift + click adds the full week range in between.',
          )}
          <ForecastChart
            weeklyBuckets={weeklyBuckets}
            categoryKeys={categoryKeys}
            projects={projects}
            selectedProjects={selectedProjects}
            onToggleProject={onToggleProject}
            hoveredProject={hoveredProject}
            onHoverProject={onHoverProject}
            selectedWeekIds={selectedShopWeekIds}
            onWeekSelect={handleSelectShopWeek}
          />
        </div>
      )}

      {activeReportTab === 'sales' && (
        <div className="report-tab-panel">
          <div className="section-header">
            <h2>Weekly Sales Forecast</h2>
            <p>Stacked weekly sales forecast with capacity line using current filters and edits.</p>
          </div>
          {renderSalesProbabilityFilter('Sales Forecast')}
          {renderWeekSelectionSummary(
            selectedSalesWeekSummary,
            () => clearSelectedWeeks(setSelectedSalesWeekIds, selectedSalesWeekAnchorRef),
            'Ctrl + click week bars or labels to toggle weeks. Ctrl + Shift + click adds the full week range in between.',
          )}
          <ForecastChart
            weeklyBuckets={salesWeeklyBuckets}
            categoryKeys={salesCategoryKeys}
            projects={salesProjects}
            selectedProjects={selectedSalesProjects}
            onToggleProject={onToggleSalesProject}
            title="Weekly Sales Forecast"
            subtitle="Stacked weekly sales forecast hours with capacity overlay."
            hoveredProject={hoveredProject}
            onHoverProject={onHoverProject}
            hoverProjectPrefix="Sales - "
            selectedWeekIds={selectedSalesWeekIds}
            onWeekSelect={handleSelectSalesWeek}
          />
          {renderDealsPanel('Pipedrive Deals — Sales Forecast')}
        </div>
      )}

      {activeReportTab === 'combined' && (
        <div className="report-tab-panel">
          <div className="section-header">
            <h2>Sales & Capacity</h2>
            <p>View sales forecast alongside operational capacity for the current scope.</p>
          </div>
          {renderSalesProbabilityFilter('Shop and Sales Forecast')}
          {renderWeekSelectionSummary(
            selectedCombinedWeekSummary,
            () => clearSelectedWeeks(setSelectedCombinedWeekIds, selectedCombinedWeekAnchorRef),
            'Ctrl + click week bars or labels to toggle weeks. Ctrl + Shift + click adds the full week range in between.',
          )}
          <ForecastChart
            weeklyBuckets={combinedWeeklyBuckets}
            categoryKeys={combinedCategoryKeys}
            projects={combinedProjects}
            selectedProjects={selectedCombinedProjects}
            onToggleProject={onToggleCombinedProject}
            title="Combined Weekly Forecast"
            subtitle="Operational + Sales projects together with capacity overlay."
            hoveredProject={hoveredProject}
            onHoverProject={onHoverProject}
            selectedWeekIds={selectedCombinedWeekIds}
            onWeekSelect={handleSelectCombinedWeek}
          />
          {renderDealsPanel('Pipedrive Deals — Shop and Sales Forecast')}
        </div>
      )}

      {activeReportTab === 'weekly' && (
        <div className="report-tab-panel">
          <ForecastTable weeklyBuckets={weeklyBuckets} />
        </div>
      )}

      {activeReportTab === 'monthly' && (
        <div className="report-tab-panel">
          <MonthlyForecastTable monthlyBuckets={monthlyBuckets} />
        </div>
      )}

      {activeReportTab === 'sales-monthly' && (
        <div className="report-tab-panel">
          <MonthlyForecastTable monthlyBuckets={salesMonthlyBuckets} />
        </div>
      )}

      {activeReportTab === 'combined-monthly' && (
        <div className="report-tab-panel">
          <MonthlyForecastTable monthlyBuckets={combinedMonthlyBuckets} />
        </div>
      )}

      {activeReportTab === 'summary' && (
        <div className="report-tab-panel">
          <div className="panel table-panel">
            <div className="section-header">
              <h2>Summary Table</h2>
              <p>Top-level report metrics for the current workbook scope and selected filters.</p>
            </div>
            <div className="table-wrap">
              <table className="report-summary-table">
                <thead>
                  <tr>
                    <th>Metric</th>
                    <th>Value</th>
                  </tr>
                </thead>
                <tbody>
                  {summaryMetrics.map((item) => (
                    <tr key={item.metric}>
                      <td>{item.metric}</td>
                      <td>{item.value}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {pdfError && <p className="export-error">{pdfError}</p>}
    </section>
  )
}
