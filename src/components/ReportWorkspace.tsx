import { useEffect, useRef, useState } from 'react'
import type { MonthlyBucket, WeeklyBucket } from '../types'
import type { SummaryMetric } from '../utils/reportExport'
import { exportReportElementToPdf } from '../utils/exportReportPdf'
import { ForecastChart } from './ForecastChart'
import { ForecastTable } from './ForecastTable'
import { MonthlyForecastTable } from './MonthlyForecastTable'
import { ExecutiveSummary, type ExecutiveData } from './ExecutiveSummary'
import { fetchPipedriveDeals, type PipedriveDeal } from '../utils/pipedrive'

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
  hoveredProject: string | null
  onHoverProject: (project: string | null) => void
  summaryMetrics: SummaryMetric[]
  reportContext: string[]
  initialTab?: ReportTab
  executiveData: ExecutiveData
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
  const envToken =
    (import.meta.env as Record<string, string | undefined>).VITE_PROJECT_47_API_TOKEN ??
    (import.meta.env as Record<string, string | undefined>).VITE_PIPEDRIVE_API_TOKEN ??
    ''
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

    fetchPipedriveDeals(pipedriveToken, controller.signal)
      .then((data) => setDeals(data))
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

  const probabilityLabels: Record<string, string> = {
    all: 'All probabilities',
    '100': '100%',
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

  const stageOptions = ['all', ...new Set(deals.map((deal) => deal.stage_name).filter(Boolean) as string[])]
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
                <th>Probability</th>
              </tr>
            </thead>
            <tbody>
              {filteredDeals.length === 0 ? (
                <tr>
                  <td colSpan={5}>{dealsLoading ? 'Loading deals...' : 'No deals match the current filters.'}</td>
                </tr>
              ) : (
                filteredDeals.slice(0, 100).map((deal) => (
                  <tr key={deal.id}>
                    <td>{deal.title}</td>
                    <td>{deal.org_name || '—'}</td>
                    <td>{deal.stage_name || '—'}</td>
                    <td>{deal.value != null ? deal.value.toLocaleString() : '—'}</td>
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
      <div className="section-header section-header-row">
        <div>
          <h2>Report Workspace</h2>
          <p>Switch between snapshot, weekly, monthly, and summary views using the same live planning dataset.</p>
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
          <ForecastChart
            weeklyBuckets={weeklyBuckets}
            categoryKeys={categoryKeys}
            projects={projects}
            selectedProjects={selectedProjects}
            onToggleProject={onToggleProject}
            hoveredProject={hoveredProject}
            onHoverProject={onHoverProject}
          />
        </div>
      )}

      {activeReportTab === 'sales' && (
        <div className="report-tab-panel">
          <div className="section-header">
            <h2>Weekly Sales Forecast</h2>
            <p>Stacked weekly sales forecast with capacity line using current filters and edits.</p>
          </div>
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
