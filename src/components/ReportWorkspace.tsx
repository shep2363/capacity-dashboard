import { useRef, useState } from 'react'
import type { MonthlyBucket, WeeklyBucket } from '../types'
import type { SummaryMetric } from '../utils/reportExport'
import { exportReportElementToPdf } from '../utils/exportReportPdf'
import { ForecastChart } from './ForecastChart'
import { ForecastTable } from './ForecastTable'
import { MonthlyForecastTable } from './MonthlyForecastTable'

export type ReportTab = 'snapshot' | 'weekly' | 'monthly' | 'summary' | 'sales' | 'combined' | 'sales-monthly'

interface ReportWorkspaceProps {
  weeklyBuckets: WeeklyBucket[]
  salesWeeklyBuckets: WeeklyBucket[]
  combinedWeeklyBuckets: WeeklyBucket[]
  monthlyBuckets: MonthlyBucket[]
  salesMonthlyBuckets: MonthlyBucket[]
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
  summaryMetrics: SummaryMetric[]
  reportContext: string[]
  initialTab?: ReportTab
}

export function ReportWorkspace({
  weeklyBuckets,
  salesWeeklyBuckets,
  combinedWeeklyBuckets,
  monthlyBuckets,
  salesMonthlyBuckets,
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
  summaryMetrics,
  reportContext,
  initialTab = 'snapshot',
}: ReportWorkspaceProps) {
  const reportRef = useRef<HTMLElement>(null)
  const [activeReportTab, setActiveReportTab] = useState<ReportTab>(initialTab)
  const [isExportingPdf, setIsExportingPdf] = useState(false)
  const [pdfError, setPdfError] = useState('')

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
      </div>

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
          />
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
          />
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
