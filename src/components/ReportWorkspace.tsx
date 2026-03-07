import { useState } from 'react'
import type { MonthlyBucket, WeeklyBucket } from '../types'
import type { SummaryMetric } from '../utils/reportExport'
import { ForecastChart } from './ForecastChart'
import { ForecastTable } from './ForecastTable'
import { MonthlyForecastTable } from './MonthlyForecastTable'

type ReportTab = 'snapshot' | 'weekly' | 'monthly' | 'summary'

interface ReportWorkspaceProps {
  weeklyBuckets: WeeklyBucket[]
  monthlyBuckets: MonthlyBucket[]
  categoryKeys: string[]
  projects: string[]
  selectedProjects: Set<string>
  onToggleProject: (project: string) => void
  summaryMetrics: SummaryMetric[]
}

export function ReportWorkspace({
  weeklyBuckets,
  monthlyBuckets,
  categoryKeys,
  projects,
  selectedProjects,
  onToggleProject,
  summaryMetrics,
}: ReportWorkspaceProps) {
  const [activeReportTab, setActiveReportTab] = useState<ReportTab>('snapshot')

  return (
    <section className="panel report-workspace">
      <div className="section-header section-header-row">
        <div>
          <h2>Report Workspace</h2>
          <p>Switch between snapshot, weekly, monthly, and summary views using the same live planning dataset.</p>
        </div>
        <div className="section-actions">
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
          Snapshot
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
      </div>

      {activeReportTab === 'snapshot' && (
        <div className="report-tab-panel">
          <div className="section-header">
            <h2>Weekly Capacity Snapshot</h2>
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
    </section>
  )
}
