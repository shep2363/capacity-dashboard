import { ForecastChart } from './ForecastChart'
import { ForecastTable } from './ForecastTable'
import { MonthlyForecastTable } from './MonthlyForecastTable'
import type { MonthlyBucket, WeeklyBucket } from '../types'

interface ReportSnapshotProps {
  weeklyBuckets: WeeklyBucket[]
  monthlyBuckets: MonthlyBucket[]
  categoryKeys: string[]
  projects: string[]
  selectedProjects: Set<string>
  onToggleProject: (project: string) => void
}

export function ReportSnapshot({
  weeklyBuckets,
  monthlyBuckets,
  categoryKeys,
  projects,
  selectedProjects,
  onToggleProject,
}: ReportSnapshotProps) {
  return (
    <section className="panel report-snapshot" id="capacity-report-snapshot">
      <div className="section-header section-header-row">
        <div>
          <h2>Capacity Report Snapshot</h2>
          <p>Presentation-ready view of chart and forecast tables using current filters, selections, and edits.</p>
        </div>
        <div className="section-actions">
          <button type="button" className="ghost-btn print-report-btn" onClick={() => window.print()}>
            Print / Save PDF
          </button>
        </div>
      </div>

      <div className="report-chart-block">
        <ForecastChart
          weeklyBuckets={weeklyBuckets}
          categoryKeys={categoryKeys}
          projects={projects}
          selectedProjects={selectedProjects}
          onToggleProject={onToggleProject}
        />
      </div>

      <div className="report-tables-grid">
        <ForecastTable weeklyBuckets={weeklyBuckets} />
        <MonthlyForecastTable monthlyBuckets={monthlyBuckets} />
      </div>
    </section>
  )
}
