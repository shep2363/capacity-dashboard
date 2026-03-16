import { useMemo } from 'react'
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import type { WorkbookDataset } from '../utils/activeWorkbookApi'
import type { GrossProfitByProjectRow, RevenueRateRow, WeeklyRevenueRow } from '../utils/revenue'

const COLOR_PALETTE = [
  '#4f46e5',
  '#0f766e',
  '#f59e0b',
  '#0284c7',
  '#7c3aed',
  '#16a34a',
  '#9333ea',
  '#be123c',
  '#0ea5e9',
  '#64748b',
  '#ea580c',
  '#047857',
]

type RevenueSaveStatus = 'idle' | 'saving' | 'saved' | 'error'
type RevenueRateField = 'revenuePerHour' | 'grossProfitPerHour'

interface RevenueWorkspaceProps {
  rateRows: RevenueRateRow[]
  weeklyRevenueRows: WeeklyRevenueRow[]
  weeklyProjectKeys: string[]
  grossProfitRows: GrossProfitByProjectRow[]
  maxRatePerHour: number
  onRateChange: (dataset: WorkbookDataset, project: string, field: RevenueRateField, value: string) => void
  mainSaveLabel: string
  salesSaveLabel: string
  mainSaveStatus: RevenueSaveStatus
  salesSaveStatus: RevenueSaveStatus
}

function formatCurrency(value: number): string {
  if (!Number.isFinite(value)) {
    return '$0'
  }
  return new Intl.NumberFormat(undefined, {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  }).format(value)
}

function statusColor(status: RevenueSaveStatus): string {
  if (status === 'error') {
    return '#fca5a5'
  }
  if (status === 'saved') {
    return '#86efac'
  }
  return '#cbd5e1'
}

interface WeeklyTooltipEntry {
  payload?: WeeklyRevenueRow
}

function WeeklyRevenueTooltip({ active, payload }: { active?: boolean; payload?: WeeklyTooltipEntry[] }) {
  if (!active || !payload || payload.length === 0) {
    return null
  }

  const row = payload[0]?.payload
  if (!row) {
    return null
  }

  return (
    <div className="revenue-tooltip">
      <div className="revenue-tooltip-title">{row.weekRangeLabel}</div>
      <div className="revenue-tooltip-summary">
        <div>Total Project Hours: {row.totalPlannedHours.toFixed(1)}</div>
        <div>Total Revenue: {formatCurrency(row.totalRevenue)}</div>
      </div>
      <div className="revenue-tooltip-grid">
        {row.details.length === 0 ? (
          <div className="revenue-tooltip-empty">No planned hours for this week.</div>
        ) : (
          row.details.map((detail) => (
            <div key={`${row.weekStartIso}-${detail.projectLabel}`} className="revenue-tooltip-row">
              <div className="revenue-tooltip-project">{detail.projectLabel}</div>
              <div className="revenue-tooltip-metric">{detail.plannedHours.toFixed(1)} h</div>
              <div className="revenue-tooltip-metric">{formatCurrency(detail.revenuePerHour)}/h</div>
              <div className="revenue-tooltip-metric">{formatCurrency(detail.revenueAmount)}</div>
            </div>
          ))
        )}
      </div>
    </div>
  )
}

interface GrossProfitTooltipEntry {
  payload?: GrossProfitByProjectRow
}

function GrossProfitTooltip({ active, payload }: { active?: boolean; payload?: GrossProfitTooltipEntry[] }) {
  if (!active || !payload || payload.length === 0) {
    return null
  }
  const row = payload[0]?.payload
  if (!row) {
    return null
  }
  return (
    <div className="revenue-tooltip">
      <div className="revenue-tooltip-title">{row.projectLabel}</div>
      <div className="revenue-tooltip-summary">
        <div>Hours: {row.plannedHours.toFixed(1)}</div>
        <div>Rate: {formatCurrency(row.grossProfitPerHour)}/h</div>
        <div>Gross Profit: {formatCurrency(row.grossProfitAmount)}</div>
      </div>
    </div>
  )
}

export function RevenueWorkspace({
  rateRows,
  weeklyRevenueRows,
  weeklyProjectKeys,
  grossProfitRows,
  maxRatePerHour,
  onRateChange,
  mainSaveLabel,
  salesSaveLabel,
  mainSaveStatus,
  salesSaveStatus,
}: RevenueWorkspaceProps) {
  const weeklyChartData = useMemo(
    () =>
      weeklyRevenueRows.map((row) => ({
        ...row,
        ...row.amountsByProject,
      })),
    [weeklyRevenueRows],
  )

  return (
    <section className="panel revenue-page">
      <div className="section-header">
        <h2>Revenue</h2>
        <p>Set project financial rates and monitor weekly revenue and gross profit from current planning hours.</p>
      </div>

      <div className="revenue-status-row">
        <span>
          <strong>Shop Rates Sync:</strong>{' '}
          <span style={{ color: statusColor(mainSaveStatus) }}>{mainSaveLabel}</span>
        </span>
        <span>
          <strong>Sales Rates Sync:</strong>{' '}
          <span style={{ color: statusColor(salesSaveStatus) }}>{salesSaveLabel}</span>
        </span>
      </div>

      <div className="panel table-panel revenue-rate-editor">
        <div className="section-header">
          <h3>Rates by Project</h3>
          <p>Changes save to shared storage and are visible to all users.</p>
        </div>
        {rateRows.length === 0 ? (
          <div className="status">No projects are available in the current filter scope.</div>
        ) : (
          <div className="table-wrap">
            <table className="revenue-rate-table">
              <thead>
                <tr>
                  <th>Project</th>
                  <th>Planned Hours (Current Scope)</th>
                  <th>Revenue per Hour</th>
                  <th>Gross Profit per Hour</th>
                </tr>
              </thead>
              <tbody>
                {rateRows.map((row) => (
                  <tr key={`${row.dataset}-${row.project}`}>
                    <td>{row.label}</td>
                    <td>{row.plannedHours.toFixed(1)}</td>
                    <td>
                      <input
                        type="number"
                        min={0}
                        max={maxRatePerHour}
                        step="0.01"
                        value={row.revenuePerHour}
                        onChange={(event) => onRateChange(row.dataset, row.project, 'revenuePerHour', event.target.value)}
                      />
                    </td>
                    <td>
                      <input
                        type="number"
                        min={0}
                        max={maxRatePerHour}
                        step="0.01"
                        value={row.grossProfitPerHour}
                        onChange={(event) =>
                          onRateChange(row.dataset, row.project, 'grossProfitPerHour', event.target.value)
                        }
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      <div className="revenue-grid">
        <div className="panel chart-panel">
          <div className="section-header">
            <h3>Weekly Revenue</h3>
            <p>Calculated as planned hours multiplied by project revenue-per-hour rates.</p>
          </div>
          {weeklyChartData.length === 0 ? (
            <div className="status">No weekly revenue data available for the current scope.</div>
          ) : (
            <div className="chart-wrap revenue-chart-wrap">
              <ResponsiveContainer width="100%" height={470}>
                <BarChart data={weeklyChartData} margin={{ top: 16, right: 20, left: 20, bottom: 28 }}>
                  <CartesianGrid stroke="#334155" vertical={false} />
                  <XAxis dataKey="weekLabel" angle={-28} textAnchor="end" interval={0} height={62} />
                  <YAxis tickFormatter={(value: number) => formatCurrency(value)} />
                  <Tooltip content={<WeeklyRevenueTooltip />} />
                  <Legend />
                  {weeklyProjectKeys.map((projectKey, index) => (
                    <Bar
                      key={projectKey}
                      dataKey={projectKey}
                      name={projectKey}
                      stackId="revenue"
                      fill={COLOR_PALETTE[index % COLOR_PALETTE.length]}
                    />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>

        <div className="panel chart-panel">
          <div className="section-header">
            <h3>Gross Profit by Project</h3>
            <p>Calculated as planned hours multiplied by project gross-profit-per-hour rates.</p>
          </div>
          {grossProfitRows.length === 0 ? (
            <div className="status">No gross profit data available for the current scope.</div>
          ) : (
            <div className="chart-wrap revenue-chart-wrap">
              <ResponsiveContainer width="100%" height={470}>
                <BarChart data={grossProfitRows} margin={{ top: 16, right: 20, left: 20, bottom: 28 }}>
                  <CartesianGrid stroke="#334155" vertical={false} />
                  <XAxis dataKey="projectLabel" interval={0} angle={-28} textAnchor="end" height={80} />
                  <YAxis tickFormatter={(value: number) => formatCurrency(value)} />
                  <Tooltip content={<GrossProfitTooltip />} />
                  <Bar dataKey="grossProfitAmount" name="Gross Profit" fill="#22c55e" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      </div>
    </section>
  )
}
