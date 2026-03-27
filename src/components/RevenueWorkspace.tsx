import { useMemo, useState } from 'react'
import {
  Bar,
  CartesianGrid,
  ComposedChart,
  Legend,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import type { WorkbookDataset } from '../utils/activeWorkbookApi'
import { computeLeftTooltipPosition, type TooltipPosition } from '../utils/chartTooltip'
import { exportRevenueMonthlyWorkbook } from '../utils/revenueExport'
import type {
  MonthlyGrossProfitRow,
  MonthlyRevenueRow,
  RevenueRateRow,
  WeeklyGrossProfitRow,
  WeeklyRevenueRow,
} from '../utils/revenue'

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

const REVENUE_TOOLTIP_BOUNDS = { width: 450, height: 320 }

type RevenueSaveStatus = 'idle' | 'saving' | 'saved' | 'error'
type RevenueRateField = 'revenuePerHour' | 'grossProfitPerHour'
type RevenueViewTab = 'data' | 'revenue-chart' | 'gross-profit-chart' | 'revenue-monthly' | 'gross-profit-monthly'

interface RevenueWorkspaceProps {
  rateRows: RevenueRateRow[]
  weeklyRevenueRows: WeeklyRevenueRow[]
  weeklyProjectKeys: string[]
  weeklyGrossProfitRows: WeeklyGrossProfitRow[]
  weeklyGrossProfitProjectKeys: string[]
  monthlyRevenueRows: MonthlyRevenueRow[]
  monthlyGrossProfitRows: MonthlyGrossProfitRow[]
  monthlyProjectKeys: string[]
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

function formatHours(value: unknown): string {
  const normalized = Array.isArray(value) ? value[0] : value
  const numeric = typeof normalized === 'number' ? normalized : Number(normalized ?? 0)
  return `${numeric.toFixed(1)} h`
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

function truncateLabel(label: string): string {
  return label.length > 26 ? `${label.slice(0, 26)}...` : label
}

function resolveBarSize(itemCount: number): number {
  if (itemCount > 80) return 14
  if (itemCount > 60) return 17
  if (itemCount > 40) return 20
  if (itemCount > 24) return 24
  return 30
}

function buildRevenueExportFileName(): string {
  const timestamp = new Date().toISOString().replace(/\.\d{3}Z$/, '').replace(/[:]/g, '').replace('T', '-')
  return `revenue-monthly-forecast-${timestamp}.xlsx`
}

function CompactLegend({
  payload,
}: {
  payload?: Array<{ value?: string | number; color?: string }>
}) {
  const items = payload ?? []
  const visibleItems = items.slice(0, 12)
  const hiddenCount = Math.max(0, items.length - visibleItems.length)

  return (
    <div className="compact-legend">
      {visibleItems.map((entry) => (
        <span key={String(entry.value)} className="legend-chip" title={String(entry.value)}>
          <span className="legend-chip-dot" style={{ backgroundColor: entry.color }} />
          {truncateLabel(String(entry.value))}
        </span>
      ))}
      {hiddenCount > 0 && <span className="legend-chip legend-chip-more">+{hiddenCount} more</span>}
    </div>
  )
}

interface WeeklyTooltipEntry {
  dataKey?: string | number
  color?: string
  payload?: WeeklyRevenueRow
}

function WeeklyRevenueTooltip({
  active,
  payload,
  projectColorMap,
}: {
  active?: boolean
  payload?: WeeklyTooltipEntry[]
  projectColorMap: Record<string, string>
}) {
  if (!active || !payload || payload.length === 0) return null
  const row = payload[0]?.payload
  if (!row) return null
  const payloadColorMap: Record<string, string> = {}
  payload.forEach((entry) => {
    const dataKey = typeof entry.dataKey === 'string' ? entry.dataKey : null
    if (!dataKey || !entry.color) {
      return
    }
    payloadColorMap[dataKey] = entry.color
  })

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
          row.details.map((detail) => {
            const projectColor = payloadColorMap[detail.projectLabel] ?? projectColorMap[detail.projectLabel] ?? '#bfdbfe'
            return (
              <div key={`${row.weekStartIso}-${detail.projectLabel}`} className="revenue-tooltip-row">
                <div className="revenue-tooltip-project" style={{ color: projectColor }}>
                  {detail.projectLabel}
                </div>
                <div className="revenue-tooltip-metric" style={{ color: projectColor }}>
                  {detail.plannedHours.toFixed(1)} h
                </div>
                <div className="revenue-tooltip-metric" style={{ color: projectColor }}>
                  {formatCurrency(detail.revenuePerHour)}/h
                </div>
                <div className="revenue-tooltip-metric" style={{ color: projectColor }}>
                  {formatCurrency(detail.revenueAmount)}
                </div>
              </div>
            )
          })
        )}
      </div>
    </div>
  )
}

interface MonthlyRevenueTooltipEntry {
  payload?: MonthlyRevenueRow
}

function MonthlyRevenueTooltip({ active, payload }: { active?: boolean; payload?: MonthlyRevenueTooltipEntry[] }) {
  if (!active || !payload || payload.length === 0) return null
  const row = payload[0]?.payload
  if (!row) return null

  return (
    <div className="revenue-tooltip">
      <div className="revenue-tooltip-title">{row.monthLabel}</div>
      <div className="revenue-tooltip-summary">
        <div>Total Project Hours: {row.totalPlannedHours.toFixed(1)}</div>
        <div>Total Revenue: {formatCurrency(row.totalRevenue)}</div>
      </div>
      <div className="revenue-tooltip-grid">
        {row.details.length === 0 ? (
          <div className="revenue-tooltip-empty">No planned hours for this month.</div>
        ) : (
          row.details.map((detail) => (
            <div key={`${row.monthKey}-${detail.projectLabel}`} className="revenue-tooltip-row">
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

interface WeeklyGrossProfitTooltipEntry {
  dataKey?: string | number
  color?: string
  payload?: WeeklyGrossProfitRow
}

function WeeklyGrossProfitTooltip({
  active,
  payload,
  projectColorMap,
}: {
  active?: boolean
  payload?: WeeklyGrossProfitTooltipEntry[]
  projectColorMap: Record<string, string>
}) {
  if (!active || !payload || payload.length === 0) return null
  const row = payload[0]?.payload
  if (!row) return null

  const payloadColorMap: Record<string, string> = {}
  payload.forEach((entry) => {
    const dataKey = typeof entry.dataKey === 'string' ? entry.dataKey : null
    if (!dataKey || !entry.color) {
      return
    }
    payloadColorMap[dataKey] = entry.color
  })

  return (
    <div className="revenue-tooltip">
      <div className="revenue-tooltip-title">{row.weekRangeLabel}</div>
      <div className="revenue-tooltip-summary">
        <div>Total Project Hours: {row.totalPlannedHours.toFixed(1)}</div>
        <div>Total Gross Profit: {formatCurrency(row.totalGrossProfit)}</div>
      </div>
      <div className="revenue-tooltip-grid">
        {row.details.length === 0 ? (
          <div className="revenue-tooltip-empty">No planned hours for this week.</div>
        ) : (
          row.details.map((detail) => {
            const projectColor = payloadColorMap[detail.projectLabel] ?? projectColorMap[detail.projectLabel] ?? '#bfdbfe'
            return (
              <div key={`${row.weekStartIso}-${detail.projectLabel}`} className="revenue-tooltip-row">
                <div className="revenue-tooltip-project" style={{ color: projectColor }}>
                  {detail.projectLabel}
                </div>
                <div className="revenue-tooltip-metric" style={{ color: projectColor }}>
                  {detail.plannedHours.toFixed(1)} h
                </div>
                <div className="revenue-tooltip-metric" style={{ color: projectColor }}>
                  {formatCurrency(detail.grossProfitPerHour)}/h
                </div>
                <div className="revenue-tooltip-metric" style={{ color: projectColor }}>
                  {formatCurrency(detail.grossProfitAmount)}
                </div>
              </div>
            )
          })
        )}
      </div>
    </div>
  )
}

interface MonthlyGrossProfitTooltipEntry {
  payload?: MonthlyGrossProfitRow
}

function MonthlyGrossProfitTooltip({
  active,
  payload,
}: {
  active?: boolean
  payload?: MonthlyGrossProfitTooltipEntry[]
}) {
  if (!active || !payload || payload.length === 0) return null
  const row = payload[0]?.payload
  if (!row) return null

  return (
    <div className="revenue-tooltip">
      <div className="revenue-tooltip-title">{row.monthLabel}</div>
      <div className="revenue-tooltip-summary">
        <div>Total Project Hours: {row.totalPlannedHours.toFixed(1)}</div>
        <div>Total Gross Profit: {formatCurrency(row.totalGrossProfit)}</div>
      </div>
      <div className="revenue-tooltip-grid">
        {row.details.length === 0 ? (
          <div className="revenue-tooltip-empty">No planned hours for this month.</div>
        ) : (
          row.details.map((detail) => (
            <div key={`${row.monthKey}-${detail.projectLabel}`} className="revenue-tooltip-row">
              <div className="revenue-tooltip-project">{detail.projectLabel}</div>
              <div className="revenue-tooltip-metric">{detail.plannedHours.toFixed(1)} h</div>
              <div className="revenue-tooltip-metric">{formatCurrency(detail.grossProfitPerHour)}/h</div>
              <div className="revenue-tooltip-metric">{formatCurrency(detail.grossProfitAmount)}</div>
            </div>
          ))
        )}
      </div>
    </div>
  )
}

export function RevenueWorkspace({
  rateRows,
  weeklyRevenueRows,
  weeklyProjectKeys,
  weeklyGrossProfitRows,
  weeklyGrossProfitProjectKeys,
  monthlyRevenueRows,
  monthlyGrossProfitRows,
  monthlyProjectKeys,
  maxRatePerHour,
  onRateChange,
  mainSaveLabel,
  salesSaveLabel,
  mainSaveStatus,
  salesSaveStatus,
}: RevenueWorkspaceProps) {
  const [activeTab, setActiveTab] = useState<RevenueViewTab>('data')
  const [exportError, setExportError] = useState('')
  const [weeklyRevenueTooltipPosition, setWeeklyRevenueTooltipPosition] = useState<TooltipPosition | undefined>(undefined)
  const [weeklyGrossProfitTooltipPosition, setWeeklyGrossProfitTooltipPosition] = useState<TooltipPosition | undefined>(undefined)
  const [monthlyRevenueTooltipPosition, setMonthlyRevenueTooltipPosition] = useState<TooltipPosition | undefined>(undefined)
  const [monthlyGrossProfitTooltipPosition, setMonthlyGrossProfitTooltipPosition] = useState<TooltipPosition | undefined>(undefined)

  const weeklyChartData = useMemo(
    () =>
      weeklyRevenueRows.map((row) => ({
        ...row,
        ...row.amountsByProject,
      })),
    [weeklyRevenueRows],
  )

  const monthlyRevenueChartData = useMemo(
    () =>
      monthlyRevenueRows.map((row) => ({
        ...row,
        ...row.amountsByProject,
      })),
    [monthlyRevenueRows],
  )

  const monthlyGrossProfitChartData = useMemo(
    () =>
      monthlyGrossProfitRows.map((row) => ({
        ...row,
        ...row.amountsByProject,
      })),
    [monthlyGrossProfitRows],
  )
  const weeklyGrossProfitChartData = useMemo(
    () =>
      weeklyGrossProfitRows.map((row) => ({
        ...row,
        ...row.amountsByProject,
      })),
    [weeklyGrossProfitRows],
  )
  const weeklyProjectColorMap = useMemo(() => {
    const map: Record<string, string> = {}
    weeklyProjectKeys.forEach((projectKey, index) => {
      map[projectKey] = COLOR_PALETTE[index % COLOR_PALETTE.length]
    })
    return map
  }, [weeklyProjectKeys])
  const weeklyGrossProfitProjectColorMap = useMemo(() => {
    const map: Record<string, string> = {}
    weeklyGrossProfitProjectKeys.forEach((projectKey, index) => {
      map[projectKey] = COLOR_PALETTE[index % COLOR_PALETTE.length]
    })
    return map
  }, [weeklyGrossProfitProjectKeys])

  const exportRevenueWorkbook = (): void => {
    setExportError('')
    try {
      exportRevenueMonthlyWorkbook({
        monthlyRevenueRows,
        monthlyGrossProfitRows,
        fileName: buildRevenueExportFileName(),
      })
    } catch (saveError) {
      console.error('[capacity-dashboard] revenue workbook export failed', saveError)
      setExportError('Failed to export revenue workbook. Please try again.')
    }
  }

  return (
    <section className="panel revenue-page">
      <div className="section-header section-header-row revenue-header-row">
        <div>
          <h2>Revenue</h2>
          <p>Set project financial rates and monitor weekly revenue and gross profit from current planning hours.</p>
        </div>
        <div className="revenue-header-actions">
          <button type="button" onClick={exportRevenueWorkbook}>
            Export Monthly Forecast Excel
          </button>
          {exportError && <div className="export-error revenue-export-error">{exportError}</div>}
        </div>
      </div>

      <div className="report-tabs" aria-label="Revenue Views">
        <button
          type="button"
          className={activeTab === 'data' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveTab('data')}
        >
          Data
        </button>
        <button
          type="button"
          className={activeTab === 'revenue-chart' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveTab('revenue-chart')}
        >
          Revenue Chart
        </button>
        <button
          type="button"
          className={activeTab === 'gross-profit-chart' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveTab('gross-profit-chart')}
        >
          Gross Profit Chart
        </button>
        <button
          type="button"
          className={activeTab === 'revenue-monthly' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveTab('revenue-monthly')}
        >
          Revenue Monthly Forecast
        </button>
        <button
          type="button"
          className={activeTab === 'gross-profit-monthly' ? 'report-tab-btn report-tab-btn-active' : 'report-tab-btn'}
          onClick={() => setActiveTab('gross-profit-monthly')}
        >
          Gross Profit Monthly Forecast
        </button>
      </div>

      {activeTab === 'data' && (
        <div className="report-tab-panel revenue-data-panel">
          <div className="revenue-data-shell">
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

            <div className="panel table-panel revenue-rate-editor revenue-rate-panel">
              <div className="section-header">
                <h3>Rates by Project</h3>
                <p>Changes save to shared storage and are visible to all users.</p>
              </div>
              {rateRows.length === 0 ? (
                <div className="status">No projects are available in the current filter scope.</div>
              ) : (
                <div className="table-wrap revenue-rate-table-wrap">
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
                              onChange={(event) =>
                                onRateChange(row.dataset, row.project, 'revenuePerHour', event.target.value)
                              }
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
          </div>
        </div>
      )}

      {activeTab === 'revenue-chart' && (
        <div className="report-tab-panel">
          <div className="panel chart-panel">
            <div className="section-header">
              <h3>Weekly Revenue</h3>
              <p>Calculated as planned hours multiplied by project revenue-per-hour rates.</p>
            </div>
            {weeklyChartData.length === 0 || weeklyProjectKeys.length === 0 ? (
              <div className="status">No weekly revenue data available for the current scope.</div>
            ) : (
              <div className="chart-wrap revenue-chart-wrap">
                <ResponsiveContainer width="100%" height={560}>
                  <ComposedChart
                    data={weeklyChartData}
                    margin={{ top: 20, right: 20, left: 22, bottom: 36 }}
                    barCategoryGap="2%"
                    barGap={0}
                    barSize={resolveBarSize(weeklyChartData.length)}
                    onMouseMove={(state) => {
                      setWeeklyRevenueTooltipPosition(computeLeftTooltipPosition(state, REVENUE_TOOLTIP_BOUNDS))
                    }}
                    onMouseLeave={() => {
                      setWeeklyRevenueTooltipPosition(undefined)
                    }}
                  >
                    <CartesianGrid vertical={false} stroke="#334155" />
                    <XAxis
                      dataKey="weekLabel"
                      angle={-34}
                      textAnchor="end"
                      interval={0}
                      minTickGap={0}
                      height={72}
                      tickMargin={8}
                      tick={{ fontSize: 12, fill: '#e5e7eb', fontWeight: 600 }}
                      axisLine={{ stroke: '#475569' }}
                      tickLine={false}
                    />
                    <YAxis
                      tick={{ fontSize: 13, fill: '#e5e7eb', fontWeight: 600 }}
                      tickFormatter={(value: number) => formatCurrency(value)}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip
                      formatter={(value) => formatHours(value)}
                      content={<WeeklyRevenueTooltip projectColorMap={weeklyProjectColorMap} />}
                      position={weeklyRevenueTooltipPosition}
                    />
                    <Legend verticalAlign="top" align="left" content={<CompactLegend />} />
                    {weeklyProjectKeys.map((projectKey, index) => (
                      <Bar
                        key={projectKey}
                        dataKey={projectKey}
                        name={projectKey}
                        stackId="weekly-revenue"
                        fill={weeklyProjectColorMap[projectKey] ?? COLOR_PALETTE[index % COLOR_PALETTE.length]}
                      />
                    ))}
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === 'gross-profit-chart' && (
        <div className="report-tab-panel">
          <div className="panel chart-panel">
            <div className="section-header">
              <h3>Gross Profit by Project</h3>
              <p>Calculated as planned hours multiplied by project gross-profit-per-hour rates.</p>
            </div>
            {weeklyGrossProfitChartData.length === 0 || weeklyGrossProfitProjectKeys.length === 0 ? (
              <div className="status">No gross profit data available for the current scope.</div>
            ) : (
              <div className="chart-wrap revenue-chart-wrap">
                <ResponsiveContainer width="100%" height={560}>
                  <ComposedChart
                    data={weeklyGrossProfitChartData}
                    margin={{ top: 20, right: 20, left: 22, bottom: 36 }}
                    barCategoryGap="2%"
                    barGap={0}
                    barSize={resolveBarSize(weeklyGrossProfitChartData.length)}
                    onMouseMove={(state) => {
                      setWeeklyGrossProfitTooltipPosition(computeLeftTooltipPosition(state, REVENUE_TOOLTIP_BOUNDS))
                    }}
                    onMouseLeave={() => {
                      setWeeklyGrossProfitTooltipPosition(undefined)
                    }}
                  >
                    <CartesianGrid vertical={false} stroke="#334155" />
                    <XAxis
                      dataKey="weekLabel"
                      angle={-34}
                      textAnchor="end"
                      interval={0}
                      minTickGap={0}
                      height={72}
                      tickMargin={8}
                      tick={{ fontSize: 12, fill: '#e5e7eb', fontWeight: 600 }}
                      axisLine={{ stroke: '#475569' }}
                      tickLine={false}
                    />
                    <YAxis
                      tick={{ fontSize: 13, fill: '#e5e7eb', fontWeight: 600 }}
                      tickFormatter={(value: number) => formatCurrency(value)}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip
                      formatter={(value) => formatHours(value)}
                      content={<WeeklyGrossProfitTooltip projectColorMap={weeklyGrossProfitProjectColorMap} />}
                      position={weeklyGrossProfitTooltipPosition}
                    />
                    <Legend verticalAlign="top" align="left" content={<CompactLegend />} />
                    {weeklyGrossProfitProjectKeys.map((projectKey, index) => (
                      <Bar
                        key={projectKey}
                        dataKey={projectKey}
                        name={projectKey}
                        stackId="weekly-gross-profit"
                        fill={
                          weeklyGrossProfitProjectColorMap[projectKey] ?? COLOR_PALETTE[index % COLOR_PALETTE.length]
                        }
                      />
                    ))}
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === 'revenue-monthly' && (
        <div className="report-tab-panel">
          <div className="panel chart-panel">
            <div className="section-header">
              <h3>Revenue Monthly Forecast</h3>
              <p>Monthly revenue forecast aggregated from weekly planned hours and project rates.</p>
            </div>
            {monthlyRevenueChartData.length === 0 || monthlyProjectKeys.length === 0 ? (
              <div className="status">No monthly revenue data available for the current scope.</div>
            ) : (
              <div className="chart-wrap revenue-chart-wrap">
                <ResponsiveContainer width="100%" height={560}>
                  <ComposedChart
                    data={monthlyRevenueChartData}
                    margin={{ top: 20, right: 20, left: 22, bottom: 36 }}
                    barCategoryGap="4%"
                    barGap={0}
                    barSize={resolveBarSize(monthlyRevenueChartData.length)}
                    onMouseMove={(state) => {
                      setMonthlyRevenueTooltipPosition(computeLeftTooltipPosition(state, REVENUE_TOOLTIP_BOUNDS))
                    }}
                    onMouseLeave={() => {
                      setMonthlyRevenueTooltipPosition(undefined)
                    }}
                  >
                    <CartesianGrid vertical={false} stroke="#334155" />
                    <XAxis
                      dataKey="monthLabel"
                      angle={-34}
                      textAnchor="end"
                      interval={0}
                      minTickGap={0}
                      height={72}
                      tickMargin={8}
                      tick={{ fontSize: 12, fill: '#e5e7eb', fontWeight: 600 }}
                      axisLine={{ stroke: '#475569' }}
                      tickLine={false}
                    />
                    <YAxis
                      tick={{ fontSize: 13, fill: '#e5e7eb', fontWeight: 600 }}
                      tickFormatter={(value: number) => formatCurrency(value)}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip
                      formatter={(value) => formatHours(value)}
                      content={<MonthlyRevenueTooltip />}
                      position={monthlyRevenueTooltipPosition}
                    />
                    <Legend verticalAlign="top" align="left" content={<CompactLegend />} />
                    {monthlyProjectKeys.map((projectKey, index) => (
                      <Bar
                        key={projectKey}
                        dataKey={projectKey}
                        name={projectKey}
                        stackId="monthly-revenue"
                        fill={COLOR_PALETTE[index % COLOR_PALETTE.length]}
                      />
                    ))}
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        </div>
      )}

      {activeTab === 'gross-profit-monthly' && (
        <div className="report-tab-panel">
          <div className="panel chart-panel">
            <div className="section-header">
              <h3>Gross Profit Monthly Forecast</h3>
              <p>Monthly gross profit forecast aggregated from weekly planned hours and project rates.</p>
            </div>
            {monthlyGrossProfitChartData.length === 0 || monthlyProjectKeys.length === 0 ? (
              <div className="status">No monthly gross profit data available for the current scope.</div>
            ) : (
              <div className="chart-wrap revenue-chart-wrap">
                <ResponsiveContainer width="100%" height={560}>
                  <ComposedChart
                    data={monthlyGrossProfitChartData}
                    margin={{ top: 20, right: 20, left: 22, bottom: 36 }}
                    barCategoryGap="4%"
                    barGap={0}
                    barSize={resolveBarSize(monthlyGrossProfitChartData.length)}
                    onMouseMove={(state) => {
                      setMonthlyGrossProfitTooltipPosition(computeLeftTooltipPosition(state, REVENUE_TOOLTIP_BOUNDS))
                    }}
                    onMouseLeave={() => {
                      setMonthlyGrossProfitTooltipPosition(undefined)
                    }}
                  >
                    <CartesianGrid vertical={false} stroke="#334155" />
                    <XAxis
                      dataKey="monthLabel"
                      angle={-34}
                      textAnchor="end"
                      interval={0}
                      minTickGap={0}
                      height={72}
                      tickMargin={8}
                      tick={{ fontSize: 12, fill: '#e5e7eb', fontWeight: 600 }}
                      axisLine={{ stroke: '#475569' }}
                      tickLine={false}
                    />
                    <YAxis
                      tick={{ fontSize: 13, fill: '#e5e7eb', fontWeight: 600 }}
                      tickFormatter={(value: number) => formatCurrency(value)}
                      axisLine={false}
                      tickLine={false}
                    />
                    <Tooltip
                      formatter={(value) => formatHours(value)}
                      content={<MonthlyGrossProfitTooltip />}
                      position={monthlyGrossProfitTooltipPosition}
                    />
                    <Legend verticalAlign="top" align="left" content={<CompactLegend />} />
                    {monthlyProjectKeys.map((projectKey, index) => (
                      <Bar
                        key={projectKey}
                        dataKey={projectKey}
                        name={projectKey}
                        stackId="monthly-gross-profit"
                        fill={COLOR_PALETTE[index % COLOR_PALETTE.length]}
                      />
                    ))}
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        </div>
      )}
    </section>
  )
}
