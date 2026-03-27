import { useState } from 'react'
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import {
  chartTooltipGlassStyle,
  resolveTooltipPosition,
  type TooltipPosition,
  type RechartsTooltipSnapshot,
} from '../utils/chartTooltip'

export interface KpiSet {
  bookedYtd: number
  capacityYtd: number
  utilization: number
  remaining: number
  activeProjects: number
  status: 'green' | 'yellow' | 'red'
  label: string
}

export interface ExecutiveData {
  combinedKpis: KpiSet
  opsKpis: KpiSet
  salesKpis: KpiSet
  monthlyComparison: Array<{ month: string; opsBooked: number; salesBooked: number; totalBooked: number; capacity: number }>
  quarterlySummary: Array<{ quarter: string; booked: number; capacity: number; utilization: number }>
  annual: { booked: number; capacity: number; utilization: number; status: 'Under Capacity' | 'Within Capacity' | 'Over Capacity' }
  riskMonths: Array<{ monthLabel: string; variance: number }>
  topProjects: Array<{ project: string; hours: number; percent: number }>
  utilizationTrend: Array<{ monthLabel: string; utilization: number }>
}

interface ExecutiveSummaryProps {
  data: ExecutiveData
}

function formatHours(value: number): string {
  if (Math.abs(value) >= 1000) {
    return `${(value / 1000).toFixed(1)}k`
  }
  return value.toFixed(0)
}

function utilizationStatus(utilization: number): 'green' | 'red' {
  return utilization >= 0.95 ? 'green' : 'red'
}

const statusColors: Record<string, string> = {
  green: '#22c55e',
  red: '#ef4444',
}

export function ExecutiveSummary({ data }: ExecutiveSummaryProps) {
  const [monthlyTooltipPosition, setMonthlyTooltipPosition] = useState<TooltipPosition | undefined>(undefined)
  const [utilizationTooltipPosition, setUtilizationTooltipPosition] = useState<TooltipPosition | undefined>(undefined)
  const { combinedKpis, opsKpis, salesKpis, monthlyComparison, quarterlySummary, annual, riskMonths, topProjects, utilizationTrend } = data

  const monthlyChartData = monthlyComparison

  const utilizationChartData = utilizationTrend.map((row) => ({
    month: row.monthLabel,
    utilization: Number((row.utilization * 100).toFixed(1)),
  }))

  return (
    <div className="panel executive-panel">
      <div className="section-header">
        <h2>Executive Summary</h2>
        <p>High-level view of workload, capacity, and risk based on current filters.</p>
      </div>

      <div
        className="kpi-grid"
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
          gap: '12px',
        }}
      >
        {[
          { label: 'Booked Hours (YTD)', value: combinedKpis.bookedYtd, suffix: '', color: '#38bdf8' },
          { label: 'Available Capacity (YTD)', value: combinedKpis.capacityYtd, suffix: '', color: '#a855f7' },
          {
            label: 'Utilization',
            value: combinedKpis.utilization * 100,
            suffix: '%',
            color: statusColors[utilizationStatus(combinedKpis.utilization)],
          },
          { label: 'Remaining Capacity', value: combinedKpis.remaining, suffix: '', color: '#facc15' },
          { label: 'Active Projects', value: combinedKpis.activeProjects, suffix: '', color: '#22c55e' },
        ].map((card) => (
          <div
            key={card.label}
            className="kpi-card"
            style={{
              border: '1px solid #1f2937',
              borderRadius: 10,
              padding: '12px 14px',
              background: '#0b1220',
              borderColor: card.color,
            }}
          >
            <span className="kpi-label" style={{ color: '#9ca3af', fontSize: '0.95rem' }}>
              {card.label}
            </span>
            <div className="kpi-value" style={{ color: card.color, fontSize: '1.6rem', fontWeight: 700 }}>
              {card.label === 'Utilization'
                ? `${card.value.toFixed(1)}${card.suffix}`
                : `${formatHours(card.value)}${card.suffix}`}
            </div>
          </div>
        ))}
      </div>

      <div
        className="two-column"
        style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: 12 }}
      >
        {[opsKpis, salesKpis].map((kpi) => (
          <div key={kpi.label} className="panel sub-panel" style={{ background: '#0b1220' }}>
            <div className="section-header">
              <h3>{kpi.label} (YTD)</h3>
            </div>
            <div className="annual-summary" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 12 }}>
              <div>
                <span>Booked</span>
                <strong>{formatHours(kpi.bookedYtd)}</strong>
              </div>
              <div>
                <span>Capacity</span>
                <strong>{formatHours(kpi.capacityYtd)}</strong>
              </div>
              <div>
                <span>Utilization</span>
                <strong style={{ color: statusColors[utilizationStatus(kpi.utilization)] }}>
                  {(kpi.utilization * 100).toFixed(1)}%
                </strong>
              </div>
              <div>
                <span>Remaining</span>
                <strong>{formatHours(kpi.remaining)}</strong>
              </div>
            </div>
          </div>
        ))}
      </div>

      <div className="two-column" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: 12 }}>
        <div className="panel sub-panel" style={{ background: '#0b1220' }}>
          <div className="section-header">
            <h3>Monthly Capacity vs Bookings</h3>
          </div>
          <div style={{ height: 260 }}>
            <ResponsiveContainer>
              <BarChart
                data={monthlyChartData}
                onMouseMove={(state) =>
                  setMonthlyTooltipPosition(
                    resolveTooltipPosition(state as RechartsTooltipSnapshot, {
                      width: 220,
                      height: 120,
                    }),
                  )
                }
                onMouseLeave={() => setMonthlyTooltipPosition(undefined)}
              >
                <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
                <XAxis dataKey="month" stroke="#9ca3af" />
                <YAxis stroke="#9ca3af" />
                <Tooltip
                  contentStyle={chartTooltipGlassStyle}
                  position={monthlyTooltipPosition}
                />
                <Legend />
                <Bar dataKey="opsBooked" name="Shop Booked" stackId="booked" fill="#38bdf8" />
                <Bar dataKey="salesBooked" name="Sales Booked" stackId="booked" fill="#f472b6" />
                <Line type="monotone" dataKey="capacity" name="Capacity" stroke="#a855f7" strokeWidth={3} dot={false} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="panel sub-panel" style={{ background: '#0b1220' }}>
          <div className="section-header">
            <h3>Utilization Trend</h3>
          </div>
          <div style={{ height: 260 }}>
            <ResponsiveContainer>
              <LineChart
                data={utilizationChartData}
                onMouseMove={(state) =>
                  setUtilizationTooltipPosition(
                    resolveTooltipPosition(state as RechartsTooltipSnapshot, {
                      width: 220,
                      height: 90,
                    }),
                  )
                }
                onMouseLeave={() => setUtilizationTooltipPosition(undefined)}
              >
                <CartesianGrid strokeDasharray="3 3" stroke="#1f2937" />
                <XAxis dataKey="month" stroke="#9ca3af" />
                <YAxis stroke="#9ca3af" unit="%" />
                <Tooltip
                  contentStyle={chartTooltipGlassStyle}
                  position={utilizationTooltipPosition}
                />
                <Legend />
                <Line
                  type="monotone"
                  dataKey="utilization"
                  name="Utilization %"
                  stroke="#22c55e"
                  strokeWidth={3}
                  dot={false}
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="two-column" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: 12 }}>
        <div className="panel sub-panel" style={{ background: '#0b1220' }}>
          <div className="section-header">
            <h3>Quarterly Performance</h3>
          </div>
          <table className="compact-table">
            <thead>
              <tr>
                <th>Quarter</th>
                <th>Booked</th>
                <th>Capacity</th>
                <th>Utilization</th>
              </tr>
            </thead>
            <tbody>
              {quarterlySummary.map((row) => {
                const status = utilizationStatus(row.utilization)
                return (
                  <tr key={row.quarter}>
                    <td>{row.quarter}</td>
                    <td>{formatHours(row.booked)}</td>
                    <td>{formatHours(row.capacity)}</td>
                    <td style={{ color: statusColors[status] }}>{(row.utilization * 100).toFixed(1)}%</td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>

        <div className="panel sub-panel" style={{ background: '#0b1220' }}>
          <div className="section-header">
            <h3>Annual Utilization</h3>
          </div>
          <div className="annual-summary" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: 12 }}>
            <div>
              <span>Total Booked</span>
              <strong>{formatHours(annual.booked)}</strong>
            </div>
            <div>
              <span>Total Capacity</span>
              <strong>{formatHours(annual.capacity)}</strong>
            </div>
            <div>
              <span>Utilization</span>
              <strong style={{ color: statusColors[utilizationStatus(annual.utilization)] }}>
                {(annual.utilization * 100).toFixed(1)}%
              </strong>
            </div>
            <div
              className="status-chip"
              data-status={annual.status}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                padding: '6px 10px',
                borderRadius: 999,
                background: '#1f2937',
              }}
            >
              {annual.status}
            </div>
          </div>
        </div>
      </div>

      <div className="two-column" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: 12 }}>
        <div className="panel sub-panel" style={{ background: '#0b1220' }}>
          <div className="section-header">
            <h3>Capacity Risk Months</h3>
            <p>Months where booked hours exceed capacity.</p>
          </div>
          {riskMonths.length === 0 ? (
            <div className="muted">No capacity risks in view.</div>
          ) : (
            <ul className="risk-list" style={{ listStyle: 'none', padding: 0, margin: 0 }}>
              {riskMonths.map((month) => (
                <li
                  key={month.monthLabel}
                  className="risk-item"
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    padding: '8px 10px',
                    border: '1px solid #1f2937',
                    borderRadius: 8,
                    marginBottom: 6,
                  }}
                >
                  <span>{month.monthLabel}</span>
                  <span
                    className="risk-badge"
                    style={{
                      background: '#f97316',
                      color: '#0b1220',
                      borderRadius: 999,
                      padding: '4px 8px',
                      fontWeight: 700,
                    }}
                  >
                    Capacity Risk
                  </span>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="panel sub-panel" style={{ background: '#0b1220' }}>
          <div className="section-header">
            <h3>Top Projects by Hours</h3>
          </div>
          <table className="compact-table">
            <thead>
              <tr>
                <th>Project</th>
                <th>Hours</th>
                <th>% Workload</th>
              </tr>
            </thead>
            <tbody>
              {topProjects.map((row) => (
                <tr key={row.project}>
                  <td>{row.project}</td>
                  <td>{formatHours(row.hours)}</td>
                  <td>{row.percent.toFixed(1)}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
