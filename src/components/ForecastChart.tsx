import { useMemo, useState, type MouseEvent as ReactMouseEvent } from 'react'
import {
  Bar,
  CartesianGrid,
  Cell,
  ComposedChart,
  Legend,
  Line,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import type { WeeklyBucket } from '../types'
import { shortWeekLabel } from '../utils/planner'

interface ForecastChartProps {
  weeklyBuckets: WeeklyBucket[]
  categoryKeys: string[]
  projects: string[]
  selectedProjects: Set<string>
  onToggleProject: (project: string) => void
  title?: string
  subtitle?: string
  hoveredProject?: string | null
  onHoverProject?: (project: string | null) => void
  hoverProjectPrefix?: string
  selectedWeekIds?: Set<string>
  onWeekSelect?: (weekStartIso: string, multiSelect: boolean) => void
}

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

function formatHours(value: unknown): string {
  const normalized = Array.isArray(value) ? value[0] : value
  const numeric = typeof normalized === 'number' ? normalized : Number(normalized ?? 0)
  return `${numeric.toFixed(1)} h`
}

function truncateLabel(label: string): string {
  return label.length > 26 ? `${label.slice(0, 26)}...` : label
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

interface TooltipEntry {
  name?: string | number
  dataKey?: string | number
  value?: number | string | Array<number | string>
  color?: string
  payload?: { weekRangeLabel?: string; holidayTooltip?: string }
}

interface CustomTooltipProps {
  active?: boolean
  payload?: TooltipEntry[]
  label?: string | number
}

function ForecastTooltipCard({ active, payload, label }: CustomTooltipProps) {
  if (!active || !payload || payload.length === 0) {
    return null
  }

  const weekRangeLabel = String(payload[0]?.payload?.weekRangeLabel ?? label ?? '')
  const holidayTooltip = String(payload[0]?.payload?.holidayTooltip ?? '')
  const totalProjectHours = payload
    .filter((entry) => entry.dataKey !== 'capacity')
    .reduce((sum: number, entry: TooltipEntry) => {
      const value = Array.isArray(entry.value) ? entry.value[0] : entry.value
      const numeric = typeof value === 'number' ? value : Number(value ?? 0)
      return sum + (Number.isFinite(numeric) ? numeric : 0)
    }, 0)

  const capacityEntry = payload.find((entry) => entry.dataKey === 'capacity')
  const capacityRaw = Array.isArray(capacityEntry?.value) ? capacityEntry?.value[0] : capacityEntry?.value
  const capacityHours =
    typeof capacityRaw === 'number' ? capacityRaw : Number.isFinite(Number(capacityRaw)) ? Number(capacityRaw) : 0
  const variance = totalProjectHours - capacityHours
  const overCapacityHours = Math.max(variance, 0)
  const underCapacityHours = Math.max(-variance, 0)
  const utilizationPercent = capacityHours > 0 ? (totalProjectHours / capacityHours) * 100 : null

  return (
    <div
      style={{
        borderRadius: 12,
        border: '1px solid #1f2937',
        fontSize: '0.95rem',
        backgroundColor: '#0f172a',
        color: '#e5e7eb',
        padding: '0.65rem 0.75rem',
        minWidth: 210,
      }}
    >
      <div style={{ fontWeight: 700, marginBottom: 6 }}>{weekRangeLabel}</div>
      <div
        style={{
          fontWeight: 800,
          borderTop: '1px solid #334155',
          paddingTop: 6,
          marginBottom: 6,
          color: '#bfdbfe',
        }}
      >
        Total Project Hours: {totalProjectHours.toFixed(1)} h
      </div>
      {holidayTooltip ? (
        <div style={{ marginBottom: 6, padding: '6px 8px', background: '#111827', borderRadius: 8 }}>
          <div style={{ fontWeight: 700, color: '#fbbf24', marginBottom: 4 }}>Holidays</div>
          {holidayTooltip.split('\n').map((line) => (
            <div key={line} style={{ color: '#e5e7eb', fontSize: '0.93rem', lineHeight: 1.35 }}>
              {line}
            </div>
          ))}
        </div>
      ) : null}
      <div style={{ display: 'grid', gap: 4, marginBottom: 6 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
          <span style={{ color: '#93c5fd' }}>Capacity</span>
          <span>{capacityHours.toFixed(1)} h</span>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
          <span style={{ color: '#86efac' }}>Utilization %</span>
          <span>{utilizationPercent === null ? 'N/A' : `${utilizationPercent.toFixed(1)}%`}</span>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
          <span style={{ color: '#f59e0b', fontWeight: 700 }}>Over Capacity Hours</span>
          <span style={{ color: '#f59e0b', fontWeight: 700 }}>{overCapacityHours.toFixed(1)} h</span>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}>
          <span style={{ color: '#ef4444', fontWeight: 700 }}>Under Capacity Hours</span>
          <span style={{ color: '#ef4444', fontWeight: 700 }}>{underCapacityHours.toFixed(1)} h</span>
        </div>
      </div>
      <div style={{ display: 'grid', gap: 4 }}>
        {payload
          .filter((entry) => entry.dataKey !== 'capacity')
          .map((entry: TooltipEntry) => {
          const value = Array.isArray(entry.value) ? entry.value[0] : entry.value
          const numeric = typeof value === 'number' ? value : Number(value ?? 0)
          const safeValue = Number.isFinite(numeric) ? numeric : 0
          return (
            <div
              key={`${String(entry.name)}-${String(entry.dataKey)}`}
              style={{ display: 'flex', justifyContent: 'space-between', gap: 12 }}
            >
              <span style={{ color: entry.color ?? '#e5e7eb' }}>{String(entry.name)}</span>
              <span>{safeValue.toFixed(1)} h</span>
            </div>
          )
        })}
      </div>
    </div>
  )
}

export function ForecastChart({
  weeklyBuckets,
  categoryKeys,
  projects,
  selectedProjects,
  onToggleProject,
  title,
  subtitle = 'Stacked weekly forecast hours with total capacity from selected resources.',
  hoveredProject = null,
  onHoverProject,
  hoverProjectPrefix = '',
  selectedWeekIds,
  onWeekSelect,
}: ForecastChartProps) {
  const Y_AXIS_STEP = 500
  const MIN_Y_AXIS_MAX = 1000
  const [hoverState, setHoverState] = useState<ForecastHoverState>({ active: false })

  function buildYAxisTicks(maxValue: number): number[] {
    const ticks: number[] = []
    for (let value = 0; value <= maxValue; value += Y_AXIS_STEP) {
      ticks.push(value)
    }
    return ticks
  }

  function buildAxisScale(maxValue: number): { max: number; ticks: number[] } {
    const roundedPeak = Math.ceil(maxValue / Y_AXIS_STEP) * Y_AXIS_STEP
    const boundedMax = Math.max(MIN_Y_AXIS_MAX, roundedPeak || 0)
    return {
      max: boundedMax,
      ticks: buildYAxisTicks(boundedMax),
    }
  }

  const categoryOrder = useMemo(() => {
    const totals = new Map<string, number>()

    for (const week of weeklyBuckets) {
      for (const [category, hours] of Object.entries(week.groups)) {
        totals.set(category, (totals.get(category) ?? 0) + hours)
      }
    }

    return [...categoryKeys].sort((a, b) => (totals.get(b) ?? 0) - (totals.get(a) ?? 0))
  }, [weeklyBuckets, categoryKeys])

  const chartData = weeklyBuckets.map((bucket) => ({
    id: bucket.weekStartIso,
    weekLabel: shortWeekLabel(bucket.weekStartIso),
    weekRangeLabel: bucket.weekLabel,
    totalHours: bucket.totalHours,
    capacity: bucket.capacity,
    variance: bucket.variance,
    overCapacity: bucket.overCapacity,
    holidayTooltip:
      bucket.holidayDetails && bucket.holidayDetails.length > 0
        ? bucket.holidayDetails.map((h) => `${h.name} - ${h.date}`).join('\n')
        : '',
    ...bucket.groups,
  }))

  function handleWeekSelection(weekStartIso: string, event: Pick<ReactMouseEvent<Element>, 'ctrlKey' | 'metaKey'>): void {
    if (!onWeekSelect || !weekStartIso) {
      return
    }
    onWeekSelect(weekStartIso, event.ctrlKey || event.metaKey)
  }

  function toHoverKey(project: string): string {
    return `${hoverProjectPrefix}${project}`
  }

  function isHighlighted(category: string): boolean {
    if (!hoveredProject) {
      return true
    }
    return hoveredProject === category || hoveredProject === toHoverKey(category)
  }

  const barSize = useMemo(() => {
    if (chartData.length > 80) {
      return 14
    }
    if (chartData.length > 60) {
      return 17
    }
    if (chartData.length > 40) {
      return 20
    }
    if (chartData.length > 24) {
      return 24
    }
    return 30
  }, [chartData.length])

  function renderWeekTick(props: {
    x?: number | string
    y?: number | string
    payload?: { value?: string | number; index?: number }
  }) {
    const tickIndex = props.payload?.index ?? 0
    const week = chartData[tickIndex]
    const weekId = week?.id ?? ''
    const isSelected = weekId ? selectedWeekIds?.has(weekId) ?? false : false
    const x = Number(props.x ?? 0)
    const y = Number(props.y ?? 0)

    return (
      <g
        transform={`translate(${x},${y})`}
        onClick={(event) => {
          if (!weekId) {
            return
          }
          handleWeekSelection(weekId, event)
        }}
        style={{ cursor: onWeekSelect ? 'pointer' : 'default' }}
      >
        <text
          x={0}
          y={0}
          dy={18}
          textAnchor="end"
          transform="rotate(-34)"
          fill={isSelected ? '#fbbf24' : '#e5e7eb'}
          fontSize={12}
          fontWeight={isSelected ? 800 : 600}
        >
          {String(props.payload?.value ?? '')}
        </text>
      </g>
    )
  }

  const { axisMax, axisTicks } = useMemo(() => {
    const peak = chartData.reduce(
      (max, row) => Math.max(max, Number(row.totalHours ?? 0), Number(row.capacity ?? 0)),
      0,
    )
    const scale = buildAxisScale(peak)

    return {
      axisMax: scale.max,
      axisTicks: scale.ticks,
    }
  }, [chartData])

  return (
    <div className="panel chart-panel">
      <div className="section-header">
        {title ? <h2>{title}</h2> : null}
        {subtitle ? <p>{subtitle}</p> : null}
        <div className="toggle-chips">
          {projects.map((project) => {
            const on = selectedProjects.has(project)
            return (
              <button
                key={project}
                type="button"
                className={`chip-toggle ${on ? 'chip-on' : 'chip-off'}`}
                onClick={() => onToggleProject(project)}
                onMouseEnter={() => onHoverProject?.(toHoverKey(project))}
                onMouseLeave={() => onHoverProject?.(null)}
                aria-pressed={on}
              >
                {project}
              </button>
            )
          })}
        </div>
      </div>

      <div className="chart-layout">
        <div className="chart-wrap">
          <ResponsiveContainer width="100%" height={560}>
            <ComposedChart
              data={chartData}
              margin={{ top: 20, right: 20, left: 22, bottom: 36 }}
              barCategoryGap="2%"
              barGap={0}
              barSize={barSize}
              onMouseMove={(state) => {
                const hover = state as RechartsHoverSnapshot<TooltipEntry>
                setHoverState({
                  active: Boolean(hover.isTooltipActive),
                  payload: hover.activePayload,
                  label: hover.activeLabel,
                })
              }}
              onMouseLeave={() => {
                setHoverState({ active: false })
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
                tick={renderWeekTick}
                axisLine={{ stroke: '#475569' }}
                tickLine={false}
              />
              <YAxis
                domain={[0, axisMax]}
                ticks={axisTicks}
                tick={{ fontSize: 13, fill: '#e5e7eb', fontWeight: 600 }}
                tickFormatter={(value: number) => `${value}`}
                axisLine={false}
                tickLine={false}
              />
              <Tooltip formatter={(value) => formatHours(value)} content={() => null} />
              <Legend verticalAlign="top" align="left" content={<CompactLegend />} />

              {categoryOrder.map((category, index) => (
                <Bar
                  key={category}
                  dataKey={category}
                  stackId="weekly"
                  fill={COLOR_PALETTE[index % COLOR_PALETTE.length]}
                  fillOpacity={isHighlighted(category) ? 1 : 0.25}
                  name={category}
                  onClick={(data, _index, event) => {
                    const weekId = data?.payload?.id
                    if (!weekId) {
                      return
                    }
                    handleWeekSelection(weekId, event)
                  }}
                >
                  {chartData.map((row) => {
                    const isSelected = selectedWeekIds?.has(row.id) ?? false
                    return (
                      <Cell
                        key={`${category}-${row.id}`}
                        cursor={onWeekSelect ? 'pointer' : 'default'}
                        stroke={isSelected ? '#fbbf24' : undefined}
                        strokeWidth={isSelected ? 2 : 0}
                      />
                    )
                  })}
                </Bar>
              ))}

              <Line
                type="monotone"
                dataKey="capacity"
                name="Capacity"
                stroke="#93c5fd"
                strokeWidth={3}
                dot={false}
                strokeDasharray="6 4"
                connectNulls
                strokeLinecap="round"
                isAnimationActive={false}
              />
            </ComposedChart>
          </ResponsiveContainer>
        </div>
        <aside className="chart-hover-dock" aria-live="polite">
          {hoverState.active && hoverState.payload && hoverState.payload.length > 0 ? (
            <ForecastTooltipCard active payload={hoverState.payload} label={hoverState.label} />
          ) : (
            <div className="chart-hover-placeholder">
              <strong>Hover details</strong>
              <span>Hover a week in the chart to view hours, capacity, utilization, and project totals here.</span>
            </div>
          )}
        </aside>
      </div>

      <p className="chart-note">
        Weeks over capacity are highlighted in the pivot and summary tables.
        {onWeekSelect ? ' Ctrl + click week bars or labels to total multiple weeks.' : ''}
      </p>
    </div>
  )
}

interface ForecastHoverState {
  active: boolean
  payload?: TooltipEntry[]
  label?: string | number
}

interface RechartsHoverSnapshot<TPayload> {
  isTooltipActive?: boolean
  activePayload?: TPayload[]
  activeLabel?: string | number
}
