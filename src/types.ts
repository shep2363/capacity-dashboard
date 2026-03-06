export type ChartGroupBy = 'project' | 'resource'
export type PivotRowGrouping = 'project' | 'resource'

export interface TaskRow {
  id: string
  name: string
  workHours: number
  start: Date
  finish: Date
  resourceName: string
  project: string
}

export interface AppFilters {
  resource: string
  dateFrom: string
  dateTo: string
  year: string
}

export interface WeeklyBucket {
  weekStartIso: string
  weekEndIso: string
  weekLabel: string
  totalHours: number
  capacity: number
  variance: number
  overCapacity: boolean
  groups: Record<string, number>
}

export interface LeafCell {
  project: string
  resource: string
  weekStartIso: string
  hours: number
}

export interface PivotRow {
  rowKey: string
  rowLabel: string
  valuesByWeek: Record<string, number>
  totalHours: number
}

export interface PivotTableModel {
  rows: PivotRow[]
  weekKeys: string[]
  columnTotals: Record<string, number>
  grandTotal: number
  editedRowWeekKeys: Set<string>
}
