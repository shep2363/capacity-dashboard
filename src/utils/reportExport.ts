import * as XLSX from 'xlsx'
import type { MonthlyBucket, WeeklyBucket } from '../types'

export interface SummaryMetric {
  metric: string
  value: string | number
}

interface ExportReportWorkbookInput {
  weeklyBuckets: WeeklyBucket[]
  monthlyBuckets: MonthlyBucket[]
  chartCategoryKeys: string[]
  summary: SummaryMetric[]
  fileName: string
}

function normalizeNumber(value: number): number {
  return Number.isFinite(value) ? value : 0
}

function autoWidth(rows: Array<Array<string | number>>): Array<{ wch: number }> {
  if (rows.length === 0) {
    return []
  }

  const widths: number[] = []
  for (const row of rows) {
    row.forEach((cell, index) => {
      const cellLength = String(cell ?? '').length
      widths[index] = Math.max(widths[index] ?? 0, Math.min(60, cellLength + 2))
    })
  }

  return widths.map((wch) => ({ wch }))
}

function setHeaderRowFeatures(sheet: XLSX.WorkSheet): void {
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
  // Freeze header row (supported by most spreadsheet viewers consuming this metadata).
  sheet['!freeze'] = { xSplit: 0, ySplit: 1 }
  // Filter on header row for quick stakeholder sorting/filtering in Excel.
  sheet['!autofilter'] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: range.s.c },
      e: { r: 0, c: range.e.c },
    }),
  }
}

export function exportReportWorkbook(input: ExportReportWorkbookInput): void {
  const workbook = XLSX.utils.book_new()

  const chartRows: Array<Array<string | number>> = [
    [
      'Week Start',
      'Week End',
      'Week Label',
      'Total Forecast Hours',
      'Capacity',
      'Variance',
      'Status',
      ...input.chartCategoryKeys,
    ],
    ...input.weeklyBuckets.map((bucket) => [
      bucket.weekStartIso,
      bucket.weekEndIso,
      bucket.weekLabel,
      normalizeNumber(bucket.totalHours),
      normalizeNumber(bucket.capacity),
      normalizeNumber(bucket.variance),
      bucket.status,
      ...input.chartCategoryKeys.map((key) => normalizeNumber(bucket.groups[key] ?? 0)),
    ]),
  ]
  const chartSheet = XLSX.utils.aoa_to_sheet(chartRows)
  chartSheet['!cols'] = autoWidth(chartRows)
  setHeaderRowFeatures(chartSheet)
  XLSX.utils.book_append_sheet(workbook, chartSheet, 'Weekly Capacity Chart')

  const weeklyRows: Array<Array<string | number>> = [
    ['Week Start', 'Week End', 'Week Label', 'Forecast Hours', 'Capacity', 'Variance', 'Status'],
    ...input.weeklyBuckets.map((bucket) => [
      bucket.weekStartIso,
      bucket.weekEndIso,
      bucket.weekLabel,
      normalizeNumber(bucket.totalHours),
      normalizeNumber(bucket.capacity),
      normalizeNumber(bucket.variance),
      bucket.status,
    ]),
  ]
  const weeklySheet = XLSX.utils.aoa_to_sheet(weeklyRows)
  weeklySheet['!cols'] = autoWidth(weeklyRows)
  setHeaderRowFeatures(weeklySheet)
  XLSX.utils.book_append_sheet(workbook, weeklySheet, 'Weekly Forecast')

  const monthlyRows: Array<Array<string | number>> = [
    ['Month', 'Planned Hours', 'Capacity', 'Variance', 'Status'],
    ...input.monthlyBuckets.map((bucket) => [
      bucket.monthLabel,
      normalizeNumber(bucket.plannedHours),
      normalizeNumber(bucket.capacity),
      normalizeNumber(bucket.variance),
      bucket.status,
    ]),
  ]
  const monthlySheet = XLSX.utils.aoa_to_sheet(monthlyRows)
  monthlySheet['!cols'] = autoWidth(monthlyRows)
  setHeaderRowFeatures(monthlySheet)
  XLSX.utils.book_append_sheet(workbook, monthlySheet, 'Monthly Forecast')

  const summaryRows: Array<Array<string | number>> = [
    ['Metric', 'Value'],
    ...input.summary.map((item) => [item.metric, typeof item.value === 'number' ? normalizeNumber(item.value) : item.value]),
  ]
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows)
  summarySheet['!cols'] = autoWidth(summaryRows)
  setHeaderRowFeatures(summarySheet)
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary')

  XLSX.writeFile(workbook, input.fileName)
}
