import type { SummaryMetric } from './reportExport'
import type { MonthlyBucket, WeeklyBucket } from '../types'
import { buildExportApiUrl, withAuth } from './apiBase'

interface ExportWithChartApiInput {
  weeklyBuckets: WeeklyBucket[]
  monthlyBuckets: MonthlyBucket[]
  summary: SummaryMetric[]
  chartCategoryKeys: string[]
  fileName: string
}

function triggerDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  URL.revokeObjectURL(url)
}

export async function exportReportWorkbookWithChartApi(input: ExportWithChartApiInput): Promise<boolean> {
  const apiUrl = buildExportApiUrl()
  const weeklyCapacityChartRows = input.weeklyBuckets.map((bucket) => [
    bucket.weekStartIso,
    bucket.weekEndIso,
    bucket.weekLabel,
    bucket.totalHours,
    bucket.capacity,
    bucket.variance,
    bucket.status,
    ...input.chartCategoryKeys.map((key) => bucket.groups[key] ?? 0),
  ])
  const weeklyForecastRows = input.weeklyBuckets.map((bucket) => [
    bucket.weekStartIso,
    bucket.weekEndIso,
    bucket.weekLabel,
    bucket.totalHours,
    bucket.capacity,
    bucket.variance,
    bucket.status,
  ])
  const monthlyForecastRows = input.monthlyBuckets.map((bucket) => [
    bucket.monthLabel,
    bucket.plannedHours,
    bucket.capacity,
    bucket.variance,
    bucket.status,
  ])
  const summaryRows = input.summary.map((item) => [item.metric, item.value])

  try {
    const response = await fetch(apiUrl, {
      ...withAuth({
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileName: input.fileName,
          chartCategoryKeys: input.chartCategoryKeys,
          weeklyCapacityChartRows,
          weeklyForecastRows,
          monthlyForecastRows,
          summaryRows,
        }),
      }),
    })

    if (!response.ok) {
      return false
    }

    const blob = await response.blob()
    triggerDownload(blob, input.fileName)
    return true
  } catch {
    return false
  }
}
