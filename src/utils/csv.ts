import type { WeeklyBucket } from '../types'

function escapeCsvCell(value: string | number): string {
  const text = String(value)
  if (text.includes(',') || text.includes('"') || text.includes('\n')) {
    return `"${text.replace(/"/g, '""')}"`
  }
  return text
}

export function weeklyBucketsToCsv(rows: WeeklyBucket[]): string {
  const header = ['Week', 'Week Start', 'Week End', 'Total Hours', 'Capacity', 'Variance', 'Over Capacity']
  const lines = rows.map((row) => [
    row.weekLabel,
    row.weekStartIso,
    row.weekEndIso,
    row.totalHours.toFixed(2),
    row.capacity.toFixed(2),
    row.variance.toFixed(2),
    row.overCapacity ? 'Yes' : 'No',
  ])

  return [header, ...lines]
    .map((line) => line.map(escapeCsvCell).join(','))
    .join('\n')
}

export function downloadCsv(fileName: string, csvContent: string): void {
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  URL.revokeObjectURL(url)
}
