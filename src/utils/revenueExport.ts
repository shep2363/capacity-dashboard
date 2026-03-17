import * as XLSX from 'xlsx'
import type { MonthlyGrossProfitRow, MonthlyRevenueRow } from './revenue'

type CellValue = string | number

interface ExportRevenueMonthlyWorkbookInput {
  monthlyRevenueRows: MonthlyRevenueRow[]
  monthlyGrossProfitRows: MonthlyGrossProfitRow[]
  fileName: string
}

function safeNumber(value: number): number {
  return Number.isFinite(value) ? value : 0
}

function autoWidth(rows: CellValue[][]): Array<{ wch: number }> {
  if (rows.length === 0) {
    return []
  }

  const widths: number[] = []
  rows.forEach((row) => {
    row.forEach((cell, index) => {
      const cellLength = String(cell ?? '').length
      widths[index] = Math.max(widths[index] ?? 0, Math.min(60, cellLength + 2))
    })
  })

  return widths.map((wch) => ({ wch }))
}

function setHeaderRowFeatures(sheet: XLSX.WorkSheet): void {
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
  sheet['!freeze'] = { xSplit: 0, ySplit: 1 }
  sheet['!autofilter'] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: range.s.c },
      e: { r: 0, c: range.e.c },
    }),
  }
}

function buildMonthlyRevenueRows(rows: MonthlyRevenueRow[]): CellValue[][] {
  const sheetRows: CellValue[][] = [
    [
      'Month',
      'Project',
      'Planned Hours',
      'Applied Rate (/h)',
      'Computed Amount',
      'Monthly Planned Hours Total',
      'Monthly Revenue Total',
    ],
  ]

  rows.forEach((month) => {
    if (month.details.length === 0) {
      sheetRows.push([
        month.monthLabel,
        '(No planned hours)',
        0,
        0,
        0,
        safeNumber(month.totalPlannedHours),
        safeNumber(month.totalRevenue),
      ])
      return
    }

    month.details.forEach((detail, detailIndex) => {
      sheetRows.push([
        month.monthLabel,
        detail.projectLabel,
        safeNumber(detail.plannedHours),
        safeNumber(detail.revenuePerHour),
        safeNumber(detail.revenueAmount),
        detailIndex === 0 ? safeNumber(month.totalPlannedHours) : '',
        detailIndex === 0 ? safeNumber(month.totalRevenue) : '',
      ])
    })

    sheetRows.push([
      month.monthLabel,
      'Monthly Total',
      safeNumber(month.totalPlannedHours),
      '',
      safeNumber(month.totalRevenue),
      safeNumber(month.totalPlannedHours),
      safeNumber(month.totalRevenue),
    ])
  })

  return sheetRows
}

function buildMonthlyGrossProfitRows(rows: MonthlyGrossProfitRow[]): CellValue[][] {
  const sheetRows: CellValue[][] = [
    [
      'Month',
      'Project',
      'Planned Hours',
      'Applied Rate (/h)',
      'Computed Amount',
      'Monthly Planned Hours Total',
      'Monthly Gross Profit Total',
    ],
  ]

  rows.forEach((month) => {
    if (month.details.length === 0) {
      sheetRows.push([
        month.monthLabel,
        '(No planned hours)',
        0,
        0,
        0,
        safeNumber(month.totalPlannedHours),
        safeNumber(month.totalGrossProfit),
      ])
      return
    }

    month.details.forEach((detail, detailIndex) => {
      sheetRows.push([
        month.monthLabel,
        detail.projectLabel,
        safeNumber(detail.plannedHours),
        safeNumber(detail.grossProfitPerHour),
        safeNumber(detail.grossProfitAmount),
        detailIndex === 0 ? safeNumber(month.totalPlannedHours) : '',
        detailIndex === 0 ? safeNumber(month.totalGrossProfit) : '',
      ])
    })

    sheetRows.push([
      month.monthLabel,
      'Monthly Total',
      safeNumber(month.totalPlannedHours),
      '',
      safeNumber(month.totalGrossProfit),
      safeNumber(month.totalPlannedHours),
      safeNumber(month.totalGrossProfit),
    ])
  })

  return sheetRows
}

export function exportRevenueMonthlyWorkbook(input: ExportRevenueMonthlyWorkbookInput): void {
  const workbook = XLSX.utils.book_new()

  const revenueRows = buildMonthlyRevenueRows(input.monthlyRevenueRows)
  const revenueSheet = XLSX.utils.aoa_to_sheet(revenueRows)
  revenueSheet['!cols'] = autoWidth(revenueRows)
  setHeaderRowFeatures(revenueSheet)
  XLSX.utils.book_append_sheet(workbook, revenueSheet, 'Monthly Revenue Forecast')

  const grossProfitRows = buildMonthlyGrossProfitRows(input.monthlyGrossProfitRows)
  const grossProfitSheet = XLSX.utils.aoa_to_sheet(grossProfitRows)
  grossProfitSheet['!cols'] = autoWidth(grossProfitRows)
  setHeaderRowFeatures(grossProfitSheet)
  XLSX.utils.book_append_sheet(workbook, grossProfitSheet, 'Monthly Gross Profit Forecast')

  XLSX.writeFile(workbook, input.fileName)
}
