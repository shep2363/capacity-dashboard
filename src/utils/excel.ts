import * as XLSX from 'xlsx'
import { isValid, parse } from 'date-fns'
import type { TaskRow } from '../types'

const CANDIDATE_COLUMNS = {
  name: ['name', 'task name'],
  work: ['work', 'hours', 'estimated work'],
  start: ['start', 'start date'],
  finish: ['finish', 'end', 'finish date', 'due date'],
  resource: ['resource names', 'resource', 'resources', 'assignee'],
  project: ['project', 'proje', 'project name'],
}

function normalizeKey(input: string): string {
  return input.toLowerCase().replace(/[^a-z0-9]/g, '')
}

function findCellValue(row: Record<string, unknown>, candidates: string[]): unknown {
  const entries = Object.entries(row)
  const normalizedEntries = entries.map(([key, value]) => ({
    key,
    normalizedKey: normalizeKey(key),
    value,
  }))

  for (const candidate of candidates) {
    const target = normalizeKey(candidate)
    const exact = normalizedEntries.find((entry) => entry.normalizedKey === target)
    if (exact) {
      return exact.value
    }
  }

  for (const candidate of candidates) {
    const target = normalizeKey(candidate)
    const startsWithMatch = normalizedEntries.find((entry) => entry.normalizedKey.startsWith(target))
    if (startsWithMatch) {
      return startsWithMatch.value
    }
  }

  return undefined
}

function parseWorkHours(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }

  if (typeof value === 'string') {
    const cleaned = value.replace(/[^\d.-]/g, '')
    const parsed = Number.parseFloat(cleaned)
    return Number.isFinite(parsed) ? parsed : null
  }

  return null
}

function parseExcelDate(value: unknown): Date | null {
  if (value instanceof Date && isValid(value)) {
    return value
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = XLSX.SSF.parse_date_code(value)
    if (!date) {
      return null
    }
    return new Date(date.y, date.m - 1, date.d)
  }

  if (typeof value === 'string' && value.trim()) {
    const maybeDate = new Date(value)
    if (isValid(maybeDate)) {
      return maybeDate
    }

    const parsed = parse(value, 'M/d/yy', new Date())
    if (isValid(parsed)) {
      return parsed
    }
  }

  return null
}

export function parseSpreadsheet(arrayBuffer: ArrayBuffer): TaskRow[] {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true })
  const firstSheetName = workbook.SheetNames[0]

  if (!firstSheetName) {
    return []
  }

  const worksheet = workbook.Sheets[firstSheetName]
  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    defval: null,
    // Use displayed cell values so duration-formatted Work cells stay in hours.
    raw: false,
  })

  return rawRows
    .map((row, index) => {
      const nameValue = findCellValue(row, CANDIDATE_COLUMNS.name)
      const workValue = findCellValue(row, CANDIDATE_COLUMNS.work)
      const startValue = findCellValue(row, CANDIDATE_COLUMNS.start)
      const finishValue = findCellValue(row, CANDIDATE_COLUMNS.finish)
      const resourceValue = findCellValue(row, CANDIDATE_COLUMNS.resource)
      const projectValue = findCellValue(row, CANDIDATE_COLUMNS.project)

      const name = typeof nameValue === 'string' && nameValue.trim() ? nameValue.trim() : `Task ${index + 1}`
      const workHours = parseWorkHours(workValue)
      const start = parseExcelDate(startValue)
      const finish = parseExcelDate(finishValue)

      if (!workHours || workHours <= 0 || !start || !finish) {
        return null
      }

      return {
        id: `${index}-${name}`,
        name,
        workHours,
        start,
        finish,
        resourceName:
          typeof resourceValue === 'string' && resourceValue.trim() ? resourceValue.trim() : 'Unassigned',
        project: typeof projectValue === 'string' && projectValue.trim() ? projectValue.trim() : 'Unspecified',
      } satisfies TaskRow
    })
    .filter((task): task is TaskRow => task !== null)
}
