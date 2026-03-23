import * as XLSX from 'xlsx'
import { isValid, parse, parse as parseStrict, compareAsc } from 'date-fns'
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

function parseProbability(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (value >= 0 && value < 1) {
      return value * 100
    }
    return value
  }

  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) {
      return null
    }
    const cleaned = trimmed.replace(/[^\d.-]/g, '')
    const parsed = Number.parseFloat(cleaned)
    if (!Number.isFinite(parsed)) {
      return null
    }
    if (!trimmed.includes('%') && parsed >= 0 && parsed < 1) {
      return parsed * 100
    }
    return parsed
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
    .map<TaskRow | null>((row, index) => {
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

function parseDateSheetName(sheetName: string): Date | null {
  // Accept sheet names like "3-5-2026" or "03-05-2026".
  const match = sheetName.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/)
  if (!match) return null
  const [_, m, d, y] = match
  const parsed = parseStrict(`${m}-${d}-${y}`, 'M-d-yyyy', new Date())
  return isValid(parsed) ? parsed : null
}

export function parseSalesSpreadsheet(arrayBuffer: ArrayBuffer): TaskRow[] {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true })
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    return []
  }

  const datedSheets = workbook.SheetNames.map((name) => ({
    name,
    date: parseDateSheetName(name),
  }))
    .filter((entry) => entry.date)
    .sort((a, b) => compareAsc(a.date as Date, b.date as Date))

  const targetSheetName = datedSheets.length > 0 ? datedSheets[datedSheets.length - 1].name : workbook.SheetNames[0]
  const worksheet = workbook.Sheets[targetSheetName]

  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(worksheet, {
    defval: null,
    raw: false,
  })

  const salesColumnCandidates = {
    fab: ['fabhrs', 'fab hrs', 'fab hours', 'fab'],
    shipping: ['shipping', 'shipping hrs', 'shipping hours'],
    blast: ['blast', 'blast hrs', 'blast hours'],
    paint: ['paint', 'paint hrs', 'paint hours'],
    start: ['expectedshopstart', 'expected shop start', 'shop start', 'expected start'],
    finish: ['expectedshopcomplete', 'expected shop complete', 'shop complete', 'expected complete'],
    sir: ['sir'],
    quote: ['quote'],
    title: ['title', 'project title'],
  }
  const probabilityCandidates = ['probability', 'prob', 'chance']

  function getNumeric(row: Record<string, unknown>, candidates: string[]): number {
    const value = findCellValue(row, candidates)
    const parsed = parseWorkHours(value)
    return parsed ?? 0
  }

  function getDate(row: Record<string, unknown>, candidates: string[]): Date | null {
    return parseExcelDate(findCellValue(row, candidates))
  }

  return rawRows
    .map((row, index) => {
      const fabHours = getNumeric(row, salesColumnCandidates.fab)
      const shippingHours = getNumeric(row, salesColumnCandidates.shipping)
      const blastHours = getNumeric(row, salesColumnCandidates.blast)
      const paintHours = getNumeric(row, salesColumnCandidates.paint)
      // Sales Production Report hours are defined as the sum of Fab + Shipping + Blast + Paint.
      const totalHours = fabHours + shippingHours + blastHours + paintHours

      const start = getDate(row, salesColumnCandidates.start)
      const finish = getDate(row, salesColumnCandidates.finish)

      if (!totalHours || totalHours <= 0 || !start || !finish) {
        return null
      }

      const sir = findCellValue(row, salesColumnCandidates.sir)
      const quote = findCellValue(row, salesColumnCandidates.quote)
      const title = findCellValue(row, salesColumnCandidates.title)
      const probabilityValue = findCellValue(row, probabilityCandidates)
      const probability = parseProbability(probabilityValue)

      const pieces = [sir, quote, title]
        .map((value) => (typeof value === 'string' && value.trim() ? value.trim() : ''))
        .filter(Boolean)
      const projectName = pieces.length > 0 ? pieces.join(' - ') : `Sales Project ${index + 1}`

      const task: TaskRow = {
        id: `sales-${index}-${projectName}`,
        name: projectName,
        project: projectName,
        resourceName: 'Sales',
        workHours: totalHours,
        start,
        finish,
        salesProbability: probability,
      }

      return task
    })
    .filter((task): task is TaskRow => task !== null)
}
