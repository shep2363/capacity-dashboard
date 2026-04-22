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

// ── Office Script pre-processing ──────────────────────────────────────────
// Replicates the CleanupWorkbook Office Script logic client-side so the
// workbook is normalised before the flexible column parser runs.

const BLOCKED_RESOURCES = new Set(['Inventory', 'Purchasing', 'Detailing'])

function worksheetText(ws: XLSX.WorkSheet, r: number, c: number): string {
  const cell = ws[XLSX.utils.encode_cell({ r, c })]
  return cell?.v != null ? String(cell.v) : ''
}

function worksheetSetText(ws: XLSX.WorkSheet, r: number, c: number, v: string): void {
  const addr = XLSX.utils.encode_cell({ r, c })
  ws[addr] = { ...(ws[addr] ?? {}), t: 's', v, w: v }
}

function worksheetFontBold(ws: XLSX.WorkSheet, r: number, c: number): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as any
  return cell?.s?.font?.bold === true
}

function worksheetFontStrike(ws: XLSX.WorkSheet, r: number, c: number): boolean {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cell = ws[XLSX.utils.encode_cell({ r, c })] as any
  return cell?.s?.font?.strike === true
}

/**
 * Applies the same transformations as the CleanupWorkbook Office Script:
 *
 * 1. Normalise Project (col 0), Name (col 1), Hours (col 4), Resource (col 9).
 * 2. Delete rows whose Resource column is blank, bold, struck-through, or in
 *    the blocked list (Inventory, Purchasing, Detailing).
 * 3. Move Job # (col 0) to col 10, then shift left — net effect: col 0 moves
 *    to the last position and all other columns shift one place left.
 * 4. Delete the two Baseline columns that land at index 4 after the shift.
 *
 * Because the downstream parser uses flexible header-name detection (not fixed
 * column positions), step 3 and 4 do not affect parsing — they are included
 * for fidelity with the original script.
 */
function preprocessWorksheet(ws: XLSX.WorkSheet): XLSX.WorkSheet {
  if (!ws['!ref']) return ws

  const range = XLSX.utils.decode_range(ws['!ref'])
  const R0 = range.s.r
  const C0 = range.s.c
  const rowCount = range.e.r - R0 + 1
  const colCount = range.e.c - C0 + 1

  // ── Step 1: Normalise fields (bottom-up, matching Office Script iteration) ──
  for (let i = R0 + rowCount - 1; i >= R0; i--) {
    worksheetSetText(ws, i, C0 + 0, worksheetText(ws, i, C0 + 0).substring(0, 5))
    worksheetSetText(ws, i, C0 + 9, worksheetText(ws, i, C0 + 9).split('[')[0].trim())
    worksheetSetText(ws, i, C0 + 4, worksheetText(ws, i, C0 + 4).split(' ')[0])
    worksheetSetText(ws, i, C0 + 1, worksheetText(ws, i, C0 + 1).trim())
  }

  // ── Step 2: Collect rows to keep ────────────────────────────────────────────
  const keepRows: number[] = []
  for (let i = R0; i < R0 + rowCount; i++) {
    const resourceVal = worksheetText(ws, i, C0 + 9).trim()
    const shouldDelete =
      resourceVal === '' ||
      worksheetFontBold(ws, i, C0 + 9) ||
      worksheetFontStrike(ws, i, C0 + 9) ||
      BLOCKED_RESOURCES.has(resourceVal)
    if (!shouldDelete) keepRows.push(i)
  }

  // ── Step 3: Rebuild worksheet with only the kept rows ──────────────────────
  const out: XLSX.WorkSheet = {}
  keepRows.forEach((oldRow, newRow) => {
    for (let c = C0; c <= range.e.c; c++) {
      const src = XLSX.utils.encode_cell({ r: oldRow, c })
      const dst = XLSX.utils.encode_cell({ r: newRow, c })
      if (ws[src]) out[dst] = { ...ws[src] }
    }
  })

  const newRowCount = keepRows.length

  // ── Step 4: Move col 0 → col 10, then shift cols 1-10 left by 1 ───────────
  // (mirrors: moveTo col 10, then delete col 0 with shift-left)
  for (let r = 0; r < newRowCount; r++) {
    // Copy col C0 → col C0+10
    const src0 = XLSX.utils.encode_cell({ r, c: C0 })
    const dst10 = XLSX.utils.encode_cell({ r, c: C0 + 10 })
    if (out[src0]) out[dst10] = { ...out[src0] }

    // Shift cols C0+1 … C0+10 one place left (fills the gap left by col 0)
    for (let c = C0; c < C0 + 10; c++) {
      const src = XLSX.utils.encode_cell({ r, c: c + 1 })
      const dst = XLSX.utils.encode_cell({ r, c })
      if (out[src]) out[dst] = { ...out[src] }
      else delete out[dst]
    }
    delete out[XLSX.utils.encode_cell({ r, c: C0 + 10 })]
  }

  // ── Step 5: Delete baseline cols at index 4 (twice) ────────────────────────
  // After step 4 the sheet still has colCount columns (0-indexed 0 … colCount-1).
  let currentLastC = colCount - 1 // last column index (0-based from C0)
  for (let pass = 0; pass < 2; pass++) {
    const delC = C0 + 4
    for (let r = 0; r < newRowCount; r++) {
      for (let c = delC; c < C0 + currentLastC; c++) {
        const src = XLSX.utils.encode_cell({ r, c: c + 1 })
        const dst = XLSX.utils.encode_cell({ r, c })
        if (out[src]) out[dst] = { ...out[src] }
        else delete out[dst]
      }
      delete out[XLSX.utils.encode_cell({ r, c: C0 + currentLastC })]
    }
    currentLastC--
  }

  // Update the sheet's used-range reference
  out['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: C0 },
    e: { r: newRowCount - 1, c: C0 + currentLastC },
  })

  return out
}

// ── Main spreadsheet parser ───────────────────────────────────────────────

export function parseSpreadsheet(arrayBuffer: ArrayBuffer): TaskRow[] {
  // cellStyles: true is required so preprocessWorksheet can read bold / strikethrough
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true, cellStyles: true })
  const firstSheetName = workbook.SheetNames[0]

  if (!firstSheetName) {
    return []
  }

  const worksheet = preprocessWorksheet(workbook.Sheets[firstSheetName])
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
      const workHours = parseWorkHours(workValue) ?? 0
      const start = parseExcelDate(startValue)
      const finish = parseExcelDate(finishValue)
      const resourceName =
        typeof resourceValue === 'string' && resourceValue.trim() ? resourceValue.trim() : 'Unassigned'

      // Milestones (Work = 0) are allowed through only when they have a valid resource and dates.
      // Tasks with no hours and no resource are header/summary rows — skip them.
      if (workHours < 0 || !start || !finish) {
        return null
      }
      if (workHours === 0 && resourceName === 'Unassigned') {
        return null
      }

      return {
        id: `${index}-${name}`,
        name,
        workHours,
        start,
        finish,
        resourceName,
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
