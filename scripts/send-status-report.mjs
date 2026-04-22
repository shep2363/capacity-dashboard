/**
 * Weekly Status Report mailer
 *
 * Required GitHub secret:
 *   EMAIL_PASSWORD  – password for colden.sheppard@inframod.com (Microsoft 365)
 *
 * Optional GitHub secrets:
 *   WORKBOOK_API_URL – base URL of the shared-data API (e.g. https://your-api.vercel.app)
 *                      When set, the script downloads the live workbook, builds the Status
 *                      Report Excel file, and attaches it to the email.
 *   DASHBOARD_URL    – public URL of the dashboard (used in the email body link)
 */

import nodemailer from 'nodemailer'
import * as XLSX from 'xlsx'
import { format, isAfter, startOfDay } from 'date-fns'

// ── config ─────────────────────────────────────────────────────────────────

const FROM = 'colden.sheppard@inframod.com'
const DISPLAY_NAME = 'Colden Sheppard | infraMOD'

const RECIPIENTS = [
  'maciej.jedrzejowski@inframod.com',
  'jeffrey.reboya@inframod.com',
  'sandro.penner@inframod.com',
  'dylan.barnes@inframod.com',
  'austin.blezy@inframod.com',
  'sergii.vorotylo@inframod.com',
  'john.pondang@inframod.com',
]

const EMAIL_PASSWORD  = process.env.EMAIL_PASSWORD
const WORKBOOK_API_URL = (process.env.WORKBOOK_API_URL ?? '').replace(/\/$/, '')
const DASHBOARD_URL   = process.env.DASHBOARD_URL || 'https://capacity-dashboard.vercel.app'

// ── workbook parsing (mirrors src/utils/excel.ts) ──────────────────────────

function normalizeKey(s) {
  return s.toLowerCase().replace(/[^a-z0-9]/g, '')
}

function findCell(row, candidates) {
  const entries = Object.entries(row).map(([k, v]) => ({ norm: normalizeKey(k), v }))
  for (const c of candidates) {
    const t = normalizeKey(c)
    const hit = entries.find((e) => e.norm === t) ?? entries.find((e) => e.norm.startsWith(t))
    if (hit) return hit.v
  }
  return undefined
}

function parseNum(val) {
  if (typeof val === 'number' && isFinite(val)) return val
  if (typeof val === 'string') {
    const n = parseFloat(val.replace(/[^\d.-]/g, ''))
    return isFinite(n) ? n : null
  }
  return null
}

function parseDate(val) {
  if (val instanceof Date && !isNaN(val.getTime())) return val
  if (typeof val === 'number' && isFinite(val)) {
    const d = XLSX.SSF.parse_date_code(val)
    return d ? new Date(d.y, d.m - 1, d.d) : null
  }
  if (typeof val === 'string' && val.trim()) {
    const d = new Date(val)
    return isNaN(d.getTime()) ? null : d
  }
  return null
}

function parseTasks(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false })

  return rows
    .map((row, i) => {
      const nameVal = findCell(row, ['name', 'task name'])
      const name = typeof nameVal === 'string' && nameVal.trim() ? nameVal.trim() : `Task ${i + 1}`
      const workHours = parseNum(findCell(row, ['work', 'hours', 'estimated work'])) ?? 0
      const start = parseDate(findCell(row, ['start', 'start date']))
      const finish = parseDate(findCell(row, ['finish', 'end', 'finish date', 'due date']))
      const resVal = findCell(row, ['resource names', 'resource', 'resources', 'assignee'])
      const resourceName = typeof resVal === 'string' && resVal.trim() ? resVal.trim() : 'Unassigned'
      const projVal = findCell(row, ['project', 'proje', 'project name'])
      const project = typeof projVal === 'string' && projVal.trim() ? projVal.trim() : 'Unspecified'

      if (workHours < 0 || !start || !finish) return null
      if (workHours === 0 && resourceName === 'Unassigned') return null

      return { name, workHours, start, finish, resourceName, project }
    })
    .filter(Boolean)
}

// ── report generation (mirrors exportStatusReport in App.tsx) ───────────────

function generateReport(tasks) {
  const wb = XLSX.utils.book_new()
  const now = startOfDay(new Date())
  const headers = ['Project', 'Sequence', 'Percent Complete', 'Start Date', 'Finish Date']

  const autoWidth = (rows) => {
    const widths = []
    rows.forEach((row) =>
      row.forEach((cell, i) => {
        widths[i] = Math.max(widths[i] ?? 0, Math.min(60, String(cell ?? '').length + 2))
      }),
    )
    return widths.map((wch) => ({ wch }))
  }

  const tabs = [
    { label: 'Detailing',   resourceName: 'Detailer' },
    { label: 'Processing',  resourceName: 'Processing' },
    { label: 'Fabrication', resourceName: 'Fabrication' },
    { label: 'Shipping',    resourceName: 'Shipping' },
  ]

  tabs.forEach(({ label, resourceName }) => {
    const filtered = tasks.filter((t) => t.resourceName === resourceName)
    let dataRows

    if (resourceName === 'Detailer') {
      // Exclude milestones whose finish date has already passed
      dataRows = filtered
        .filter((t) => !isAfter(now, t.finish))
        .map((t) => [
          t.project,
          t.name,
          '—',
          format(t.start, 'MM/dd/yyyy'),
          format(t.finish, 'MM/dd/yyyy'),
        ])
    } else {
      // Percent complete comes from Smartsheet (not available server-side); show placeholder
      dataRows = filtered.map((t) => [
        t.project,
        t.name,
        'See Dashboard',
        format(t.start, 'MM/dd/yyyy'),
        format(t.finish, 'MM/dd/yyyy'),
      ])
    }

    const aoa = [headers, ...dataRows]
    const sheet = XLSX.utils.aoa_to_sheet(aoa)
    sheet['!cols'] = autoWidth(aoa)
    sheet['!freeze'] = { xSplit: 0, ySplit: 1 }
    sheet['!autofilter'] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: 0, c: headers.length - 1 },
      }),
    }
    XLSX.utils.book_append_sheet(wb, sheet, label)
  })

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })
}

// ── email body ──────────────────────────────────────────────────────────────

function buildEmailHtml(dateStr, hasAttachment) {
  const reportNote = hasAttachment
    ? `<p>Please find this week's <strong>Status Report</strong> attached (Excel file with Detailing, Processing, Fabrication, and Shipping tabs).</p>`
    : `<p>To download this week's <strong>Status Report</strong>, log into the dashboard and click <strong>Export Status Report</strong>.</p>`

  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;color:#1f2937;line-height:1.6;max-width:600px;margin:0 auto;padding:20px">
  <div style="background:#0b1625;padding:20px 24px;border-radius:8px 8px 0 0">
    <h2 style="color:#a5b4fc;margin:0">infraMOD Production Dashboard</h2>
    <p style="color:#94a3b8;margin:4px 0 0">Weekly Status Report — ${dateStr}</p>
  </div>

  <div style="border:1px solid #e5e7eb;border-top:none;padding:24px;border-radius:0 0 8px 8px">
    <p>Hi team,</p>

    ${reportNote}

    <p>You can also view live capacity and schedule data on the dashboard:</p>

    <p style="text-align:center;margin:24px 0">
      <a href="${DASHBOARD_URL}"
         style="background:#6366f1;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:600">
        Open Dashboard
      </a>
    </p>

    <p style="color:#64748b;font-size:0.9rem">
      This email is sent automatically every Friday at 1:00 PM CST.<br>
      Reply to this email or contact Colden directly with any questions.
    </p>

    <p>Best,<br><strong>Colden Sheppard</strong><br>infraMOD</p>
  </div>
</body>
</html>`
}

// ── main ───────────────────────────────────────────────────────────────────

async function main() {
  if (!EMAIL_PASSWORD) {
    throw new Error('EMAIL_PASSWORD secret is not set — cannot send email.')
  }

  const dateStr = format(new Date(), 'MMMM d, yyyy')
  const attachments = []

  // Try to download the workbook and generate the Excel attachment
  if (WORKBOOK_API_URL) {
    try {
      console.log(`Downloading workbook from ${WORKBOOK_API_URL}/api/workbook-file?dataset=main …`)
      const res = await fetch(`${WORKBOOK_API_URL}/api/workbook-file?dataset=main`)
      if (!res.ok) throw new Error(`API responded ${res.status}`)
      const buf = Buffer.from(await res.arrayBuffer())
      const tasks = parseTasks(buf)
      console.log(`Parsed ${tasks.length} tasks from workbook`)
      const reportBuf = generateReport(tasks)
      attachments.push({
        filename: `Status_Report_${format(new Date(), 'MM-dd-yyyy')}.xlsx`,
        content: reportBuf,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
      console.log('Status Report Excel generated and attached.')
    } catch (err) {
      console.warn(`Could not generate report attachment: ${err.message}`)
      console.warn('Sending notification-only email instead.')
    }
  } else {
    console.log('WORKBOOK_API_URL not configured — sending notification-only email.')
  }

  const transporter = nodemailer.createTransport({
    host: 'smtp.office365.com',
    port: 587,
    secure: false, // STARTTLS
    auth: { user: FROM, pass: EMAIL_PASSWORD },
  })

  await transporter.verify()
  console.log('SMTP connection verified.')

  const info = await transporter.sendMail({
    from: `${DISPLAY_NAME} <${FROM}>`,
    to: RECIPIENTS.join(', '),
    subject: `Weekly Status Report — ${dateStr}`,
    html: buildEmailHtml(dateStr, attachments.length > 0),
    attachments,
  })

  console.log(`Email sent to ${RECIPIENTS.length} recipients. Message ID: ${info.messageId}`)
}

main().catch((err) => {
  console.error('Fatal error:', err.message)
  process.exit(1)
})
