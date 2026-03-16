import { parseISO } from 'date-fns'
import type { WorkbookDataset } from './activeWorkbookApi'
import type { RevenueRateMap } from './revenueRatesApi'
import { parseLeafKey, shortWeekLabel, weekRangeLabel } from './planner'

export const SALES_PROJECT_PREFIX = 'Sales - '

export interface RevenueRateRow {
  dataset: WorkbookDataset
  project: string
  label: string
  plannedHours: number
  revenuePerHour: number
  grossProfitPerHour: number
}

export interface WeeklyRevenueProjectDetail {
  projectLabel: string
  plannedHours: number
  revenuePerHour: number
  revenueAmount: number
}

export interface WeeklyRevenueRow {
  weekStartIso: string
  weekLabel: string
  weekRangeLabel: string
  totalPlannedHours: number
  totalRevenue: number
  amountsByProject: Record<string, number>
  details: WeeklyRevenueProjectDetail[]
}

export interface GrossProfitByProjectRow {
  projectLabel: string
  plannedHours: number
  grossProfitPerHour: number
  grossProfitAmount: number
}

interface BuildRevenueMetricsParams {
  mainFinalByKey: Record<string, number>
  salesFinalByKey: Record<string, number>
  mainWeekKeys: string[]
  salesWeekKeys: string[]
  mainRates: RevenueRateMap
  salesRates: RevenueRateMap
}

interface ProjectHoursByWeek {
  weekByProject: Map<string, Map<string, number>>
  totalByProject: Map<string, number>
  allWeeks: string[]
}

function sortWeekKeys(weekKeys: string[]): string[] {
  return [...weekKeys].sort((a, b) => parseISO(a).getTime() - parseISO(b).getTime())
}

function toNonNegativeNumber(value: unknown): number {
  const numeric = Number(value)
  if (!Number.isFinite(numeric) || numeric < 0) {
    return 0
  }
  return numeric
}

export function normalizeRateMap(rates: RevenueRateMap): RevenueRateMap {
  const normalized: RevenueRateMap = {}
  Object.entries(rates).forEach(([project, value]) => {
    if (!project) {
      return
    }
    normalized[project] = {
      revenuePerHour: toNonNegativeNumber(value?.revenuePerHour),
      grossProfitPerHour: toNonNegativeNumber(value?.grossProfitPerHour),
    }
  })
  return normalized
}

function revenueRateForLabel(projectLabel: string, mainRates: RevenueRateMap, salesRates: RevenueRateMap): number {
  if (projectLabel.startsWith(SALES_PROJECT_PREFIX)) {
    const raw = projectLabel.slice(SALES_PROJECT_PREFIX.length)
    return toNonNegativeNumber(salesRates[raw]?.revenuePerHour)
  }
  return toNonNegativeNumber(mainRates[projectLabel]?.revenuePerHour)
}

function grossProfitRateForLabel(projectLabel: string, mainRates: RevenueRateMap, salesRates: RevenueRateMap): number {
  if (projectLabel.startsWith(SALES_PROJECT_PREFIX)) {
    const raw = projectLabel.slice(SALES_PROJECT_PREFIX.length)
    return toNonNegativeNumber(salesRates[raw]?.grossProfitPerHour)
  }
  return toNonNegativeNumber(mainRates[projectLabel]?.grossProfitPerHour)
}

function collectProjectHoursByWeek(
  mainFinalByKey: Record<string, number>,
  salesFinalByKey: Record<string, number>,
  mainWeekKeys: string[],
  salesWeekKeys: string[],
): ProjectHoursByWeek {
  const mainWeekSet = new Set(mainWeekKeys)
  const salesWeekSet = new Set(salesWeekKeys)
  const allWeekKeys = sortWeekKeys([...new Set([...mainWeekKeys, ...salesWeekKeys])])
  const weekByProject = new Map<string, Map<string, number>>()
  const totalByProject = new Map<string, number>()

  const appendHours = (
    weekStartIso: string,
    projectLabel: string,
    hours: number,
  ): void => {
    const numericHours = toNonNegativeNumber(hours)
    if (numericHours <= 0) {
      return
    }
    const perProject = weekByProject.get(weekStartIso) ?? new Map<string, number>()
    perProject.set(projectLabel, (perProject.get(projectLabel) ?? 0) + numericHours)
    weekByProject.set(weekStartIso, perProject)
    totalByProject.set(projectLabel, (totalByProject.get(projectLabel) ?? 0) + numericHours)
  }

  Object.entries(mainFinalByKey).forEach(([leafKey, hours]) => {
    const { project, weekStartIso } = parseLeafKey(leafKey)
    if (!mainWeekSet.has(weekStartIso)) {
      return
    }
    appendHours(weekStartIso, project, hours)
  })

  Object.entries(salesFinalByKey).forEach(([leafKey, hours]) => {
    const { project, weekStartIso } = parseLeafKey(leafKey)
    if (!salesWeekSet.has(weekStartIso)) {
      return
    }
    appendHours(weekStartIso, `${SALES_PROJECT_PREFIX}${project}`, hours)
  })

  return { weekByProject, totalByProject, allWeeks: allWeekKeys }
}

export function buildRevenueRateRows(
  availableProjects: string[],
  salesAvailableProjects: string[],
  projectTotals: Map<string, number>,
  salesProjectTotals: Map<string, number>,
  mainRates: RevenueRateMap,
  salesRates: RevenueRateMap,
): RevenueRateRow[] {
  const rows: RevenueRateRow[] = []

  availableProjects.forEach((project) => {
    const rate = mainRates[project]
    rows.push({
      dataset: 'main',
      project,
      label: project,
      plannedHours: toNonNegativeNumber(projectTotals.get(project) ?? 0),
      revenuePerHour: toNonNegativeNumber(rate?.revenuePerHour),
      grossProfitPerHour: toNonNegativeNumber(rate?.grossProfitPerHour),
    })
  })

  salesAvailableProjects.forEach((project) => {
    const rate = salesRates[project]
    rows.push({
      dataset: 'sales',
      project,
      label: `${SALES_PROJECT_PREFIX}${project}`,
      plannedHours: toNonNegativeNumber(salesProjectTotals.get(project) ?? 0),
      revenuePerHour: toNonNegativeNumber(rate?.revenuePerHour),
      grossProfitPerHour: toNonNegativeNumber(rate?.grossProfitPerHour),
    })
  })

  return rows.sort((a, b) => a.label.localeCompare(b.label))
}

export function buildRevenueMetrics({
  mainFinalByKey,
  salesFinalByKey,
  mainWeekKeys,
  salesWeekKeys,
  mainRates,
  salesRates,
}: BuildRevenueMetricsParams): {
  weeklyRevenueRows: WeeklyRevenueRow[]
  weeklyProjectKeys: string[]
  grossProfitRows: GrossProfitByProjectRow[]
} {
  const { weekByProject, totalByProject, allWeeks } = collectProjectHoursByWeek(
    mainFinalByKey,
    salesFinalByKey,
    mainWeekKeys,
    salesWeekKeys,
  )

  const weeklyRevenueRows = allWeeks.map((weekStartIso) => {
    const projectHours = weekByProject.get(weekStartIso) ?? new Map<string, number>()
    const details = [...projectHours.entries()]
      .map(([projectLabel, plannedHours]) => {
        const revenuePerHour = revenueRateForLabel(projectLabel, mainRates, salesRates)
        const revenueAmount = plannedHours * revenuePerHour
        return { projectLabel, plannedHours, revenuePerHour, revenueAmount }
      })
      .sort((a, b) => b.revenueAmount - a.revenueAmount)

    const amountsByProject: Record<string, number> = {}
    details.forEach((detail) => {
      amountsByProject[detail.projectLabel] = detail.revenueAmount
    })
    const totalPlannedHours = details.reduce((sum, detail) => sum + detail.plannedHours, 0)
    const totalRevenue = details.reduce((sum, detail) => sum + detail.revenueAmount, 0)

    return {
      weekStartIso,
      weekLabel: shortWeekLabel(weekStartIso),
      weekRangeLabel: weekRangeLabel(weekStartIso),
      totalPlannedHours,
      totalRevenue,
      amountsByProject,
      details,
    }
  })

  const weeklyProjectKeys = [...totalByProject.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([projectLabel]) => projectLabel)

  const grossProfitRows = [...totalByProject.entries()]
    .map(([projectLabel, plannedHours]) => {
      const grossProfitPerHour = grossProfitRateForLabel(projectLabel, mainRates, salesRates)
      return {
        projectLabel,
        plannedHours,
        grossProfitPerHour,
        grossProfitAmount: plannedHours * grossProfitPerHour,
      }
    })
    .sort((a, b) => b.grossProfitAmount - a.grossProfitAmount)

  return { weeklyRevenueRows, weeklyProjectKeys, grossProfitRows }
}
