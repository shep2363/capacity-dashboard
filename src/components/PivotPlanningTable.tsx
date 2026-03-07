import clsx from 'clsx'
import { useRef, useState } from 'react'
import type { PivotTableModel, PivotRowGrouping } from '../types'
import { shortWeekLabel } from '../utils/planner'

interface PivotPlanningTableProps {
  model: PivotTableModel
  rowGrouping: PivotRowGrouping
  overCapacityWeeks: Set<string>
  visibleWeekKeys: string[]
  weekWindowLabel: string
  canPageBack: boolean
  canPageForward: boolean
  onPageBack: () => void
  onPageForward: () => void
  weekWindowSize: number
  onWeekWindowSizeChange: (size: number) => void
  isCollapsed: boolean
  onToggleCollapsed: () => void
  onEditCell: (rowKey: string, weekStartIso: string, newValue: number) => void
  onResetEdits: () => void
  title?: string
  subtitle?: string
}

export function PivotPlanningTable({
  model,
  rowGrouping,
  overCapacityWeeks,
  visibleWeekKeys,
  weekWindowLabel,
  canPageBack,
  canPageForward,
  onPageBack,
  onPageForward,
  weekWindowSize,
  onWeekWindowSizeChange,
  isCollapsed,
  onToggleCollapsed,
  onEditCell,
  onResetEdits,
  title = 'Pivot Planning Table',
  subtitle = 'Editable planning grid. Cell edits become the forecast source of truth.',
}: PivotPlanningTableProps) {
  const [editingCell, setEditingCell] = useState<{ rowKey: string; weekStartIso: string } | null>(null)
  const [draftValue, setDraftValue] = useState('')
  const skipBlurSaveRef = useRef(false)

  function startEditing(rowKey: string, weekStartIso: string, currentValue: number): void {
    setEditingCell({ rowKey, weekStartIso })
    setDraftValue(currentValue.toFixed(2))
    skipBlurSaveRef.current = false
  }

  function cancelEditing(): void {
    skipBlurSaveRef.current = true
    setEditingCell(null)
    setDraftValue('')
  }

  function saveEditing(rowKey: string, weekStartIso: string): void {
    const trimmed = draftValue.trim()
    const parsed = trimmed === '' ? 0 : Number(trimmed)
    if (!Number.isFinite(parsed) || parsed < 0) {
      setEditingCell(null)
      setDraftValue('')
      return
    }

    onEditCell(rowKey, weekStartIso, parsed)
    setEditingCell(null)
    setDraftValue('')
  }

  return (
    <section className="panel pivot-panel">
      <div className="section-header section-header-row">
        <div>
          <h2>{title}</h2>
          <p>
            {subtitle} Showing {weekWindowLabel}.
          </p>
        </div>
        <div className="section-actions">
          <label className="inline-field">
            Weeks Visible
            <select value={weekWindowSize} onChange={(event) => onWeekWindowSizeChange(Number(event.target.value))}>
              <option value={8}>8</option>
              <option value={12}>12</option>
              <option value={16}>16</option>
            </select>
          </label>
          <button type="button" className="ghost-btn" onClick={onPageBack} disabled={!canPageBack}>
            Previous Weeks
          </button>
          <button type="button" className="ghost-btn" onClick={onPageForward} disabled={!canPageForward}>
            Next Weeks
          </button>
          <button type="button" className="ghost-btn" onClick={onToggleCollapsed}>
            {isCollapsed ? 'Expand Pivot' : 'Collapse Pivot'}
          </button>
          <button type="button" className="ghost-btn" onClick={onResetEdits}>
            Reset Manual Edits
          </button>
        </div>
      </div>

      {!isCollapsed && (
        <div className="pivot-wrap">
          <table className="pivot-table">
            <thead>
              <tr>
                <th className="sticky-col">{rowGrouping === 'project' ? 'Project' : 'Resource'}</th>
                {visibleWeekKeys.map((week) => (
                  <th key={week} className={clsx({ 'over-week': overCapacityWeeks.has(week) })}>
                    {shortWeekLabel(week)}
                  </th>
                ))}
                <th>Row Total</th>
              </tr>
            </thead>
            <tbody>
              {model.rows.map((row) => (
                <tr key={row.rowKey}>
                  <td className="sticky-col row-label">{row.rowLabel}</td>
                  {visibleWeekKeys.map((week) => {
                    const value = row.valuesByWeek[week] ?? 0
                    const edited = model.editedRowWeekKeys.has(`${row.rowKey}\u0001${week}`)
                    const isEditing = editingCell?.rowKey === row.rowKey && editingCell.weekStartIso === week
                    return (
                      <td
                        key={`${row.rowKey}-${week}`}
                        className={clsx({
                          'over-week': overCapacityWeeks.has(week),
                          edited,
                          editing: isEditing,
                        })}
                      >
                        {isEditing ? (
                          <input
                            autoFocus
                            type="text"
                            inputMode="decimal"
                            value={draftValue}
                            onChange={(event) => {
                              const next = event.target.value
                              if (/^\d*\.?\d*$/.test(next)) {
                                setDraftValue(next)
                              }
                            }}
                            onBlur={() => {
                              if (skipBlurSaveRef.current) {
                                skipBlurSaveRef.current = false
                                return
                              }
                              saveEditing(row.rowKey, week)
                            }}
                            onKeyDown={(event) => {
                              if (event.key === 'Enter') {
                                event.preventDefault()
                                saveEditing(row.rowKey, week)
                              }
                              if (event.key === 'Escape') {
                                event.preventDefault()
                                cancelEditing()
                              }
                            }}
                            aria-label={`Edit hours for ${row.rowLabel} in week ${week}`}
                          />
                        ) : (
                          <button
                            type="button"
                            className="cell-display"
                            onClick={() => startEditing(row.rowKey, week, value)}
                            aria-label={`Edit hours for ${row.rowLabel} in week ${week}`}
                          >
                            {value.toFixed(2)}
                          </button>
                        )}
                      </td>
                    )
                  })}
                  <td className="row-total">{row.totalHours.toFixed(2)}</td>
                </tr>
              ))}
            </tbody>
            <tfoot>
              <tr>
                <th className="sticky-col">Column Total</th>
                {visibleWeekKeys.map((week) => (
                  <th key={`total-${week}`} className={clsx({ 'over-week': overCapacityWeeks.has(week) })}>
                    {(model.columnTotals[week] ?? 0).toFixed(2)}
                  </th>
                ))}
                <th>{model.grandTotal.toFixed(2)}</th>
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </section>
  )
}
