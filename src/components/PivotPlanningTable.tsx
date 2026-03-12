import clsx from 'clsx'
import { useEffect, useMemo, useRef, useState } from 'react'
import type { PivotTableModel, PivotRowGrouping } from '../types'
import { shortWeekLabel } from '../utils/planner'

const ALWAYS_SELECTABLE = () => true

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
  isCellSelectable?: (rowKey: string, weekStartIso: string, value: number) => boolean
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
  isCellSelectable = ALWAYS_SELECTABLE,
  title = 'Pivot Planning Table',
  subtitle = 'Editable planning grid. Cell edits become the forecast source of truth.',
}: PivotPlanningTableProps) {
  type CellRef = { rowKey: string; weekStartIso: string }

  const [editingCell, setEditingCell] = useState<{ rowKey: string; weekStartIso: string } | null>(null)
  const [draftValue, setDraftValue] = useState('')
  const skipBlurSaveRef = useRef(false)
  const skipNextClickRef = useRef(false)
  const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set())
  const [dragAnchorCell, setDragAnchorCell] = useState<CellRef | null>(null)
  const [dragCurrentCell, setDragCurrentCell] = useState<CellRef | null>(null)
  const [copyFeedback, setCopyFeedback] = useState('')

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

  function toggleSelectedCell(rowKey: string, weekStartIso: string): void {
    setSelectedCells((current) => {
      const next = new Set(current)
      const key = `${rowKey}|${weekStartIso}`
      if (next.has(key)) {
        next.delete(key)
      } else {
        next.add(key)
      }
      return next
    })
  }

  function clearSelection(): void {
    setSelectedCells(new Set())
  }

  const rowIndexByKey = useMemo(() => {
    const byRow = new Map<string, number>()
    model.rows.forEach((row, index) => {
      byRow.set(row.rowKey, index)
    })
    return byRow
  }, [model.rows])

  const weekIndexByKey = useMemo(() => {
    const byWeek = new Map<string, number>()
    visibleWeekKeys.forEach((weekStartIso, index) => {
      byWeek.set(weekStartIso, index)
    })
    return byWeek
  }, [visibleWeekKeys])

  const selectableCellKeysByWeek = useMemo(() => {
    const byWeek: Record<string, string[]> = {}
    visibleWeekKeys.forEach((weekStartIso) => {
      byWeek[weekStartIso] = model.rows
        .filter((row) => isCellSelectable(row.rowKey, weekStartIso, row.valuesByWeek[weekStartIso] ?? 0))
        .map((row) => `${row.rowKey}|${weekStartIso}`)
    })
    return byWeek
  }, [isCellSelectable, model.rows, visibleWeekKeys])

  function buildRangeSelection(anchor: CellRef, current: CellRef): Set<string> {
    const anchorRowIndex = rowIndexByKey.get(anchor.rowKey)
    const anchorWeekIndex = weekIndexByKey.get(anchor.weekStartIso)
    const currentRowIndex = rowIndexByKey.get(current.rowKey)
    const currentWeekIndex = weekIndexByKey.get(current.weekStartIso)

    if (
      anchorRowIndex === undefined ||
      anchorWeekIndex === undefined ||
      currentRowIndex === undefined ||
      currentWeekIndex === undefined
    ) {
      return new Set()
    }

    const minRow = Math.min(anchorRowIndex, currentRowIndex)
    const maxRow = Math.max(anchorRowIndex, currentRowIndex)
    const minWeek = Math.min(anchorWeekIndex, currentWeekIndex)
    const maxWeek = Math.max(anchorWeekIndex, currentWeekIndex)

    const next = new Set<string>()
    for (let rowIndex = minRow; rowIndex <= maxRow; rowIndex += 1) {
      const row = model.rows[rowIndex]
      if (!row) continue
      for (let weekIndex = minWeek; weekIndex <= maxWeek; weekIndex += 1) {
        const weekStartIso = visibleWeekKeys[weekIndex]
        if (!weekStartIso) continue
        const value = row.valuesByWeek[weekStartIso] ?? 0
        if (!isCellSelectable(row.rowKey, weekStartIso, value)) {
          continue
        }
        next.add(`${row.rowKey}|${weekStartIso}`)
      }
    }

    return next
  }

  function startDragSelection(rowKey: string, weekStartIso: string): void {
    const anchor: CellRef = { rowKey, weekStartIso }
    setDragAnchorCell(anchor)
    setDragCurrentCell(anchor)
    setSelectedCells(buildRangeSelection(anchor, anchor))
  }

  function updateDragSelection(rowKey: string, weekStartIso: string): void {
    if (!dragAnchorCell) return
    const current: CellRef = { rowKey, weekStartIso }
    setDragCurrentCell(current)
    setSelectedCells(buildRangeSelection(dragAnchorCell, current))
  }

  function finishDragSelection(): void {
    setDragAnchorCell(null)
    setDragCurrentCell(null)
  }

  function buildSelectableWeekCellKeys(weekStartIso: string): string[] {
    return selectableCellKeysByWeek[weekStartIso] ?? []
  }

  function toggleSelectWeek(weekStartIso: string): void {
    setSelectedCells((current) => {
      const next = new Set(current)
      const weekKeys = buildSelectableWeekCellKeys(weekStartIso)
      if (weekKeys.length === 0) {
        return next
      }
      const allSelected = weekKeys.every((key) => next.has(key))
      weekKeys.forEach((key) => {
        if (allSelected) {
          next.delete(key)
        } else {
          next.add(key)
        }
      })
      return next
    })
  }

  function getSelectedWeekCount(weekStartIso: string): number {
    const selectableWeekKeys = buildSelectableWeekCellKeys(weekStartIso)
    if (selectableWeekKeys.length === 0) {
      return 0
    }
    let selectedCountForWeek = 0
    selectableWeekKeys.forEach((key) => {
      if (selectedCells.has(key)) {
        selectedCountForWeek += 1
      }
    })
    return selectedCountForWeek
  }

  function isWeekFullySelected(weekStartIso: string): boolean {
    const selectableWeekKeys = buildSelectableWeekCellKeys(weekStartIso)
    return selectableWeekKeys.length > 0 && selectableWeekKeys.every((key) => selectedCells.has(key))
  }

  const weeksWithPlannedWork = useMemo(() => {
    return visibleWeekKeys.filter((weekStartIso) =>
      model.rows.some((row) => {
        const value = row.valuesByWeek[weekStartIso] ?? 0
        return value > 0 && isCellSelectable(row.rowKey, weekStartIso, value)
      }),
    )
  }, [isCellSelectable, model.rows, visibleWeekKeys])

  const plannedWeekCellKeys = useMemo(() => {
    const keys = new Set<string>()
    weeksWithPlannedWork.forEach((weekStartIso) => {
      buildSelectableWeekCellKeys(weekStartIso).forEach((key) => keys.add(key))
    })
    return keys
  }, [selectableCellKeysByWeek, weeksWithPlannedWork])

  function arePlannedWeeksFullySelected(): boolean {
    return plannedWeekCellKeys.size > 0 && [...plannedWeekCellKeys].every((key) => selectedCells.has(key))
  }

  function toggleSelectPlannedWeeks(): void {
    setSelectedCells((current) => {
      if (plannedWeekCellKeys.size === 0) {
        return current
      }
      const allSelected = [...plannedWeekCellKeys].every((key) => current.has(key))
      const next = new Set(current)
      plannedWeekCellKeys.forEach((key) => {
        if (allSelected) {
          next.delete(key)
        } else {
          next.add(key)
        }
      })
      return next
    })
  }

  useEffect(() => {
    if (!dragAnchorCell) {
      return
    }

    function handleGlobalMouseUp(): void {
      finishDragSelection()
    }

    window.addEventListener('mouseup', handleGlobalMouseUp)
    return () => {
      window.removeEventListener('mouseup', handleGlobalMouseUp)
    }
  }, [dragAnchorCell])

  useEffect(() => {
    setSelectedCells((current) => {
      const next = new Set<string>()
      current.forEach((key) => {
        const [rowKey, weekStartIso] = key.split('|')
        const rowIndex = rowIndexByKey.get(rowKey)
        const weekIndex = weekIndexByKey.get(weekStartIso)
        if (rowIndex === undefined || weekIndex === undefined) {
          return
        }
        const row = model.rows[rowIndex]
        const value = row.valuesByWeek[weekStartIso] ?? 0
        if (!isCellSelectable(rowKey, weekStartIso, value)) {
          return
        }
        next.add(key)
      })

      if (next.size === current.size) {
        let identical = true
        next.forEach((key) => {
          if (!current.has(key)) {
            identical = false
          }
        })
        if (identical) {
          return current
        }
      }
      return next
    })
  }, [isCellSelectable, model.rows, rowIndexByKey, visibleWeekKeys, weekIndexByKey])

  useEffect(() => {
    if (!copyFeedback) {
      return
    }
    const timer = window.setTimeout(() => setCopyFeedback(''), 1800)
    return () => window.clearTimeout(timer)
  }, [copyFeedback])

  async function copyTextToClipboard(text: string): Promise<void> {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text)
      return
    }

    const textarea = document.createElement('textarea')
    textarea.value = text
    textarea.setAttribute('readonly', 'true')
    textarea.style.position = 'fixed'
    textarea.style.opacity = '0'
    document.body.appendChild(textarea)
    textarea.select()
    document.execCommand('copy')
    document.body.removeChild(textarea)
  }

  async function copySelectionForExcel(): Promise<void> {
    const selectedCoordinates: Array<{ rowIndex: number; weekIndex: number }> = []

    selectedCells.forEach((key) => {
      const [rowKey, weekStartIso] = key.split('|')
      const rowIndex = rowIndexByKey.get(rowKey)
      const weekIndex = weekIndexByKey.get(weekStartIso)
      if (rowIndex === undefined || weekIndex === undefined) return

      const row = model.rows[rowIndex]
      const value = row.valuesByWeek[weekStartIso] ?? 0
      if (!isCellSelectable(rowKey, weekStartIso, value)) return
      selectedCoordinates.push({ rowIndex, weekIndex })
    })

    if (selectedCoordinates.length === 0) {
      setCopyFeedback('Nothing selected to copy.')
      return
    }

    const rowIndexes = selectedCoordinates.map((item) => item.rowIndex)
    const weekIndexes = selectedCoordinates.map((item) => item.weekIndex)
    const minRow = Math.min(...rowIndexes)
    const maxRow = Math.max(...rowIndexes)
    const minWeek = Math.min(...weekIndexes)
    const maxWeek = Math.max(...weekIndexes)

    const lines: string[] = []
    for (let rowIndex = minRow; rowIndex <= maxRow; rowIndex += 1) {
      const row = model.rows[rowIndex]
      if (!row) continue
      const values: string[] = []
      for (let weekIndex = minWeek; weekIndex <= maxWeek; weekIndex += 1) {
        const weekStartIso = visibleWeekKeys[weekIndex]
        if (!weekStartIso) {
          values.push('')
          continue
        }
        const key = `${row.rowKey}|${weekStartIso}`
        const rawValue = row.valuesByWeek[weekStartIso] ?? 0
        const selectable = isCellSelectable(row.rowKey, weekStartIso, rawValue)
        values.push(selectedCells.has(key) && selectable ? rawValue.toFixed(2) : '')
      }
      lines.push(values.join('\t'))
    }

    const output = lines.join('\n')
    try {
      await copyTextToClipboard(output)
      setCopyFeedback('Copied selection for Excel.')
    } catch {
      setCopyFeedback('Copy failed.')
    }
  }

  function handleCellMouseDown(
    event: React.MouseEvent<HTMLButtonElement>,
    rowKey: string,
    weekStartIso: string,
    cellSelectable: boolean,
  ): void {
    if (event.button !== 0 || event.ctrlKey || event.metaKey || !cellSelectable) {
      return
    }
    event.preventDefault()
    skipNextClickRef.current = false
    startDragSelection(rowKey, weekStartIso)
  }

  function handleCellMouseEnter(rowKey: string, weekStartIso: string): void {
    if (!dragAnchorCell) {
      return
    }
    skipNextClickRef.current = true
    updateDragSelection(rowKey, weekStartIso)
  }

  function handleSectionKeyDown(event: React.KeyboardEvent<HTMLElement>): void {
    if (!selectedCells.size || !(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 'c') {
      return
    }

    const target = event.target as HTMLElement | null
    if (target && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable)) {
      return
    }

    event.preventDefault()
    void copySelectionForExcel()
  }

  const selectedTotal = useMemo(() => {
    let total = 0
    selectedCells.forEach((key) => {
      const [rowKey, weekStartIso] = key.split('|')
      const row = model.rows.find((r) => r.rowKey === rowKey)
      if (!row) return
      total += row.valuesByWeek[weekStartIso] ?? 0
    })
    return total
  }, [selectedCells, model.rows])
  const selectedCount = selectedCells.size

  return (
    <section
      className={clsx('panel pivot-panel', { 'pivot-drag-selecting': Boolean(dragAnchorCell || dragCurrentCell) })}
      onKeyDownCapture={handleSectionKeyDown}
    >
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
            <span className={`chevron ${isCollapsed ? 'chevron-closed' : 'chevron-open'}`} aria-hidden="true">
              ▾
            </span>
            {isCollapsed ? 'Expand Pivot' : 'Collapse Pivot'}
          </button>
          <button type="button" className="ghost-btn" onClick={onResetEdits}>
            Reset Manual Edits
          </button>
          <button
            type="button"
            className="ghost-btn"
            onClick={toggleSelectPlannedWeeks}
            disabled={weeksWithPlannedWork.length === 0}
            title="Select all visible weeks that contain planned hours"
          >
            {arePlannedWeeksFullySelected() ? 'Clear Planned Weeks' : 'Select Planned Weeks'} ({weeksWithPlannedWork.length})
          </button>
          <div className="inline-field" style={{ marginLeft: 12 }}>
            <span style={{ color: '#9ca3af', marginRight: 8 }}>
              Click+drag to select range. Ctrl+click adds cells. Ctrl/Cmd+C copies for Excel.
            </span>
          </div>
          {selectedCount > 0 && (
            <div
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 8,
                padding: '6px 10px',
                borderRadius: 12,
                background: '#0b1220',
                border: '1px solid #38bdf8',
                color: '#e5e7eb',
              }}
            >
              <span>Selected: {selectedCount}</span>
              <strong>{selectedTotal.toFixed(2)} hrs</strong>
              <button type="button" className="ghost-btn" onClick={() => void copySelectionForExcel()}>
                Copy
              </button>
              <button type="button" className="ghost-btn" onClick={clearSelection}>
                Clear
              </button>
            </div>
          )}
          {copyFeedback && (
            <div className="inline-field" style={{ marginLeft: 8 }}>
              <span style={{ color: '#cbd5e1' }}>{copyFeedback}</span>
            </div>
          )}
        </div>
      </div>

      {!isCollapsed && (
        <div className="pivot-wrap">
          {selectedCount > 0 && (
            <div
              style={{
                position: 'sticky',
                top: 0,
                zIndex: 2,
                marginBottom: 8,
                padding: '8px 12px',
                borderRadius: 10,
                background: '#0b1220',
                border: '1px solid #38bdf8',
                color: '#e5e7eb',
                display: 'inline-flex',
                alignItems: 'center',
                gap: 10,
              }}
            >
              <strong>Selection</strong>
              <span>{selectedCount} cell{selectedCount === 1 ? '' : 's'}</span>
              <span>{selectedTotal.toFixed(2)} hrs</span>
              <button type="button" className="ghost-btn" onClick={clearSelection}>
                Clear
              </button>
              <button type="button" className="ghost-btn" onClick={() => void copySelectionForExcel()}>
                Copy
              </button>
            </div>
          )}
          <table className="pivot-table">
            <thead>
              <tr>
                <th className="sticky-col">{rowGrouping === 'project' ? 'Project' : 'Resource'}</th>
                {visibleWeekKeys.map((week) => {
                  const weekSelectedCount = getSelectedWeekCount(week)
                  const weekSelectableCount = buildSelectableWeekCellKeys(week).length
                  const fullySelected = isWeekFullySelected(week)
                  const selectLabel = fullySelected ? 'Clear' : 'Select'
                  return (
                    <th key={week} className={clsx({ 'over-week': overCapacityWeeks.has(week) })}>
                      <div className="week-header-cell">
                        <span>{shortWeekLabel(week)}</span>
                        <button
                          type="button"
                          className={clsx('week-select-btn', { 'week-select-btn-active': fullySelected })}
                          onClick={() => toggleSelectWeek(week)}
                          disabled={weekSelectableCount === 0}
                          title={`${selectLabel} all visible rows for week ${week}`}
                          aria-label={`${selectLabel} all visible rows for week ${week}`}
                        >
                          {selectLabel}
                        </button>
                        <span className="week-select-count">
                          {weekSelectedCount}/{weekSelectableCount}
                        </span>
                      </div>
                    </th>
                  )
                })}
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
                    const cellKey = `${row.rowKey}|${week}`
                    const isSelected = selectedCells.has(cellKey)
                    const cellSelectable = isCellSelectable(row.rowKey, week, value)
                    return (
                      <td
                        key={`${row.rowKey}-${week}`}
                        className={clsx({
                          'over-week': overCapacityWeeks.has(week),
                          edited,
                          editing: isEditing,
                          selected: isSelected,
                        })}
                        style={
                          isSelected
                            ? {
                                outline: '2px solid #38bdf8',
                                outlineOffset: '-2px',
                                background: 'rgba(56, 189, 248, 0.18)',
                              }
                            : undefined
                        }
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
                            disabled={!cellSelectable}
                            onMouseDown={(event) => handleCellMouseDown(event, row.rowKey, week, cellSelectable)}
                            onMouseEnter={() => handleCellMouseEnter(row.rowKey, week)}
                            onClick={(event) => {
                              if (skipNextClickRef.current) {
                                skipNextClickRef.current = false
                                return
                              }
                              if (!cellSelectable) {
                                return
                              }
                              if (event.ctrlKey || event.metaKey) {
                                toggleSelectedCell(row.rowKey, week)
                                return
                              }
                              clearSelection()
                              startEditing(row.rowKey, week, value)
                            }}
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
              {selectedCount > 0 && (
                <tr>
                  <th className="sticky-col">Selected Total</th>
                  <th colSpan={visibleWeekKeys.length + 1} style={{ textAlign: 'left' }}>
                    {selectedTotal.toFixed(2)} hrs across {selectedCount} cell{selectedCount === 1 ? '' : 's'}
                  </th>
                </tr>
              )}
            </tfoot>
          </table>
        </div>
      )}
    </section>
  )
}
