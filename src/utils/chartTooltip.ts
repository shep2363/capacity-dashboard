import type { CSSProperties } from 'react'

export interface TooltipPosition {
  x: number
  y: number
}

export interface RechartsTooltipSnapshot {
  isTooltipActive?: boolean
  activeCoordinate?: {
    x?: number
    y?: number
  }
  viewBox?: {
    x?: number
    y?: number
    width?: number
    height?: number
  }
}

interface TooltipSizing {
  width: number
  height: number
  offset?: number
  padding?: number
}

export const chartTooltipGlassStyle: CSSProperties = {
  background: 'rgba(8, 20, 36, 0.32)',
  border: '1px solid rgba(148, 163, 184, 0.18)',
  boxShadow: '0 14px 28px rgba(2, 8, 23, 0.18)',
  backdropFilter: 'blur(18px)',
  WebkitBackdropFilter: 'blur(18px)',
  color: '#f8fafc',
}

export const chartTooltipInsetStyle: CSSProperties = {
  background: 'rgba(15, 23, 42, 0.12)',
  border: '1px solid rgba(148, 163, 184, 0.1)',
  backdropFilter: 'blur(14px)',
  WebkitBackdropFilter: 'blur(14px)',
}

function clamp(value: number, min: number, max: number): number {
  if (max < min) {
    return min
  }
  return Math.min(Math.max(value, min), max)
}

export function resolveTooltipPosition(
  snapshot: RechartsTooltipSnapshot | null | undefined,
  sizing: TooltipSizing,
): TooltipPosition | undefined {
  if (!snapshot?.isTooltipActive || !snapshot.activeCoordinate || !snapshot.viewBox) {
    return undefined
  }

  const pointerX = Number(snapshot.activeCoordinate.x ?? Number.NaN)
  const pointerY = Number(snapshot.activeCoordinate.y ?? Number.NaN)
  const boxX = Number(snapshot.viewBox.x ?? 0)
  const boxY = Number(snapshot.viewBox.y ?? 0)
  const boxWidth = Number(snapshot.viewBox.width ?? Number.NaN)
  const boxHeight = Number(snapshot.viewBox.height ?? Number.NaN)

  if (!Number.isFinite(pointerX) || !Number.isFinite(pointerY) || !Number.isFinite(boxWidth) || !Number.isFinite(boxHeight)) {
    return undefined
  }

  const width = Math.max(0, sizing.width)
  const height = Math.max(0, sizing.height)
  const offset = sizing.offset ?? 18
  const padding = sizing.padding ?? 8
  const minX = boxX + padding
  const maxX = boxX + boxWidth - width - padding
  const preferredLeftX = pointerX - width - offset
  const fallbackRightX = pointerX + offset
  const tooltipX = preferredLeftX >= minX ? preferredLeftX : clamp(fallbackRightX, minX, maxX)
  const minY = boxY + padding
  const maxY = boxY + boxHeight - height - padding
  const tooltipY = clamp(pointerY - height / 2, minY, maxY)

  return { x: tooltipX, y: tooltipY }
}
