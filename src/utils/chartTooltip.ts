export interface TooltipPosition {
  x: number
  y: number
}

interface ChartPointerState {
  isTooltipActive?: boolean
  activeCoordinate?: { x?: number; y?: number }
  offset?: { left?: number; top?: number; width?: number; height?: number }
}

interface TooltipBounds {
  width: number
  height: number
  gap?: number
  edgePadding?: number
}

export function computeLeftTooltipPosition(
  state: ChartPointerState | null | undefined,
  bounds: TooltipBounds,
): TooltipPosition | undefined {
  const coordinate = state?.activeCoordinate
  if (!state?.isTooltipActive || typeof coordinate?.x !== 'number' || typeof coordinate?.y !== 'number') {
    return undefined
  }

  const gap = bounds.gap ?? 18
  const edgePadding = bounds.edgePadding ?? 8
  const offsetLeft = typeof state.offset?.left === 'number' ? state.offset.left : edgePadding
  const offsetTop = typeof state.offset?.top === 'number' ? state.offset.top : edgePadding
  const chartWidth =
    typeof state.offset?.width === 'number' && state.offset.width > 0
      ? state.offset.width
      : bounds.width + gap + edgePadding * 2
  const chartHeight =
    typeof state.offset?.height === 'number' && state.offset.height > 0
      ? state.offset.height
      : bounds.height + edgePadding * 2

  const leftCandidate = coordinate.x - bounds.width - gap
  const rightCandidate = coordinate.x + gap
  const minX = offsetLeft + edgePadding
  const maxX = offsetLeft + Math.max(chartWidth - bounds.width - edgePadding, edgePadding)
  const minY = offsetTop + edgePadding
  const maxY = offsetTop + Math.max(chartHeight - bounds.height - edgePadding, edgePadding)
  const clampedLeftX = Math.min(leftCandidate, maxX)

  const resolvedX =
    clampedLeftX >= minX ? clampedLeftX : Math.min(Math.max(rightCandidate, minX), maxX)
  const resolvedY = Math.max(minY, Math.min(coordinate.y - bounds.height / 2, maxY))

  return {
    x: Math.round(resolvedX),
    y: Math.round(resolvedY),
  }
}
