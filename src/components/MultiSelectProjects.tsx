import { useEffect, useMemo, useRef, useState } from 'react'

interface MultiSelectProjectsProps {
  options: string[]
  selectedValues: string[]
  onChange: (nextSelected: string[]) => void
  placeholder?: string
  entityPlural?: string
  searchPlaceholder?: string
  noMatchingText?: string
  ariaLabel?: string
}

function buildTriggerLabel(options: string[], selected: string[], entityPlural: string): string {
  if (options.length === 0) {
    return `No ${entityPlural}`
  }

  if (selected.length === 0) {
    return `No ${entityPlural} Selected`
  }

  if (selected.length === options.length) {
    return `All ${entityPlural}`
  }

  if (selected.length <= 2) {
    return selected.join(', ')
  }

  return `${selected.length} selected`
}

export function MultiSelectProjects({
  options,
  selectedValues,
  onChange,
  placeholder = 'Projects',
  entityPlural = 'Projects',
  searchPlaceholder = 'Search projects...',
  noMatchingText = 'No matching projects',
  ariaLabel = 'Projects',
}: MultiSelectProjectsProps) {
  const [isOpen, setIsOpen] = useState(false)
  const [search, setSearch] = useState('')
  const rootRef = useRef<HTMLDivElement>(null)
  const searchRef = useRef<HTMLInputElement>(null)

  const selectedSet = useMemo(() => new Set(selectedValues), [selectedValues])
  const filteredOptions = useMemo(
    () => options.filter((option) => option.toLowerCase().includes(search.toLowerCase())),
    [options, search],
  )

  useEffect(() => {
    if (!isOpen) {
      return
    }

    const timer = window.setTimeout(() => {
      searchRef.current?.focus()
    }, 0)

    function onDocumentMouseDown(event: MouseEvent): void {
      const target = event.target as Node
      if (rootRef.current && !rootRef.current.contains(target)) {
        setIsOpen(false)
      }
    }

    function onEscape(event: KeyboardEvent): void {
      if (event.key === 'Escape') {
        setIsOpen(false)
      }
    }

    document.addEventListener('mousedown', onDocumentMouseDown)
    document.addEventListener('keydown', onEscape)

    return () => {
      window.clearTimeout(timer)
      document.removeEventListener('mousedown', onDocumentMouseDown)
      document.removeEventListener('keydown', onEscape)
    }
  }, [isOpen])

  function toggleOption(value: string): void {
    const next = new Set(selectedSet)
    if (next.has(value)) {
      next.delete(value)
    } else {
      next.add(value)
    }

    onChange(options.filter((option) => next.has(option)))
  }

  function selectAll(): void {
    onChange([...options])
  }

  function clearAll(): void {
    onChange([])
  }

  const triggerLabel = buildTriggerLabel(options, selectedValues, entityPlural)

  return (
    <div className="multi-select" ref={rootRef}>
      <span className="multi-label">{placeholder}</span>
      <button
        type="button"
        className="multi-trigger"
        onClick={() => setIsOpen((current) => !current)}
        aria-haspopup="listbox"
        aria-expanded={isOpen}
      >
        <span className="multi-trigger-text">{triggerLabel}</span>
        <span className="multi-caret" aria-hidden="true">
          {isOpen ? '▴' : '▾'}
        </span>
      </button>

      {isOpen && (
        <div className="multi-menu" role="listbox" aria-label={ariaLabel}>
          <div className="multi-actions">
            <button type="button" className="ghost-btn" onClick={selectAll}>
              Select All
            </button>
            <button type="button" className="ghost-btn" onClick={clearAll}>
              Clear All
            </button>
          </div>

          <input
            ref={searchRef}
            type="text"
            className="multi-search"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder={searchPlaceholder}
            aria-label={searchPlaceholder}
          />

          <div className="multi-options">
            {filteredOptions.map((value) => (
              <label key={value} className="multi-option">
                <input type="checkbox" checked={selectedSet.has(value)} onChange={() => toggleOption(value)} />
                <span>{value}</span>
              </label>
            ))}

            {filteredOptions.length === 0 && <div className="multi-empty">{noMatchingText}</div>}
          </div>
        </div>
      )}
    </div>
  )
}
