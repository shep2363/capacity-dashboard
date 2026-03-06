import { useEffect, useMemo, useRef, useState } from 'react'

interface MultiSelectProjectsProps {
  options: string[]
  selectedValues: string[]
  onChange: (nextSelected: string[]) => void
  placeholder?: string
}

function buildTriggerLabel(options: string[], selected: string[]): string {
  if (options.length === 0) {
    return 'No Projects'
  }

  if (selected.length === 0) {
    return 'No Projects Selected'
  }

  if (selected.length === options.length) {
    return 'All Projects'
  }

  if (selected.length <= 2) {
    return selected.join(', ')
  }

  return `${selected.length} projects selected`
}

export function MultiSelectProjects({
  options,
  selectedValues,
  onChange,
  placeholder = 'Projects',
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

  function toggleOption(project: string): void {
    const next = new Set(selectedSet)
    if (next.has(project)) {
      next.delete(project)
    } else {
      next.add(project)
    }

    onChange(options.filter((option) => next.has(option)))
  }

  function selectAll(): void {
    onChange([...options])
  }

  function clearAll(): void {
    onChange([])
  }

  const triggerLabel = buildTriggerLabel(options, selectedValues)

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
        <div className="multi-menu" role="listbox" aria-label="Projects">
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
            placeholder="Search projects..."
            aria-label="Search projects"
          />

          <div className="multi-options">
            {filteredOptions.map((project) => (
              <label key={project} className="multi-option">
                <input
                  type="checkbox"
                  checked={selectedSet.has(project)}
                  onChange={() => toggleOption(project)}
                />
                <span>{project}</span>
              </label>
            ))}

            {filteredOptions.length === 0 && <div className="multi-empty">No matching projects</div>}
          </div>
        </div>
      )}
    </div>
  )
}
