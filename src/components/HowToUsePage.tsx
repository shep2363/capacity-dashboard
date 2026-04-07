interface HowToUsePageProps {
  isUserMode: boolean
  onOpenReport: () => void
  onOpenProcessing: () => void
  onOpenRevenue: () => void
  onOpenPlanning: () => void
  onLock: () => void
}

const tabGuide = [
  {
    title: 'Report Workspace',
    description:
      'See weekly and monthly forecast trends, project totals, and summary metrics in one place. You can also select one week, toggle weeks with Ctrl + click, or add a full range with Ctrl + Shift + click.',
  },
  {
    title: 'Department Tabs',
    description: 'Review Processing, Fabrication, Assembly, Paint, and Shipping work grouped by week and sequence.',
  },
  {
    title: 'Planning',
    description: 'Use the planning page to review the active workbook, current filter window, and capacity context.',
  },
  {
    title: 'Revenue',
    description: 'Check revenue and gross profit views tied to the currently loaded project data.',
  },
]

const quickStartSteps = [
  'Open Report Workspace first to get the overall forecast picture for the selected year and week range.',
  'When comparing capacity across several weeks, click one week first, then use Ctrl + Shift + click on a later week to select the full range in between.',
  'Move into the department tabs when you want to answer who is working on what and when each sequence is planned.',
  'Use the Planning page when you need to confirm the source workbook, current filters, and capacity assumptions.',
  'Lock the dashboard when you are finished so the next person starts from a clean access screen.',
]

export function HowToUsePage({
  isUserMode,
  onOpenReport,
  onOpenProcessing,
  onOpenRevenue,
  onOpenPlanning,
  onLock,
}: HowToUsePageProps) {
  return (
    <section className="how-to-use-page">
      <div className="panel how-to-use-hero">
        <div className="how-to-use-hero-copy">
          <span className="how-to-use-badge">{isUserMode ? 'User Guide' : 'Dashboard Guide'}</span>
          <div className="section-header">
            <h2>How to Use App</h2>
            <p>
              This page gives users a quick map of the dashboard so they can move from summary views to department
              detail without guessing where to click next.
            </p>
          </div>
          <div className="how-to-use-callout">
            {isUserMode
              ? 'User mode is best for reviewing the shared plan, checking department progress, and exporting information.'
              : 'Admin mode includes the same viewing flow, plus workbook upload and planning edit controls.'}
          </div>
          <div className="how-to-use-actions">
            <button type="button" onClick={onOpenReport}>
              Open Report Workspace
            </button>
            <button type="button" className="ghost-btn" onClick={onOpenProcessing}>
              Open Processing
            </button>
            <button type="button" className="ghost-btn" onClick={onOpenRevenue}>
              Open Revenue
            </button>
            <button type="button" className="ghost-btn" onClick={onOpenPlanning}>
              Open Planning
            </button>
            <button type="button" className="ghost-btn" onClick={onLock}>
              Lock App
            </button>
          </div>
        </div>
      </div>

      <div className="how-to-use-grid">
        <section className="panel how-to-use-card">
          <div className="section-header">
            <h2>Quick Start</h2>
            <p>Follow this order when you are opening the dashboard for a normal review session.</p>
          </div>
          <div className="how-to-use-steps">
            {quickStartSteps.map((step, index) => (
              <div key={step} className="how-to-use-step">
                <span className="how-to-use-step-number">{index + 1}</span>
                <p>{step}</p>
              </div>
            ))}
          </div>
        </section>

        <section className="panel how-to-use-card">
          <div className="section-header">
            <h2>What Each Tab Does</h2>
            <p>Use this as a quick reminder when someone asks where a specific answer lives.</p>
          </div>
          <div className="how-to-use-tab-list">
            {tabGuide.map((item) => (
              <article key={item.title} className="how-to-use-tab-item">
                <strong>{item.title}</strong>
                <p>{item.description}</p>
              </article>
            ))}
          </div>
        </section>
      </div>

      <section className="panel how-to-use-card">
        <div className="section-header">
          <h2>Helpful Tips</h2>
          <p>These habits make the shared dashboard easier to use across the team.</p>
        </div>
        <div className="how-to-use-tip-grid">
          <article className="how-to-use-tip">
            <strong>Start broad, then narrow</strong>
            <p>Use Report Workspace for the big picture first, then jump into the department tabs for sequence detail.</p>
          </article>
          <article className="how-to-use-tip">
            <strong>Check the active files</strong>
            <p>The Planning page shows the active workbook names and sync status, which helps confirm you are reviewing the right data.</p>
          </article>
          <article className="how-to-use-tip">
            <strong>Use week range selection in reports</strong>
            <p>
              In the forecast charts, click a week to anchor your selection. Ctrl + click toggles individual weeks, and
              Ctrl + Shift + click adds every week between the anchor week and the week you clicked.
            </p>
          </article>
          <article className="how-to-use-tip">
            <strong>Refresh department progress when needed</strong>
            <p>Department pages let you refresh Smartsheet progress so completion percentages stay current during a review.</p>
          </article>
          <article className="how-to-use-tip">
            <strong>Know the user-mode limit</strong>
            <p>
              {isUserMode
                ? 'Planning edit sections stay hidden in user mode, so this login is intended for review, reporting, and department follow-up.'
                : 'Admin mode can manage the shared planning data, so be intentional before changing workbook or capacity settings.'}
            </p>
          </article>
        </div>
      </section>
    </section>
  )
}
