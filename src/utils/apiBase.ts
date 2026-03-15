function defaultLocalApiBase(): string {
  if (typeof window === 'undefined') {
    return ''
  }
  if (window.location.hostname === 'localhost') {
    return 'http://localhost:8000'
  }
  if (window.location.hostname === '127.0.0.1') {
    return 'http://127.0.0.1:8000'
  }
  return ''
}

export function buildSharedApiUrl(path: string): string {
  const configuredBase = import.meta.env.VITE_SHARED_DATA_API_URL ?? defaultLocalApiBase()
  if (!configuredBase) {
    return path
  }
  return `${configuredBase.replace(/\/$/, '')}${path}`
}

export function buildExportApiUrl(): string {
  const configuredBase = import.meta.env.VITE_EXPORT_API_URL
  if (configuredBase) {
    return configuredBase
  }
  return buildSharedApiUrl('/api/export-report')
}

export function withAuth(init: RequestInit = {}): RequestInit {
  return {
    cache: 'no-store',
    credentials: 'include',
    ...init,
  }
}
