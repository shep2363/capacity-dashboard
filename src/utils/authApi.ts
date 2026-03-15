import { buildSharedApiUrl, withAuth } from './apiBase'

export type AuthRole = 'admin' | 'user'

export interface AuthSessionPayload {
  authenticated: boolean
  role: AuthRole | null
}

export class AuthApiError extends Error {
  status: number

  constructor(status: number, message: string) {
    super(message)
    this.status = status
  }
}

async function extractErrorMessage(response: Response): Promise<string> {
  try {
    const payload = await response.json()
    if (payload && typeof payload.error === 'string') {
      return payload.error
    }
    if (payload && typeof payload.message === 'string') {
      return payload.message
    }
  } catch {
    // no-op
  }
  return `Request failed (${response.status}).`
}

async function parseJson<T>(response: Response): Promise<T> {
  if (!response.ok) {
    throw new AuthApiError(response.status, await extractErrorMessage(response))
  }
  return (await response.json()) as T
}

export async function fetchAuthSession(): Promise<AuthSessionPayload> {
  const response = await fetch(buildSharedApiUrl('/api/auth/session'), withAuth())
  return parseJson<AuthSessionPayload>(response)
}

export async function loginWithPassword(password: string): Promise<AuthSessionPayload> {
  const response = await fetch(
    buildSharedApiUrl('/api/auth/login'),
    withAuth({
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password }),
    }),
  )
  return parseJson<AuthSessionPayload>(response)
}

export async function logoutSession(): Promise<void> {
  const response = await fetch(
    buildSharedApiUrl('/api/auth/logout'),
    withAuth({
      method: 'POST',
    }),
  )
  if (!response.ok) {
    throw new AuthApiError(response.status, await extractErrorMessage(response))
  }
}
