/**
 * Legacy cloud-auth compatibility stub.
 *
 * The legacy external cloud auth backend was removed. Exports are preserved so
 * callers compile; operations are no-ops or report signed-out/unavailable.
 */

export interface GskLoginProgress {
  phase: 'url' | 'success' | 'error'
  url?: string
  expiresInSec?: number
  error?: string
}

export function isAllowedAuthUrl(_raw: string, _origin: string): URL | null {
  return null
}

import { homedir } from 'node:os'
import { join } from 'node:path'

export function genofficeAuthPath(): string {
  return join(process.env.GENOFFICE_AUTH_DIR || join(homedir(), '.genoffice'), 'auth.json')
}

export interface GenofficeAuth {
  apiKey: string
  keyId?: string
  accessToken?: string
}

export function loadGenofficeAuth(): GenofficeAuth | null {
  return null
}

export function reloadGenofficeAuth(): void {}

export function genofficeApiKey(): string {
  return ''
}

export const FRESH_AUTH_REQUIRED = 'fresh_browser_authentication_required'

export type MintResult =
  | { kind: 'ok'; apiKey: string; keyId?: string }
  | { kind: 'fresh_auth_required' }
  | { kind: 'error'; reason: string }

export function classifyMintResponse(status: number, json: Record<string, unknown>): MintResult {
  if (status === 403 && json.detail === FRESH_AUTH_REQUIRED) return { kind: 'fresh_auth_required' }
  const data =
    json.data && typeof json.data === 'object' ? (json.data as Record<string, unknown>) : {}
  const apiKey = typeof data.token === 'string' ? data.token : ''
  if (status >= 200 && status < 300 && apiKey) {
    return {
      kind: 'ok',
      apiKey,
      ...(typeof data.key_id === 'string' ? { keyId: data.key_id } : {}),
    }
  }
  const message =
    typeof json.message === 'string' && json.message
      ? json.message
      : typeof json.detail === 'string' && json.detail
        ? json.detail
        : status >= 200 && status < 300
          ? 'API key creation failed'
          : `API key creation failed (HTTP ${status})`
  return { kind: 'error', reason: message }
}

export function startGenofficeLogin(_onEvent?: (progress: GskLoginProgress) => void): boolean {
  return false
}

export type GskRecoveryOutcome = 'reminted' | 'login_started' | 'login_required'

export function setGenofficeLoginOpener(_openUrl: ((url: string) => void) | null): void {}

export function setGenofficeLoginSuccessListener(_listener: (() => void) | null): void {}

export function invalidateGenofficeKey(_key: string): boolean {
  return false
}

export function recoverFromRejectedGskKey(
  _key: string,
  _options: { allowLogin?: boolean } = {},
): Promise<GskRecoveryOutcome> {
  return Promise.resolve('login_required')
}

export function genofficeLoginInFlight(): boolean {
  return false
}

export function ensureGenofficeLogin(_openUrl: (url: string) => void): void {}

export async function genofficeLogout(): Promise<void> {}

export function resetGenofficeAuthCache(): void {}

export function genofficeProxyFallbackPreferred(): boolean {
  return false
}
