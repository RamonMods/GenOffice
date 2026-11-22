/**
 * Legacy cloud-backend compatibility layer.
 *
 * The legacy external CLI cloud backend has been removed. Functions
 * that depended on it are kept as graceful no-ops/unavailable so callers and
 * imports do not break. BYOK and keyless search/media providers are handled
 * elsewhere in this package.
 */

import { unwatchFile, watchFile } from 'node:fs'
import {
  asRecord,
  firstItem,
  ramonmodsProxyUrl,
  isCopyrightHost,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'
import {
  genofficeApiKey,
  genofficeAuthPath,
  recoverFromRejectedGskKey,
  reloadGenofficeAuth,
  type GskRecoveryOutcome,
} from './genoffice-auth'

const SEARCH_TIMEOUT_MS = 60_000
const GENERATE_TIMEOUT_MS = 600_000
export const GSK_BACKEND_UNAVAILABLE_ERROR =
  'O backend em nuvem legado não está disponível nesta versão; configure um provedor BYOK ou use a busca integrada'

// ── CLI resolution & auth ───────────────────────────────────────────

/** Legacy CLI entry resolution. Always unavailable (CLI backend removed). */
export function resolveGskEntry(): string | null {
  return null
}

// ── Rejected keys (401) ─────────────────────────────────────────────

/**
 * Keys the server answered 401 for. Kept for API compatibility; no CLI backend
 * is active in this build.
 */
const rejectedKeys = new Set<string>()

/** Thrown when a legacy cloud-key call is refused/unavailable. */
export class GskAuthError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'GskAuthError'
  }
}

/** Whether an error text reports an HTTP 401. */
export function isGskUnauthorizedText(text: unknown): boolean {
  return /\bHTTP 401\b|\b401 Unauthorized\b/i.test(String(text ?? ''))
}

/**
 * Records that `key` was refused with 401. Kept for API compatibility.
 */
export function noteGskKeyRejected(
  key: string,
  options: { allowLogin?: boolean } = {},
): Promise<GskRecoveryOutcome> {
  if (key) rejectedKeys.add(key)
  return recoverFromRejectedGskKey(key, options)
}

/**
 * For callers that used the legacy cloud key outside this module: when `error`
 * reports a 401, drop the key and start recovery in background.
 */
export function noteGskCallFailure(key: string, error: unknown): boolean {
  if (!key) return false
  const text = error instanceof Error ? error.message : String(error ?? '')
  if (!isGskUnauthorizedText(text)) return false
  void noteGskKeyRejected(key).catch(() => {})
  return true
}

/** Test hook: forget rejected keys. */
export function resetGskRejectedKeys(): void {
  rejectedKeys.clear()
}

/**
 * API key for legacy cloud auth; '' when not logged in / unavailable.
 * Kept for API compatibility; the external CLI backend is removed.
 */
export function ramonmodsApiKey(): string {
  const env = process.env.RAMONMODS_API_KEY
  if (env && !rejectedKeys.has(env)) return env
  const own = genofficeApiKey()
  return own && !rejectedKeys.has(own) ? own : ''
}

/**
 * Fires when the effective legacy cloud key changes on disk. Kept for API
 * compatibility; no CLI backend is active.
 */
export function watchGskApiKey(onChange: (key: string) => void, intervalMs = 2000): () => void {
  let last = ramonmodsApiKey()
  const check = (): void => {
    reloadGenofficeAuth()
    const key = ramonmodsApiKey()
    if (key === last) return
    last = key
    onChange(key)
  }
  const files = [genofficeAuthPath()]
  for (const f of files) watchFile(f, { persistent: false, interval: intervalMs }, check)
  return () => {
    for (const f of files) unwatchFile(f, check)
  }
}

/**
 * Whether the legacy cloud backend is usable. Always false: the external CLI
 * backend was removed. Callers should fall back to BYOK/DDG/Parallel.
 */
export function hasGskAuth(): boolean {
  return false
}

// ── Child-process proxy plumbing ────────────────────────────────────

export { setGskProxyUrl, ramonmodsProxyUrl } from './shared'

/**
 * env for legacy CLI children. Kept for API compatibility.
 */
export function ramonmodsChildEnv(base: NodeJS.ProcessEnv = process.env): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = { ...base, ELECTRON_RUN_AS_NODE: '1' }
  const proxy = [
    ramonmodsProxyUrl(),
    base.HTTPS_PROXY,
    base.https_proxy,
    base.HTTP_PROXY,
    base.http_proxy,
    base.ALL_PROXY,
    base.all_proxy,
  ].find((v) => v && /^https?:\/\//.test(v))
  if (proxy) {
    delete env.https_proxy
    delete env.http_proxy
    delete env.ALL_PROXY
    delete env.all_proxy
    env.NODE_USE_ENV_PROXY = '1'
    env.HTTPS_PROXY = proxy
    env.HTTP_PROXY = proxy
  }
  return env
}

// ── Low-level execution ─────────────────────────────────────────────

/**
 * Distill a backend failure to one readable line. Exported for tests and kept
 * for compatibility with callers that format legacy errors.
 */
export function summarizeGskFailure(raw: unknown, fallback = 'unknown error'): string {
  const text = (typeof raw === 'string' ? raw : raw == null ? '' : String(raw)).trim()
  if (!text) return fallback
  const isHtml = /<!doctype|<html[\s>]|<\/html>|<body[\s>]/i.test(text)
  if (!isHtml) {
    const lines = text
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
    const errors = lines.filter((l) => /^\[ERROR\]/i.test(l))
    const kept = errors.length ? errors : lines.filter((l) => !/^\[INFO\]/i.test(l))
    const joined = (kept.length ? kept : lines).join(' ')
    return joined.length > 300 ? `${joined.slice(0, 300)}…` : joined
  }
  const plain = text
    .replace(/<head[\s\S]*?<\/head>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<form[\s\S]*?<\/form>/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&(?:[a-z][a-z0-9]*|#\d+|#x[0-9a-f]+);/gi, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^HTTP\s+\d{3}\s*:?\s*/i, '')
    .trim()
  const head = /HTTP\s+(\d{3})/.exec(text)?.[1]
  const label = head ? `HTTP ${head} (HTML error page)` : 'an HTML error page'
  if (!plain) return label
  return `${label}: ${plain.length > 200 ? `${plain.slice(0, 200)}…` : plain}`
}

interface ScanBudget {
  chars: number
}

const MAX_RECOVERY_SCAN_CHARS = 4 * 1024 * 1024

function jsonBlockAt(text: string, start: number, budget: ScanBudget): string | null {
  let depth = 0
  let inString = false
  let escaped = false
  for (let i = start; i < text.length; i++) {
    if (budget.chars <= 0) return null
    budget.chars--
    const c = text[i]!
    if (inString) {
      if (escaped) escaped = false
      else if (c === '\\') escaped = true
      else if (c === '"') inString = false
      continue
    }
    if (c === '"') inString = true
    else if (c === '{' || c === '[') depth++
    else if (c === '}' || c === ']') {
      depth--
      if (depth === 0) return text.slice(start, i + 1)
    }
  }
  return null
}

/**
 * Parse JSON output from a legacy CLI response. Kept for compatibility/tests.
 */
export function parseGskOutput(stdout: string): unknown {
  const trimmed = stdout.trim()
  try {
    return JSON.parse(trimmed)
  } catch {
    /* fall through to the recovery scan */
  }
  let offset = 0
  const budget: ScanBudget = { chars: MAX_RECOVERY_SCAN_CHARS }
  for (const line of trimmed.split('\n')) {
    if (budget.chars <= 0) break
    const opener = line.trimStart()[0]
    if (opener === '{' || opener === '[') {
      const block = jsonBlockAt(trimmed, offset + line.indexOf(opener), budget)
      if (block) {
        try {
          return JSON.parse(block)
        } catch {
          /* a log line that only looks like JSON; try the next opener */
        }
      }
    }
    offset += line.length + 1
  }
  throw new Error(`No JSON found in output: ${summarizeGskFailure(stdout, '')}`)
}

function unavailable(): never {
  throw new GskAuthError(GSK_BACKEND_UNAVAILABLE_ERROR)
}

// ── Search ──────────────────────────────────────────────────────────

export const MAX_RAMONMODS_RESULTS = 20
export const MAX_RAMONMODS_SNIPPET_CHARS = 2_000
const MAX_RAMONMODS_URL_CHARS = 8_192

function normalizeMaxResults(n: number): number {
  if (!Number.isFinite(n)) return 6
  return Math.min(20, Math.max(1, Math.floor(n)))
}

function clipField(v: unknown, max = MAX_RAMONMODS_SNIPPET_CHARS): string {
  const s = String(v ?? '')
  return s.length > max ? s.slice(0, max) : s
}

/** Parses legacy web search response shape (exported for tests/compat). */
export function parseGskWebSearch(
  raw: unknown,
  maxResults: number,
): { results: WebSearchResult[]; answer?: string } {
  const bounded = normalizeMaxResults(maxResults)
  const data = asRecord(asRecord(raw).data ?? raw)
  const organic: unknown[] = Array.isArray(data.organic_results) ? data.organic_results : []
  const results: WebSearchResult[] = organic.slice(0, bounded).map((item) => {
    const o = asRecord(item)
    return {
      title: clipField(o.title),
      url: clipField(o.link, MAX_RAMONMODS_URL_CHARS),
      snippet: clipField(o.snippet),
    }
  })
  const answerRaw = typeof data.answer === 'string' && data.answer ? data.answer : undefined
  const answer =
    answerRaw !== undefined && answerRaw.length > MAX_RAMONMODS_SNIPPET_CHARS
      ? `${answerRaw.slice(0, MAX_RAMONMODS_SNIPPET_CHARS)}…`
      : answerRaw
  return answer !== undefined ? { results, answer } : { results }
}

export async function ramonmodsWebSearch(
  _query: string,
  _maxResults = 6,
): Promise<{ results: WebSearchResult[]; answer?: string }> {
  unavailable()
}

/** Parses legacy image search response shape (exported for tests/compat). */
export function parseGskImageSearch(raw: unknown, maxResults: number): ImageSearchResult[] {
  const bounded = normalizeMaxResults(maxResults)
  const dataRaw = asRecord(raw).data
  const data: unknown[] = Array.isArray(dataRaw) ? dataRaw : []
  const images: ImageSearchResult[] = []
  for (const item of data) {
    const img = asRecord(item)
    const imageUrl = String(img.image_url ?? img.imageUrl ?? '')
    if (!imageUrl) continue
    if (isCopyrightHost(imageUrl)) continue
    const width = Number(img.width)
    const height = Number(img.height)
    const entry: ImageSearchResult = {
      title: String(img.title ?? ''),
      imageUrl,
      sourceUrl: String(img.link ?? ''),
      source: String(img.source ?? safeHost(img.link)),
    }
    if (Number.isFinite(width) && width > 0) entry.width = width
    if (Number.isFinite(height) && height > 0) entry.height = height
    images.push(entry)
    if (images.length >= bounded) break
  }
  return images
}

export async function ramonmodsImageSearch(
  _query: string,
  _maxResults = 8,
): Promise<ImageSearchResult[]> {
  unavailable()
}

// ── Image generation ────────────────────────────────────────────────

export interface GskGenerateImageOptions {
  prompt: string
  model?: string
  referenceImageUrls?: string[]
  aspectRatio?: string
  imageSize?: string
}

export interface GskGeneratedImage {
  url: string
  taskId: string
}

/** Parses legacy generated image response (exported for tests/compat). */
export function parseGskGeneratedImage(raw: unknown): GskGeneratedImage {
  const data = asRecord(asRecord(raw).data ?? raw)
  const img = asRecord(firstItem(data.generated_images ?? data.images ?? []))
  const url = firstItem(img.image_urls_nowatermark) ?? firstItem(img.image_urls) ?? img.url ?? ''
  if (!url) {
    throw new Error(`image generation returned no result: ${JSON.stringify(raw).slice(0, 200)}`)
  }
  return { url: String(url), taskId: String(img.task_id ?? '') }
}

export async function ramonmodsGenerateImage(
  _options: GskGenerateImageOptions,
  _signal?: AbortSignal,
): Promise<GskGeneratedImage> {
  unavailable()
}

export async function ramonmodsResolveDownloadUrl(url: string): Promise<string> {
  return url
}

// ── Media analysis / transcription ──────────────────────────────────

/** Best-effort text extraction from legacy analysis output (exported for tests/compat). */
export function extractGskText(raw: unknown): string {
  const data = asRecord(raw).data ?? raw
  if (typeof data === 'string') return data
  if (data && typeof data === 'object') {
    for (const key of [
      'analysis',
      'analysis_result',
      'result',
      'content',
      'text',
      'answer',
      'transcript',
    ]) {
      const v = (data as Record<string, unknown>)[key]
      if (typeof v === 'string' && v.trim()) return v
    }
  }
  return JSON.stringify(data ?? raw)
}

export interface GskAnalyzeMediaOptions {
  mediaUrls: string[]
  requirements: string
}

export async function ramonmodsAnalyzeMedia(
  _options: GskAnalyzeMediaOptions,
  _signal?: AbortSignal,
): Promise<string> {
  unavailable()
}

export interface GskTranscribeOptions {
  audioUrls: string[]
  prompt?: string
  model?: string
}

export async function ramonmodsTranscribe(
  _options: GskTranscribeOptions,
  _signal?: AbortSignal,
): Promise<string> {
  unavailable()
}

// ── Past projects (legacy) ──────────────────────────────────────────

export interface GskPastProject {
  projectId: string
  type: string
  title: string
  ctime: string
  projectUrl: string
}

export interface GskPastProjectsPage {
  projects: GskPastProject[]
  total: number
  hasMore: boolean
}

export function parseGskPastProjects(raw: unknown): GskPastProjectsPage {
  const rec = asRecord(raw)
  const data = asRecord(rec.data ?? raw)
  const urlById = new Map<string, string>()
  const sessionProjects = asRecord(asRecord(rec.session_state).past_projects).projects
  if (Array.isArray(sessionProjects)) {
    for (const item of sessionProjects) {
      const p = asRecord(item)
      if (p.project_id && typeof p.project_url === 'string' && p.project_url) {
        urlById.set(String(p.project_id), p.project_url)
      }
    }
  }
  const listRaw: unknown[] = Array.isArray(data.projects) ? data.projects : []
  const projects: GskPastProject[] = []
  for (const item of listRaw) {
    const p = asRecord(item)
    const projectId = String(p.project_id ?? '')
    if (!projectId) continue
    projects.push({
      projectId,
      type: String(p.type ?? ''),
      title: String(p.title ?? ''),
      ctime: String(p.ctime ?? ''),
      projectUrl: urlById.get(projectId) ?? `/agents?id=${projectId}`,
    })
  }
  const total = Number(data.total)
  return {
    projects,
    total: Number.isFinite(total) ? total : projects.length,
    hasMore: data.has_more === true,
  }
}

export interface GskListPastProjectsOptions {
  artifactTypes?: string[]
  limit?: number
  offset?: number
  signal?: AbortSignal
}

export async function ramonmodsListPastProjects(
  _options: GskListPastProjectsOptions = {},
): Promise<GskPastProjectsPage> {
  unavailable()
}

export interface GskLoginInfo {
  email: string
  plan: string
  creditBalance?: number
}

export async function ramonmodsLoginInfo(): Promise<GskLoginInfo | null> {
  return null
}
