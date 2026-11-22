/**
 * The deck outline an agent writes between the style sheet and the page specs:
 * the same shape the app's planning stage produces (core hook + one entry per
 * page). The check mirrors the rules that stage is prompted with, so a CLI
 * agent gets them as findings instead of relying on having read them.
 */
import { MAX_DECK_PAGES } from './deck-spec'

export const PAGE_TYPES = ['cover', 'content', 'data', 'closing'] as const
export type PageType = (typeof PAGE_TYPES)[number]

export const LAYOUT_VARIANTS: Readonly<Record<PageType, readonly string[]>> = {
  cover: [
    'cover_typography_hero',
    'cover_dark_minimal',
    'cover_split_color',
    'cover_full_image_overlay',
    'cover_magazine',
    'cover_split_image',
  ],
  content: [
    'left_text_right_image',
    'three_column_cards',
    'hero_big_number',
    'two_column_comparison',
    'timeline_horizontal',
    'full_image_text_overlay',
  ],
  data: ['kpi_cards_row', 'chart_with_insight', 'two_by_two_grid'],
  closing: ['closing_cta', 'closing_thank_you'],
}

export interface OutlinePage {
  title: string
  type: PageType
  layout: string
  brief: string
  image_queries: string[]
}

export interface DeckOutline {
  topic?: string
  core_hook: string
  pages: OutlinePage[]
}

export interface OutlineIssue {
  /** 0-based page index; absent for deck-level findings */
  page?: number
  level: 'error' | 'warning'
  message: string
}

export const PLACEHOLDER = /\bXX+(?:%|\b)|\blorem\b|\bTBD\b|\bTODO\b|\{\{|\[insert\b|\bN\/A\b/i
const HAS_CJK = new RegExp('[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]')
const MIN_BRIEF = 40
/** Raw LLM outline budget; larger payloads are rejected before JSON.parse. */
const MAX_OUTLINE_RAW_CHARS = 512_000
/** Per-page image query budget; extras are dropped with a warning. */
const MAX_IMAGE_QUERIES_PER_PAGE = 8

export function parseOutline(
  raw: string,
): { ok: true; outline: DeckOutline; issues: OutlineIssue[] } | { ok: false; error: string } {
  // LLM output is untrusted: refuse a megabyte dump before JSON.parse builds
  // a giant object graph from it.
  if (typeof raw !== 'string' || raw.length > MAX_OUTLINE_RAW_CHARS) {
    return { ok: false, error: `esboço muito grande (limite de ${MAX_OUTLINE_RAW_CHARS} caracteres)` }
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(raw)
  } catch (e) {
    return { ok: false, error: `JSON inválido: ${e instanceof Error ? e.message : String(e)}` }
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    return {
      ok: false,
      error:
        'esperado { "core_hook", "pages": [ { "title", "type", "layout", "brief", "image_queries" } ] }',
    }
  }
  const root = parsed as Record<string, unknown>
  if (!Array.isArray(root.pages) || root.pages.length === 0) {
    return { ok: false, error: 'o array "pages" está ausente ou vazio' }
  }
  if (root.pages.length > MAX_DECK_PAGES) {
    return {
      ok: false,
      error: `muitas páginas (${root.pages.length}; o limite é ${MAX_DECK_PAGES})`,
    }
  }
  const issues: OutlineIssue[] = []
  const pages: OutlinePage[] = root.pages.map((p, i) => normalizePage(p, i, issues))
  const coreHook = typeof root.core_hook === 'string' ? root.core_hook.trim() : ''
  if (!coreHook)
    issues.push({
      level: 'error',
      message:
        '"core_hook" está ausente: uma frase com tensão, um número ou um contraste contraintuitivo',
    })
  const topic = typeof root.topic === 'string' ? root.topic.trim() : undefined
  checkDeck(pages, issues)
  return { ok: true, outline: { ...(topic ? { topic } : {}), core_hook: coreHook, pages }, issues }
}

function normalizePage(raw: unknown, i: number, issues: OutlineIssue[]): OutlinePage {
  const p = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const title = typeof p.title === 'string' ? p.title.trim() : ''
  const type = typeof p.type === 'string' ? (p.type.trim() as PageType) : ('' as PageType)
  const layout = typeof p.layout === 'string' ? p.layout.trim() : ''
  const brief = typeof p.brief === 'string' ? p.brief.trim() : ''
  const queries = Array.isArray(p.image_queries) ? p.image_queries : []
  const err = (message: string) => issues.push({ page: i, level: 'error', message })
  const warn = (message: string) => issues.push({ page: i, level: 'warning', message })

  if (!title) err('"title" está ausente')
  if (!(PAGE_TYPES as readonly string[]).includes(type)) {
    err(`"type" deve ser um de ${PAGE_TYPES.join(', ')}`)
  } else if (!layout) {
    err(`"layout" está ausente; escolha um de ${LAYOUT_VARIANTS[type].join(', ')}`)
  } else if (!LAYOUT_VARIANTS[type].includes(layout)) {
    err(
      `"layout" ${layout} não é uma variante de ${type}; escolha um de ${LAYOUT_VARIANTS[type].join(', ')}`,
    )
  }
  if (!brief) err('"brief" está ausente: o que vai em cada região, com os fatos e números reais')
  else if (brief.length < MIN_BRIEF)
    warn(`"brief" é muito curto (${brief.length} caracteres); nomeie o conteúdo de cada região`)
  const placeholder = PLACEHOLDER.exec(brief) ?? PLACEHOLDER.exec(title)
  if (placeholder)
    err(
      `placeholder "${placeholder[0]}" no texto; todo número e nome vem do material`,
    )
  if (!Array.isArray(p.image_queries))
    warn('"image_queries" está ausente; use [] para uma página sem fotos')
  const image_queries: string[] = []
  for (const q of queries) {
    if (image_queries.length >= MAX_IMAGE_QUERIES_PER_PAGE) {
      warn(`"image_queries" limitado a ${MAX_IMAGE_QUERIES_PER_PAGE} por página; extras descartados`)
      break
    }
    if (typeof q !== 'string' || !q.trim()) {
      err(
        '"image_queries" devem ser strings não vazias (uma consulta de cena em inglês ou uma URL http(s))',
      )
      continue
    }
    const s = q.trim()
    if (!/^https?:\/\//i.test(s) && HAS_CJK.test(s)) {
      warn(`a consulta de imagem "${s}" deve ser uma frase em inglês que nomeie uma cena concreta`)
    }
    image_queries.push(s)
  }
  return { title, type, layout, brief, image_queries }
}

function checkDeck(pages: OutlinePage[], issues: OutlineIssue[]): void {
  if (pages[0] && pages[0].type !== 'cover') {
    issues.push({ page: 0, level: 'warning', message: 'a primeira página não é uma capa' })
  }
  const last = pages.length - 1
  if (pages.length > 2 && pages[last] && pages[last].type !== 'closing') {
    issues.push({ page: last, level: 'warning', message: 'a última página não é uma página de encerramento' })
  }
  const body = pages
    .map((p, i) => ({ p, i }))
    .filter(({ p }) => p.type === 'content' || p.type === 'data')
  for (let k = 1; k < body.length; k++) {
    const prev = body[k - 1]!
    const cur = body[k]!
    if (cur.p.layout && cur.p.layout === prev.p.layout && cur.i === prev.i + 1) {
      issues.push({
        page: cur.i,
        level: 'error',
        message: `o layout ${cur.p.layout} se repete em relação à página anterior; varie a composição`,
      })
    }
  }
  const distinct = new Set(body.map(({ p }) => p.layout).filter(Boolean))
  if (body.length >= 4 && distinct.size < 3) {
    issues.push({
      level: 'warning',
      message: `apenas ${distinct.size} variante(s) de layout em ${body.length} páginas de conteúdo; use pelo menos três`,
    })
  }
}
