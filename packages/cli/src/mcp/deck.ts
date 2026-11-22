import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  writeFileSync,
} from 'node:fs'
import { isAbsolute, join, resolve } from 'node:path'
import { z } from 'zod'
import { assertAllowed } from '../fs'
import { CliError, EXIT, toJsonError, type JsonError } from '../result'
import { runJson, type McpContext, type Outcome } from './run'
import type { ZodShape } from './tools'

/**
 * The staged deck flow as tools. The deck directory is the state: style.md,
 * outline.json and pages/NN.json are exactly what `genoffice slides check` and
 * `genoffice create --spec <dir>` read, so a deck started here can be finished
 * from the command line and vice versa. Each tool refuses to run before the
 * stage it depends on exists, pages are written in outline order, and a page
 * that fails its check never replaces the file on disk.
 */
export interface DeckTool {
  name: string
  description: string
  shape: ZodShape
  readOnly?: boolean
  run(args: Record<string, unknown>, ctx: McpContext): Promise<Outcome>
}

/** Path policy errors from the deck tools come back as the same JSON error a command would print. */
export async function runDeckTool(
  tool: DeckTool,
  args: Record<string, unknown>,
  ctx: McpContext,
): Promise<Outcome> {
  try {
    return await tool.run(args, ctx)
  } catch (err) {
    const error =
      err instanceof CliError
        ? err
        : new CliError(EXIT.conversion, err instanceof Error ? err.message : String(err))
    return { error: toJsonError('deck', error) }
  }
}

const jsonObject = z.record(z.string(), z.unknown())

const DIR = z
  .string()
  .describe(
    'a pasta do deck (criada se não existir); style.md, outline.json, pages/ e deck.pptx ficam aqui',
  )

export const DECK_TOOLS: DeckTool[] = [
  {
    name: 'deck_start',
    description:
      'Etapa 1 de um deck em etapas, para uma apresentação mais longa ou sensível ao design que uma pessoa irá apresentar (um deck curto de até cerca de 5 slides com conteúdo concreto é feito com uma única chamada create_pptx): escreva a folha de estilo e o esquema, valide o esquema e receba os guias de design e de especificação além das capacidades de nuvem configuradas em uma só chamada. Chame uma vez por deck, antes de qualquer deck_page. Corrija todos os erros de esquema relatados (chame novamente com o esquema corrigido) antes de escrever as páginas. O resultado lista as páginas a escrever, uma chamada deck_page para cada.',
    shape: {
      dir: DIR,
      style: z
        .string()
        .describe(
          'style.md: a folha de estilo do deck em Markdown, indicando a paleta com cores #RRGGBB, as fontes, as variantes de layout e uma frase de estilo; as páginas são validadas contra essas cores',
        ),
      outline: jsonObject.describe(
        'o objeto de esquema: { "core_hook", "pages": [{ "title", "type": cover|content|data|closing, "layout", "brief" com os números reais, "image_queries": [] }] } (formato: guide slides spec)',
      ),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      mkdirSync(join(dir, 'pages'), { recursive: true })
      writeFileSync(join(dir, 'style.md'), String(args.style))
      writeFileSync(join(dir, 'outline.json'), JSON.stringify(args.outline, null, 2))
      const check = await runJson(['slides', 'check', join(dir, 'outline.json')], ctx)
      if (check.error) return check
      const [design, spec, caps] = await Promise.all([
        runJson(['guide', 'slides', 'design'], ctx),
        runJson(['guide', 'slides', 'spec'], ctx),
        runJson(['capabilities'], ctx),
      ])
      const pages = outlinePages(dir)
      return {
        ok: {
          ...check.ok,
          detail: {
            ...check.ok.detail,
            deck: dir,
            pages: pages.map((title, index) => ({ index, title, file: pageFile(dir, index) })),
            guides: { design: design.ok?.summary ?? null, spec: spec.ok?.summary ?? null },
            capabilities: caps.ok?.detail ?? caps.error?.message ?? null,
            next: `escreva a página 0 ("${pages[0] ?? ''}") com deck_page; uma página por chamada, em ordem, relendo a paleta de style.md e a entrada da página no esquema a cada vez`,
          },
        },
      }
    },
  },
  {
    name: 'deck_page',
    description:
      'Etapa 2: uma spec de página para a entrada `index` do esquema, validada contra essa entrada e a folha de estilo antes de ser mantida; uma página que falha na validação não é gravada. Exatamente uma página por chamada, na ordem do esquema (a próxima página ausente, ou uma anterior para reescrever). O resultado lista os achados de auditoria, de esquema e de cores fora da paleta: corrija-os (chame novamente com o mesmo índice) até que os três estejam vazios e então avance para a próxima página. Recusa executar antes de deck_start.',
    shape: {
      dir: DIR,
      index: z.number().int().nonnegative().describe('índice da página no esquema (base 0)'),
      page: jsonObject.describe(
        'o objeto de spec de uma página { "title", "type", "layout", "elements": [...] } repetindo o título, o tipo e o layout da entrada do esquema (formato: guide slides spec)',
      ),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      const started = requireOutline(dir)
      if (started) return started
      const pages = outlinePages(dir)
      const index = Number(args.index)
      const range = pageInRange(index, pages.length)
      if (range) return range
      const before = missingPages(dir, pages.length)
      const next = before[0]
      if (next !== undefined && index > next) {
        return fail(
          'invalid_argument',
          `a página ${index} vem depois da página ausente ${next}; as páginas são escritas na ordem do esquema`,
          { missing_pages: before },
          `escreva primeiro a página ${next} ("${pages[next] ?? ''}")`,
        )
      }
      const check = await checkPending(dir, index, args.page, ctx)
      const missing = missingPages(dir, pages.length)
      const progress = {
        deck: dir,
        page: index,
        file: pageFile(dir, index),
        missing_pages: missing,
      }
      if (check.error) {
        return {
          error: {
            ...check.error,
            detail: { ...check.error.detail, ...progress, kept: false },
            suggestion:
              check.error.suggestion ?? 'corrija a página e chame deck_page novamente com o mesmo índice',
          },
        }
      }
      const then = missing.length
        ? `escreva a página ${missing[0]} ("${pages[missing[0]!] ?? ''}") com deck_page`
        : 'todas as páginas têm arquivo: compile com deck_build'
      return { ok: { ...check.ok, detail: { ...check.ok.detail, ...progress, next: then } } }
    },
  },
  {
    name: 'deck_build',
    description:
      'Etapa 3: compilar o .pptx a partir das páginas validadas. Recusa enquanto alguma página do esquema não tiver arquivo ou uma página divergir de sua entrada. Depois veja o resultado com slides_render e slides_audit; corrija uma página com deck_replace. No máximo duas rodadas de correção.',
    shape: {
      dir: DIR,
      out: z
        .string()
        .optional()
        .describe('caminho de saída .pptx (padrão: <dir>/deck.pptx, sobrescrito na recompilação)'),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      const started = requireOutline(dir)
      if (started) return started
      const missing = missingPages(dir, outlinePages(dir).length)
      if (missing.length) {
        return fail(
          'missing_argument',
          `as páginas ${missing.join(', ')} ainda não têm arquivo`,
          { missing_pages: missing },
          'escreva-as com deck_page antes de compilar',
        )
      }
      const out = args.out
        ? assertAllowed(resolvePath(String(args.out), ctx), ctx.env, 'write')
        : join(dir, 'deck.pptx')
      const argv = [
        'create',
        '--type',
        'pptx',
        '--spec',
        join(dir, 'pages'),
        '--outline',
        join(dir, 'outline.json'),
        '--out',
        out,
      ]
      if (!args.out) argv.push('--force')
      const built = await runJson(argv, ctx)
      if (built.error) return built
      return {
        ok: {
          ...built.ok,
          detail: {
            ...built.ok.detail,
            deck: dir,
            next: `slides_render ${out} para ver cada página, slides_audit para geometria; deck_replace recompila uma página a partir de uma spec corrigida`,
          },
        },
      }
    },
  },
  {
    name: 'deck_replace',
    description:
      'Rodada de correção: reescreva a spec de uma página, valide-a contra a entrada do esquema e a folha de estilo e recompile apenas esse slide do deck já construído; os demais slides mantêm seus ids e conteúdo. Quando a validação falha, o arquivo da página anterior e o deck permanecem como estavam.',
    shape: {
      dir: DIR,
      index: z.number().int().nonnegative().describe('página (slide) a recompilar, base 0'),
      page: jsonObject.describe('o objeto de spec de uma página corrigido'),
      pptx: z.string().optional().describe('o deck já compilado (padrão: <dir>/deck.pptx)'),
    },
    async run(args, ctx) {
      const dir = deckDir(args.dir, ctx)
      const started = requireOutline(dir)
      if (started) return started
      const pages = outlinePages(dir)
      const index = Number(args.index)
      const range = pageInRange(index, pages.length)
      if (range) return range
      const pptx = args.pptx
        ? assertAllowed(resolvePath(String(args.pptx), ctx), ctx.env, 'read')
        : join(dir, 'deck.pptx')
      if (!existsSync(pptx)) {
        return fail('file_not_found', `nenhum deck compilado em ${pptx}`, undefined, 'execute deck_build primeiro')
      }
      const check = await checkPending(dir, index, args.page, ctx)
      if (check.error) {
        return { error: { ...check.error, detail: { ...check.error.detail, kept: false } } }
      }
      const replaced = await runJson(
        ['slides', 'replace', pptx, '--slide', String(index), '--spec', pageFile(dir, index)],
        ctx,
      )
      if (replaced.error) return replaced
      return {
        ok: {
          ...replaced.ok,
          detail: {
            ...replaced.ok.detail,
            check: check.ok.detail,
            next: `slides_render ${pptx} com o slide ${index} para confirmar`,
          },
        },
      }
    },
  },
]

/**
 * The page is checked from a `.pending` file beside the real one (so
 * `slides check` finds outline.json and style.md as usual) and only renamed
 * into place when it passes; a failing spec leaves the previous file alone.
 */
async function checkPending(
  dir: string,
  index: number,
  page: unknown,
  ctx: McpContext,
): Promise<Outcome> {
  const file = pageFile(dir, index)
  const pending = `${file}.pending`
  writeFileSync(pending, JSON.stringify(page, null, 2))
  try {
    const check = await runJson(['slides', 'check', pending, '--page', String(index)], ctx)
    if (check.ok) renameSync(pending, file)
    return check
  } finally {
    rmSync(pending, { force: true })
  }
}

function resolvePath(path: string, ctx: McpContext): string {
  return isAbsolute(path) ? path : resolve(ctx.cwd, path)
}

/** The deck folder, inside GENOFFICE_ALLOWED_ROOTS when that is set: everything the tools write lives under it. */
function deckDir(raw: unknown, ctx: McpContext): string {
  return assertAllowed(resolvePath(String(raw), ctx), ctx.env, 'write')
}

function pageFile(dir: string, index: number): string {
  return join(dir, 'pages', `${String(index + 1).padStart(2, '0')}.json`)
}

function pageInRange(index: number, count: number): Outcome | null {
  if (index < count) return null
  return fail('out_of_range', `o esquema tem ${count} páginas; o índice ${index} não tem entrada`, {
    valid_range: [0, count - 1],
  })
}

function outlinePages(dir: string): string[] {
  try {
    const raw = JSON.parse(readFileSync(join(dir, 'outline.json'), 'utf-8')) as {
      pages?: { title?: unknown }[]
    }
    return (raw.pages ?? []).map((p) => (typeof p?.title === 'string' ? p.title : ''))
  } catch {
    return []
  }
}

/** Outline indexes whose page file is missing; extra files beyond the outline are the build's problem. */
function missingPages(dir: string, count: number): number[] {
  const present = new Set(
    existsSync(join(dir, 'pages'))
      ? readdirSync(join(dir, 'pages')).filter((f) => f.endsWith('.json'))
      : [],
  )
  const missing: number[] = []
  for (let i = 0; i < count; i++) {
    if (!present.has(`${String(i + 1).padStart(2, '0')}.json`)) missing.push(i)
  }
  return missing
}

function requireOutline(dir: string): Outcome | null {
  if (existsSync(join(dir, 'outline.json'))) return null
  return fail(
    'missing_argument',
    `${dir} não contém outline.json`,
    undefined,
    'chame deck_start primeiro com a folha de estilo e o esquema',
  )
}

function fail(
  reason: JsonError['error'],
  message: string,
  detail?: Record<string, unknown>,
  suggestion?: string,
): Outcome {
  return {
    error: {
      status: 'error',
      command: 'deck',
      code: 1,
      error: reason,
      message,
      ...(suggestion ? { suggestion } : {}),
      ...(detail ? { detail } : {}),
    },
  }
}
