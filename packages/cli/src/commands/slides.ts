import { readFileSync } from 'node:fs'
import { basename, dirname } from 'node:path'
import { flagBool, flagString } from '../args'
import {
  applyOps,
  describeDeck,
  PREVIEW_CHARS,
  inlineLocalFiles,
  openDeck,
  parseOps,
  saveDeck,
} from '../formats/pptx'
import { previewChars } from '../preview'
import { readInput, resolveInput, resolveOutput, writeOutput } from '../fs'
import { outputDirectory, parseScale, renderToPngs } from '../formats/render'
import {
  auditDeckBytes,
  checkPageSpec,
  deckAuditDetail,
  findStageFiles,
  pageIndexOf,
  replaceSlideFromSpec,
  stageContext,
} from '../formats/slide-spec'
import { parseOutline } from '@genoffice/pipelines/slides'
import { readOpsInput } from '../ops-input'
import type { TxnResult } from '@genoffice/pptx-ops'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'
import { txnDetail, txnFailure } from './txn'
import { BATCH_OPTIONS, batchCounts, batchMode, batchResult, failedBatch } from '../batch'
import type { OpFailure } from '../op-errors'
import { emuPer, isReadUnit, READ_UNITS, type ReadUnit } from '../length-units'

export const slidesCommand: CommandDef = {
  name: 'slides',
  summary:
    'Lê, edita, audita ou renderiza um .pptx com as ops e a auditoria de layout usadas pela IA do app; verifica um outline ou uma especificação de página; reconstrói um slide a partir da sua especificação.',
  usage:
    'slides <read|apply|audit|render|replace> <file.pptx> [options] | slides check <outline.json|page.json>',
  options: [
    {
      name: 'slide',
      value: 'n',
      description: 'read/audit/render: apenas este slide (índice a partir de 0); replace: o slide a reconstruir',
    },
    {
      name: 'full',
      description: 'read: texto completo e todas as linhas de tabela em vez de prévias, mais as notas do palestrante',
    },
    {
      name: 'layouts',
      description:
        'read: também lista os layouts da apresentação (nome, índice, placeholders) para addSlideWithLayout',
    },
    {
      name: 'max-chars',
      value: 'n',
      description:
        'read: tamanho da prévia por elemento de texto (padrão 300); texto cortado termina em …(+n caracteres)',
    },
    {
      name: 'units',
      value: 'unit',
      description: `read: informa os comprimentos de caixas e do slide em ${READ_UNITS.join(', ')} (padrão emu; px a 96 dpi)`,
    },
    { name: 'ops', value: 'file', description: 'apply: arquivo JSON de ops, ou "-" para stdin' },
    { name: 'spec', value: 'file', description: 'replace: o arquivo de especificação de uma página a construir' },
    {
      name: 'outline',
      value: 'file',
      description:
        'check <page.json> / replace: o outline usado como referência da página (padrão: outline.json ao lado do arquivo da página ou uma pasta acima)',
    },
    { name: 'dry-run', description: 'apply: valida e imprime o plano sem gravar' },
    ...BATCH_OPTIONS,
    {
      name: 'isolation',
      value: 'mode',
      description: 'apply: atomic (padrão) ou per_op (igual a --best-effort)',
    },
    { name: 'out', value: 'path', description: 'apply: grava aqui em vez de no próprio arquivo' },
    {
      name: 'force',
      description:
        'apply: sobrescreve um arquivo --out existente ou grava enquanto o GenOffice tem o arquivo aberto',
    },
    {
      name: 'page',
      value: 'n',
      description:
        "check <page.json>: a entrada do outline (índice a partir de 0) usada como referência da página (padrão: a posição do arquivo entre os arquivos de página da pasta)",
    },
    {
      name: 'scale',
      value: 'n',
      description: 'render: pixels por ponto (padrão 1 = 960x540 para uma apresentação 16:9; 2 dobra isso)',
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    switch (verb) {
      case 'read':
        return read(file, args, ctx)
      case 'apply':
        return apply(file, args, ctx)
      case 'audit':
        return audit(file, args, ctx)
      case 'render':
        return render(file, args, ctx)
      case 'check':
        return check(file, args, ctx)
      case 'replace':
        return replace(file, args, ctx)
      default:
        throw new CliError(
          EXIT.usage,
          'esperado "slides read|apply|audit|render|replace <file.pptx>" ou "slides check <outline.json|page.json>"',
          undefined,
          {
            reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
            suggestion: 'execute `genoffice help slides`',
          },
        )
    }
  },
}

function slideFlag(args: Parameters<CommandDef['run']>[0], count: number): number | undefined {
  const index = slideFlagValue(args)
  if (index === undefined) return undefined
  if (index >= count) {
    throw new CliError(
      EXIT.usage,
      `--slide fora do intervalo (0-${count - 1})`,
      { valid_range: [0, count - 1] },
      { reason: 'out_of_range', suggestion: `use um índice a partir de 0 entre 0 e ${count - 1}` },
    )
  }
  return index
}

/**
 * `--slide` as a 0-based index, before any file is opened: a malformed value
 * (a shell loop that pasted "2 03") is reported as such, not as a missing file
 * or an out-of-range slide.
 */
function slideFlagValue(args: Parameters<CommandDef['run']>[0]): number | undefined {
  const only = flagString(args, 'slide')
  if (only === undefined) return undefined
  const index = Number(only)
  if (!/^\d+$/.test(only.trim()) || !Number.isInteger(index) || index < 0) {
    throw new CliError(
      EXIT.usage,
      `--slide deve ser um índice de slide a partir de 0 (ex.: --slide 2), recebido "${only}"`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  return index
}

async function read(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const unit = unitsFlag(args)
  const path = resolveInput(file, ctx)
  const opened = await openDeck(readInput(path))
  const index = slideFlag(args, opened.deck.slides.length)
  const deck = describeDeck(
    opened,
    index,
    flagBool(args, 'full'),
    previewChars(args, PREVIEW_CHARS),
    flagBool(args, 'layouts'),
    unit,
  )
  return {
    summary: `${basename(path)}: ${deck.slides} slides${deck.layouts ? `, ${deck.layouts.length} layouts` : ''}`,
    detail: {
      ...deck,
      units: `${unitsNote(unit)}; os ids de slide s_<n> e os ids de elemento e_* são alvos persistentes de ops; layouts[].name/index alimentam addSlideWithLayout`,
    },
  }
}

function unitsFlag(args: Parameters<CommandDef['run']>[0]): ReadUnit {
  const value = flagString(args, 'units')
  if (value === undefined) return 'emu'
  if (!isReadUnit(value)) {
    throw new CliError(
      EXIT.usage,
      `--units deve ser um de ${READ_UNITS.join(', ')}, recebido "${value}"`,
      { valid_values: READ_UNITS },
      { reason: 'invalid_argument' },
    )
  }
  return value
}

function unitsNote(unit: ReadUnit): string {
  if (unit === 'emu') return 'EMU (914400 por polegada)'
  const dpi = unit === 'px' ? ' a 96 dpi' : ''
  return `${unit}${dpi} (${emuPer(unit)} EMU cada; caixa e tamanho trazem a unidade: "${unit}"); ops usam os mesmos comprimentos em strings como "1.5${unit}"`
}

async function apply(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const mode = batchMode(args)
  const { text, source } = readOpsInput(args, ctx)
  const ops = inlineLocalFiles(parseOps(text, source), ctx)
  const opened = await openDeck(readInput(path))
  const dryRun = flagBool(args, 'dry-run')
  // One transaction per op, as `create` does: the executor validates a batch
  // against the pre-transaction deck, so a slide added by op 0 would not exist
  // for op 1 in a single transaction. Atomic still means nothing is written
  // unless every op applied; the other modes keep what applied and report the rest.
  const records: NonNullable<TxnResult['records']> = []
  const failures: NonNullable<TxnResult['failures']> = []
  // dry-run plan lines keep the caller's op index, so they line up with failures
  const plan: string[] = []
  let applied = 0
  for (const [index, op] of ops.entries()) {
    // the op's own isolation mode, so a per_op failure is not worded as an atomic rollback
    const r = applyOps(opened, [op], { isolation: mode === 'atomic' ? 'atomic' : 'per_op' })
    if (r.applied) {
      applied++
      records.push(...(r.records ?? []))
      const slide = r.records?.[0]?.slideId
      plan.push(`${index}: ${op.op}${slide ? ` on ${slide}` : ''}`)
      continue
    }
    if (mode === 'atomic') throw failedBatch(txnFailure(r, index), ops.length)
    failures.push(...(r.failures ?? []).map((f) => ({ ...f, index })))
    if (mode === 'stop_on_error') break
  }
  const r: TxnResult = { applied: records.length > 0, records, failures }
  const counts = batchCounts(ops.length, applied, failures.length)
  const finish = (base: CommandResult): CommandResult => {
    const { failures: classified, ...detail } = txnDetail(r)
    return batchResult({ ...base, detail }, counts, (classified as OpFailure[] | undefined) ?? [])
  }
  if (!applied) throw failedBatch(txnFailure(r), ops.length)
  if (dryRun) {
    // the batch ran on the in-memory deck only: nothing is applied to the file
    r.applied = false
    r.dryRun = true
    r.plan = plan
    return finish({ summary: `simulação: ${applied} de ${ops.length} ops validadas` })
  }
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    // in-place edits overwrite by design; --out onto another existing file needs --force
    fresh: flagString(args, 'out') !== undefined,
  })
  writeOutput(output, await saveDeck(opened))
  return finish({
    summary: `aplicadas ${applied} de ${ops.length} ops a ${basename(output)}`,
    outputPath: output,
  })
}

/** Geometry audit per slide: out of bounds, text overflow, overlapping content, with durable ids. */
async function audit(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const bytes = readInput(path)
  const opened = await openDeck(bytes)
  const index = slideFlag(args, opened.deck.slides.length)
  const pages = await auditDeckBytes(bytes, index)
  const { counts, issues, slides } = deckAuditDetail(pages)
  const failing = pages.filter((p) => p.issues.length > 0)
  return {
    summary: failing.length
      ? `${basename(path)}: ${counts.error + counts.warning} problema(s) em ${failing.length} de ${pages.length} slide(s)`
      : `${basename(path)}: ${pages.length} slide(s), sem problemas de layout`,
    detail: {
      issues,
      counts,
      slides,
      metrics: 'larguras de glifo heurísticas; números de overflow são aproximados',
      ids: 'os ids de elemento correspondem aos alvos de `slides read` / `slides apply`; `suggest` é uma op setTransform (EMU) para `slides apply`',
    },
  }
}

/**
 * Findings for one deck-flow artifact: an outline (core hook + page plan) or a
 * single page spec, told apart by their shape. Errors exit 1 so a script can
 * gate on them; warnings are advice.
 */
async function check(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const text = readFileSync(path, 'utf-8')
  const dir = dirname(path)
  if (looksLikeOutline(text)) {
    const r = parseOutline(text)
    if (!r.ok) {
      throw new CliError(EXIT.usage, `${basename(path)}: ${r.error}`, undefined, {
        reason: 'invalid_argument',
      })
    }
    const errors = r.issues.filter((i) => i.level === 'error')
    const style = findStageFiles(dir, ctx).style
    const detail = {
      kind: 'outline',
      pages: r.outline.pages.length,
      issues: r.issues,
      notes: style
        ? []
        : ['sem style.md ao lado do outline: escreva a folha de estilo antes das páginas'],
      next: 'escreva pages/01.json … um arquivo por página do outline, executando `genoffice slides check <page.json>` em cada um',
    }
    if (errors.length) {
      throw new CliError(
        EXIT.usage,
        `${basename(path)}: ${errors.length} erro(s) no outline`,
        detail,
      )
    }
    return {
      summary: `${basename(path)}: ${r.outline.pages.length} páginas, ${r.issues.length ? `${r.issues.length} aviso(s)` : 'sem achados'}`,
      detail,
    }
  }
  const explicit = flagString(args, 'outline')
  const stage = stageContext(dir, ctx, explicit ? resolveInput(explicit, ctx) : undefined)
  const page = pageFlagValue(args) ?? pageIndexOf(path)
  const r = await checkPageSpec(text, path, ctx, { context: stage, page })
  const outline = r.stage?.outline ?? []
  const offPalette = r.stage?.offPalette ?? []
  const count =
    r.warnings.length + r.imageFailures.length + r.audit.length + outline.length + offPalette.length
  const detail = {
    kind: 'page',
    page,
    warnings: r.warnings,
    imageFailures: r.imageFailures,
    audit: r.audit,
    outline: stage.outline ? { file: stage.outline.file, findings: outline } : null,
    style: stage.style
      ? {
          file: stage.style.file,
          offPalette,
          ...(offPalette.length
            ? {
                note: 'cores que a folha de estilo não nomeia; adicione-as ao style.md ou use a paleta',
              }
            : {}),
        }
      : null,
    notes: stage.notes,
    metrics: 'larguras de glifo heurísticas; números de overflow são aproximados',
  }
  const errors = outline.filter((o) => o.level === 'error')
  if (errors.length) {
    throw new CliError(
      EXIT.usage,
      `${basename(path)}: ${errors.length} erro(s) em relação a ${basename(stage.outline!.file)} pages[${page}]`,
      detail,
    )
  }
  return {
    summary: count
      ? `${basename(path)}: ${count} achado(s)`
      : `${basename(path)}: página construída sem problemas`,
    detail,
  }
}

function pageFlagValue(args: Parameters<CommandDef['run']>[0]): number | undefined {
  const raw = flagString(args, 'page')
  if (raw === undefined) return undefined
  if (!/^\d+$/.test(raw.trim())) {
    throw new CliError(
      EXIT.usage,
      `--page deve ser um índice de outline a partir de 0 (ex.: --page 2), recebido "${raw}"`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  return Number(raw)
}

function looksLikeOutline(text: string): boolean {
  try {
    const v = JSON.parse(text) as Record<string, unknown> | null
    if (!v || typeof v !== 'object' || Array.isArray(v) || !Array.isArray(v.pages)) return false
    // a multi-page deck spec also has "pages", but its entries carry elements, not briefs
    return !v.pages.some((p) => p && typeof p === 'object' && 'elements' in (p as object))
  } catch {
    return false
  }
}

/** One page rebuilt from its spec in place of slide --slide; the other slides keep their ids and content. */
async function replace(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  if (slideFlagValue(args) === undefined)
    throw new CliError(EXIT.usage, 'falta --slide <n>', undefined, { reason: 'missing_argument' })
  const path = resolveInput(file, ctx)
  const specFile = flagString(args, 'spec')
  if (!specFile)
    throw new CliError(EXIT.usage, 'falta --spec <page.json>', undefined, {
      reason: 'missing_argument',
    })
  const specPath = resolveInput(specFile, ctx)
  const opened = await openDeck(readInput(path))
  const index = slideFlag(args, opened.deck.slides.length)!
  const explicit = flagString(args, 'outline')
  const stage = stageContext(
    dirname(specPath),
    ctx,
    explicit ? resolveInput(explicit, ctx) : undefined,
  )
  const r = await replaceSlideFromSpec(
    opened,
    index,
    readFileSync(specPath, 'utf-8'),
    specPath,
    ctx,
    stage,
  )
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    fresh: flagString(args, 'out') !== undefined,
  })
  writeOutput(output, await saveDeck(opened))
  return {
    summary: `slide ${index} de ${basename(output)} substituído a partir de ${basename(specPath)}`,
    outputPath: output,
    detail: {
      slide: index,
      issues: r.warnings,
      imageFailures: r.imageFailures,
      outline: stage.outline
        ? { file: stage.outline.file, findings: r.stage?.outline ?? [] }
        : null,
      offPalette: r.stage?.offPalette ?? [],
      notes: stage.notes,
      next: 'os ids de elemento no slide reconstruído são novos: execute `genoffice slides read` antes de usá-los como alvo',
    },
  }
}

/** One PNG per slide, through the app's PDF export (hidden GenOffice process) and pdfium. */
async function render(
  file: string | undefined,
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const outDir = outputDirectory(flagString(args, 'out'), ctx)
  const scale = parseScale(flagString(args, 'scale'))
  const opened = await openDeck(readInput(path))
  const only = slideFlag(args, opened.deck.slides.length)
  const files = await renderToPngs(path, ctx, { outDir, scale, only, log: ctx.log })
  return {
    summary: `renderizados ${files.length} slide(s) em ${outDir}`,
    outputPath: outDir,
    detail: { files: files.map(({ page, ...f }) => ({ slide: page, ...f })) },
  }
}
