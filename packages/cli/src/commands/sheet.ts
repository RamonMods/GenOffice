import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { flagBool, flagString } from '../args'
import {
  cellEditsFromInputs,
  READ_WHERE,
  readSheet,
  workbookSummary,
  writeWorkbook,
  type CellInput,
  type ReadWhere,
  readExternalLinks,
  updateExternalLinkCaches,
  breakExternalLinks,
} from '../formats/xlsx'
import { checkWorkbook, SHEET_CHECKS } from '../formats/xlsx-check'
import { runWorkbookDsl, SUPPORTED_DSL_OPS, type DslOutcome } from '../formats/xlsx-dsl'
import { hasGatewayPayloads, type GatewayPayloads } from '../formats/xlsx-gateway-ops'
import { readInput, resolveInput, resolveOutput } from '../fs'
import type { CommandContext, CommandDef } from '../registry'
import { checkResult } from '../check'
import { CliError, EXIT, type CommandResult, type Warning } from '../result'
import { didYouMean, sheetNotFoundHints } from '../suggest'
import {
  BATCH_OPTIONS,
  batchCounts,
  batchMode,
  batchResult,
  failedBatch,
  pinnedFailure,
  type BatchMode,
} from '../batch'
import type { OpFailure } from '../op-errors'

const FORMULAS_NOT_CACHED = (message: string): Warning => ({ code: 'formulas_not_cached', message })
const FORMULA_ERRORS = (message: string): Warning => ({ code: 'formula_errors', message })
const DEPENDENTS_NOT_RECALCULATED = (message: string): Warning => ({
  code: 'dependents_not_recalculated',
  message,
})

export const sheetCommand: CommandDef = {
  name: 'sheet',
  summary: 'Lê ou edita um .xlsx pelo mesmo caminho de gravação usado pelo aplicativo.',
  usage: 'sheet <read|apply|links|check> <file.xlsx> (--cells <file|-> | --ops <file|->) [options]',
  options: [
    { name: 'sheet', value: 'name', description: 'planilha (padrão: a ativa)' },
    { name: 'range', value: 'A1:D20', description: 'read: intervalo de células (padrão: até 500×100)' },
    { name: 'cols', value: 'A,C,E:G', description: 'read: apenas estas colunas do intervalo' },
    {
      name: 'max-rows',
      value: 'n',
      description: 'read: para após n linhas; o resultado informa quantas linhas o intervalo possui',
    },
    {
      name: 'where',
      value: 'kind',
      description: `read: apenas células de um tipo (${READ_WHERE.join(' | ')}) como { ref, value, formula }, sem grade de linhas`,
    },
    {
      name: 'stats',
      description:
        'read: contagens (não vazias, fórmulas, erros, números, texto), intervalo usado e lista de planilhas em vez das células',
    },
    {
      name: 'formats',
      description:
        'read: também retorna formatos de célula (negrito, preenchimento, formato de número, alinhamento) por endereço, larguras de coluna e alturas de linha',
    },
    {
      name: 'cells',
      value: 'file',
      description:
        'apply: array JSON de { "cell": "B2", "sheet"?, "value"? | "formula"?, "style"? }; "-" lê da entrada padrão',
    },
    {
      name: 'ops',
      value: 'file',
      description:
        'apply: array JSON de operações DSL da pasta de trabalho (o vocabulário propose_operations do aplicativo; `genoffice guide sheets`); "-" lê da entrada padrão',
    },
    { name: 'dry-run', description: 'apply --ops: valida e mostra o plano sem gravar' },
    ...BATCH_OPTIONS,
    { name: 'out', value: 'path', description: 'apply/links: grava neste caminho em vez de no arquivo original' },
    {
      name: 'force',
      description:
        'apply/links: sobrescreve um arquivo --out existente ou grava enquanto o GenOffice tem o arquivo aberto',
    },
    {
      name: 'update',
      description:
        'links: relê as fontes acessíveis e grava os valores calculados como valores em cache das fórmulas (fórmulas intactas)',
    },
    {
      name: 'break',
      value: 'indices',
      description:
        'links: rompe estes índices de link (base 1, separados por vírgula) — as fórmulas passam a ser seus valores em cache, as partes dos links são removidas e os links restantes são renumerados',
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    switch (verb) {
      case 'read':
        return read(file, args, ctx)
      case 'apply':
        return apply(file, args, ctx)
      case 'check':
        return check(file, ctx)
      case 'links':
        return links(file, args, ctx)
      default:
        throw new CliError(
          EXIT.usage,
          'esperado "sheet read <file>", "sheet apply <file>", "sheet links <file>" ou "sheet check <file>"',
          undefined,
          {
            reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
            suggestion: 'execute `genoffice help sheet`',
          },
        )
    }
  },
}

type Args = Parameters<CommandDef['run']>[0]

/** Findings never fail the command: the exit code is 0 and the agent decides what to fix. */
async function check(file: string | undefined, ctx: CommandContext): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const { issues } = await checkWorkbook(readInput(path))
  return checkResult(
    path,
    issues,
    SHEET_CHECKS,
    'os valores em cache são lidos do arquivo, nada é recalculado; larguras number_overflow são aproximadas (métricas Calibri)',
  )
}

async function read(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const where = readWhere(flagString(args, 'where'))
  const stats = flagBool(args, 'stats')
  const { rows, formulas, ...data } = await readSheet(path, {
    sheet: flagString(args, 'sheet'),
    range: flagString(args, 'range'),
    formats: flagBool(args, 'formats'),
    cols: flagString(args, 'cols'),
    maxRows: positiveInt(flagString(args, 'max-rows'), '--max-rows'),
    where,
    stats,
  })
  const filled = rows.filter((r) => r.some((v) => v !== null)).length
  const head = `${basename(path)} · ${data.sheet}!${data.range}`
  const summary = data.stats
    ? `${head}: ${data.stats.nonEmpty} células não vazias, ${data.stats.formulas} fórmulas, ${data.stats.errors} erros`
    : where
      ? `${head}: ${data.matches} células ${where}`
      : `${head}: ${filled} linhas não vazias`
  const detail: Record<string, unknown> = {
    ...(where || data.stats ? {} : { rows, formulas }),
    ...data,
    units:
      'as linhas seguem ordem base 0 dentro do intervalo; fórmulas e formatos são indexados por endereço A1; recursos descrevem a planilha (mesclagens e hiperlinks: o intervalo)',
  }
  const notes: string[] = []
  if (where && data.matches! > data.cells!.length)
    notes.push(
      `${data.cells!.length} de ${data.matches} células correspondentes listadas; reduza --range ou --cols`,
    )
  if (data.rowsShown < data.rowsTotal) {
    const [, col, row] = /^([A-Z]+)(\d+):/.exec(data.range)!
    const [, endCol] = /:([A-Z]+)\d+$/.exec(data.range)!
    const next = `${col}${Number(row) + data.rowsShown}:${endCol}${Number(row) + data.rowsTotal - 1}`
    notes.push(`${data.rowsShown} de ${data.rowsTotal} linhas mostradas; continue com --range ${next}`)
  }
  if (notes.length) detail.note = notes.join('; ')
  return { summary, detail }
}

function readWhere(value: string | undefined): ReadWhere | undefined {
  if (value === undefined) return undefined
  if ((READ_WHERE as readonly string[]).includes(value)) return value as ReadWhere
  const guess = didYouMean(value, READ_WHERE)
  throw new CliError(
    EXIT.usage,
    `--where deve ser um dos valores ${READ_WHERE.join(', ')} (recebido ${value})`,
    { supported: READ_WHERE },
    {
      reason: 'invalid_argument',
      ...(guess ? { suggestion: `você quis dizer --where ${guess}` } : {}),
    },
  )
}

function positiveInt(value: string | undefined, flag: string): number | undefined {
  if (value === undefined) return undefined
  const n = Number(value)
  if (!Number.isInteger(n) || n < 1)
    throw new CliError(EXIT.usage, `${flag} deve ser um inteiro positivo (recebido ${value})`, undefined, {
      reason: 'invalid_argument',
    })
  return n
}

async function links(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const breakFlag = flagString(args, 'break')
  const update = flagBool(args, 'update')
  if (breakFlag !== undefined && update) {
    throw new CliError(
      EXIT.usage,
      '--break e --update não podem ser combinados: uma atualização relê as fontes, uma ruptura as desconecta',
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  // --break and --update write, so they honour --out/--force, the allowed
  // write roots and the "open in GenOffice" guard like apply does
  const writeOutput = () =>
    resolveOutput(flagString(args, 'out'), ctx, {
      fallback: path,
      force: flagBool(args, 'force'),
      // in-place edits overwrite by design; --out onto another existing file needs --force
      fresh: flagString(args, 'out') !== undefined,
    })
  if (breakFlag !== undefined) {
    const indices = [
      ...new Set(
        breakFlag.split(',').map((part) => {
          const n = Number(part.trim())
          if (!Number.isInteger(n) || n < 1)
            throw new CliError(
              EXIT.usage,
              `--break aceita índices de link com base 1 (recebido ${part.trim()})`,
              undefined,
              { reason: 'invalid_argument' },
            )
          return n
        }),
      ),
    ].sort((a, b) => a - b)
    const output = writeOutput()
    const result = await breakExternalLinks(path, output, indices)
    return {
      summary: `rompidos ${result.removedLinks} link(s): as fórmulas agora são literais, os links restantes foram renumerados`,
      outputPath: output,
      detail: { broken: indices, ...result },
    }
  }
  if (update) {
    const output = writeOutput()
    const result = await updateExternalLinkCaches(path, output)
    return {
      summary:
        `atualizados ${result.updated} cache(s) de fórmula a partir de ${result.reachable} link(s) acessível(is)` +
        (result.broken > 0 ? `; ${result.broken} link(s) rompido(s) mantiveram seus valores em cache` : ''),
      outputPath: output,
      detail: result,
    }
  }
  const links = await readExternalLinks(path)
  const reachable = links.filter((link) => link.reachable).length
  return {
    summary:
      links.length === 0
        ? `${basename(path)} não tem links externos`
        : `${basename(path)}: ${links.length} link(s) externo(s), ${reachable} acessível(is)`,
    detail: {
      links,
      units:
        'índice é o [n] usado pelas fórmulas; source é o caminho resolvido; reachable = o arquivo existe',
    },
  }
}

async function apply(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const cellsSpec = flagString(args, 'cells')
  const opsSpec = flagString(args, 'ops')
  if (!cellsSpec && !opsSpec)
    throw new CliError(EXIT.usage, 'falta --cells <file|-> ou --ops <file|->', undefined, {
      reason: 'missing_argument',
    })
  if (opsSpec) return applyOps(path, opsSpec, args, ctx)
  return applyCells(path, cellsSpec!, args, ctx)
}

function readJsonList(spec: string, flag: string, ctx: CommandContext): unknown[] {
  const text =
    spec === '-' ? readFileSync(0, 'utf-8') : readFileSync(resolveInput(spec, ctx), 'utf-8')
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${flag}: JSON inválido (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  const list = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!list?.length)
    throw new CliError(EXIT.usage, `${flag}: esperado um array JSON não vazio`, undefined, {
      reason: 'invalid_argument',
    })
  return list
}

async function applyOps(
  path: string,
  spec: string,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const ops = readJsonList(spec, '--ops', ctx)
  const source = readFileSync(path)
  const mode = batchMode(args)
  const { outcome: r, failures, applied } = await runBatch(source, ops, mode, args, ctx, path)
  const counts = batchCounts(ops.length, applied, failures.length)
  const detail: Record<string, unknown> = {
    plan: r.plan,
    cells: r.edits.length,
    structural: r.structuralOps.reduce((n, s) => n + s.ops.length, 0),
    ...(hasGatewayPayloads(r.gateway) ? { features: featureCounts(r.gateway) } : {}),
    sheets_changed: r.sheetPlan
      ? r.sheetPlan.additions.length + r.sheetPlan.removals.length + r.sheetPlan.renames.length
      : 0,
  }
  for (const message of r.warnings) ctx.warn({ code: 'op_warning', message })
  if (flagBool(args, 'dry-run')) {
    return batchResult(
      { summary: `simulação: ${applied} de ${ops.length} operações validadas, nada gravado`, detail },
      counts,
      failures,
    )
  }
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    // in-place edits overwrite by design; --out onto another existing file needs --force
    fresh: flagString(args, 'out') !== undefined,
  })
  const w = await writeWorkbook(source, r.edits, output, {
    plan: r.sheetPlan,
    structuralOps: r.structuralOps,
    renames: r.renames,
    gateway: r.gateway,
  })
  return batchResult(
    {
      summary: `aplicadas ${applied} de ${ops.length} operações a ${basename(output)}`,
      outputPath: output,
      detail: {
        ...detail,
        formulas: w.formulas,
        cached_values: w.cachedValues,
      },
      ...writeWarnings(w),
    },
    counts,
    failures,
  )
}

/**
 * The DSL plans the whole list against one in-memory workbook and rejects on
 * the first bad op. A non-atomic run drops that op (best effort) or everything
 * from it on (stop on error) and plans the rest again from the file, so an op
 * that only worked because of a rejected one fails too instead of half-applying.
 * Rejections the DSL cannot pin to one op (a forbidden mix of ops) stay fatal.
 */
async function runBatch(
  source: Buffer,
  ops: unknown[],
  mode: BatchMode,
  args: Args,
  ctx: CommandContext,
  sourcePath: string,
): Promise<{ outcome: DslOutcome; failures: OpFailure[]; applied: number }> {
  const failures: OpFailure[] = []
  // the DSL's own error per dropped op: its detail (sheets, supported, …) and hints are the contract
  const dropped: CliError[] = []
  let pending = ops.map((op, index) => ({ op, index }))
  for (;;) {
    try {
      const outcome = await runWorkbookDsl(
        source,
        pending.map((p) => p.op),
        flagString(args, 'sheet'),
        ctx,
        { sourcePath },
      )
      return { outcome, failures, applied: pending.length }
    } catch (err) {
      const f = pinnedFailure(err)
      if (mode === 'atomic' || !f || !pending[f.index] || !(err instanceof CliError)) {
        throw mode === 'atomic' && err instanceof CliError ? failedBatch(err, ops.length) : err
      }
      const index = pending[f.index]!.index
      failures.push({ ...f, index })
      dropped.push(
        new CliError(
          err.code,
          err.message.replace(/^ops\[\d+\]/, `ops[${index}]`),
          { ...err.detail, failures: [{ ...f, index }] },
          { reason: err.reason, suggestion: err.suggestion },
        ),
      )
      pending =
        mode === 'stop_on_error'
          ? pending.slice(0, f.index)
          : pending.filter((_, i) => i !== f.index)
      if (!pending.length) {
        // gateway rejections surface after planning, so the first one found is not always the lowest index
        const first = [...dropped].sort(
          (a, b) =>
            (a.detail!.failures as OpFailure[])[0]!.index -
            (b.detail!.failures as OpFailure[])[0]!.index,
        )[0]!
        throw new CliError(
          first.code,
          `nenhuma operação foi aplicada: ${first.message}`,
          { ...first.detail, failures, batch: batchCounts(ops.length, 0, failures.length) },
          { reason: first.reason, suggestion: first.suggestion },
        )
      }
    }
  }
}

/** What the batch writes beyond cells, for the plan report. */
function featureCounts(g: GatewayPayloads): Record<string, number> {
  const counts: Record<string, number> = {
    charts: g.visualAdditions.filter((v) => v.chart).length + g.chartEdits.length,
    images: g.visualAdditions.filter((v) => v.image).length,
    shapes: g.visualAdditions.filter((v) => v.shape).length,
    tables: g.tableAdditions.length,
    hyperlinks: g.hyperlinkEdits.reduce((n, s) => n + s.edits.length, 0),
    notes: g.noteStates.length,
    filters: g.filterStates.length,
    conditional_formats: g.cfStates.reduce((n, s) => n + s.rules.length, 0),
    data_validations: g.dvStates.reduce((n, s) => n + s.rules.length, 0),
    page_setup: g.pageSetupStates.length,
    hidden_ranges: g.hiddenOps.reduce((n, s) => n + s.ops.length, 0),
    sheet_tabs: g.tabs.moves.length + g.tabs.hidden.length + g.tabs.duplicates.length,
    protections: g.sheetProtections.length,
    defined_names: g.definedNamesState?.names.length ?? 0,
    pivots: g.pivotAdditions.length,
    sparklines: g.sparklineAdditions.reduce((n, s) => n + s.cells.length, 0),
  }
  return Object.fromEntries(Object.entries(counts).filter(([, n]) => n > 0))
}

async function applyCells(
  path: string,
  spec: string,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const inputs = readJsonList(spec, '--cells', ctx)
  // addresses are checked before the workbook engine is even started
  cellEditsFromInputs(inputs as CellInput[], '')
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    fallback: path,
    force: flagBool(args, 'force'),
    // in-place edits overwrite by design; --out onto another existing file needs --force
    fresh: flagString(args, 'out') !== undefined,
  })
  const summary = await workbookSummary(path)
  const defaultSheet = flagString(args, 'sheet') ?? summary.activeSheet ?? summary.sheets[0]?.name
  if (!defaultSheet) throw new CliError(EXIT.conversion, 'a pasta de trabalho não tem planilhas')
  const known = new Set(summary.sheets.map((s) => s.name))
  const edits = cellEditsFromInputs(inputs as CellInput[], defaultSheet)
  const unknown = edits.find((e) => !known.has(e.sheetName))
  if (unknown) {
    throw new CliError(
      EXIT.usage,
      `planilha não encontrada: ${unknown.sheetName}`,
      { sheets: [...known] },
      sheetNotFoundHints(unknown.sheetName, [...known]),
    )
  }
  const r = await writeWorkbook(readFileSync(path), edits, output)
  return {
    summary: `gravadas ${r.cells} células em ${basename(output)}`,
    outputPath: output,
    detail: {
      cells: r.cells,
      formulas: r.formulas,
      cached_values: r.cachedValues,
    },
    ...writeWarnings(r),
  }
}

function writeWarnings(w: {
  warning?: string
  formulaError?: string
  dependentsWarning?: string
  notes?: Warning[]
}): {
  warnings?: Warning[]
} {
  const warnings = [
    ...(w.warning ? [FORMULAS_NOT_CACHED(w.warning)] : []),
    ...(w.formulaError ? [FORMULA_ERRORS(w.formulaError)] : []),
    ...(w.dependentsWarning ? [DEPENDENTS_NOT_RECALCULATED(w.dependentsWarning)] : []),
    ...(w.notes ?? []),
  ]
  return warnings.length ? { warnings } : {}
}

export { SUPPORTED_DSL_OPS }
