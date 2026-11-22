import { existsSync, readFileSync, statSync } from 'node:fs'
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path'
import { flagBool, flagString } from '../args'
import { assertAllowed, resolveInput, resolveOutput, writeOutput } from '../fs'
import { applyOps, blankDeck, inlineLocalFiles, parseOps, saveDeck } from '../formats/pptx'
import {
  auditDeckBytes,
  buildDeckFromDir,
  buildDeckFromSpec,
  deckAuditDetail,
  stageContext,
} from '../formats/slide-spec'
import { outputDirectory, renderToPngs } from '../formats/render'
import {
  additionPlan,
  blankWorkbook,
  cellEditsFromTable,
  csvTable,
  writeWorkbook,
  type CsvTableOptions,
  type TableInput,
} from '../formats/xlsx'
import { exportViaApp } from '../formats/app-export'
import { blankDocument, closeDocument, fillFromHtml, saveDocument } from '../formats/docx'
import { markdownToDocx } from '../formats/markdown'
import { runWorkbookDsl } from '../formats/xlsx-dsl'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import { sheetNameFromStem } from '../formats/csv'
import { readOpsInput } from '../ops-input'
import type { CommandContext, CommandDef } from '../registry'
import type { TxnResult } from '@genoffice/pptx-ops'
import { CliError, EXIT, type CommandResult } from '../result'
import { txnFailure, txnDetail } from './txn'

/** Document types that can be built here. */
const TYPES = ['pptx', 'xlsx', 'docx', 'pdf'] as const

export const createCommand: CommandDef = {
  name: 'create',
  summary: 'Cria um novo documento a partir de conteúdo estruturado.',
  usage:
    'create --type <pptx|xlsx|docx|pdf> (--ops <file|-> | --spec <file|dir|-> [--outline <file>] | --from <file>) --out <path> [--force]',
  options: [
    { name: 'type', value: 'type', description: `tipo de documento: ${TYPES.join(', ')}` },
    {
      name: 'ops',
      value: 'file',
      description:
        'pptx: array JSON de operações (ou {"ops":[...]}) aplicadas uma a uma a uma apresentação em branco de um slide (operações posteriores podem atingir slides adicionados antes); "-" lê da stdin. Veja `genoffice guide slides`.',
    },
    {
      name: 'spec',
      value: 'file',
      description:
        'pptx: uma especificação de apresentação (páginas com texto, formas e imagens posicionadas em px numa tela de 1280x720) convertida em um slide por página, ou um diretório de arquivos de especificação de uma página cada, lidos em ordem de nome; "-" lê da stdin. Veja `genoffice guide slides spec` e `genoffice guide slides design`.',
    },
    {
      name: 'outline',
      value: 'file',
      description:
        'pptx --spec <dir>: o roteiro da apresentação a partir do qual as páginas foram escritas (padrão: outline.json dentro ou ao lado do diretório); cada arquivo de página é verificado contra sua entrada, e a compilação só começa quando todas as páginas do roteiro tiverem seu arquivo',
    },
    {
      name: 'from',
      value: 'file',
      description:
        'xlsx: um .csv ou um .json com uma matriz 2D de valores de células ou { "sheets": [{ "name", "rows" }] }; strings que começam com "=" são fórmulas. docx: um arquivo .md ou .html contendo um fragmento de HTML restrito (veja `genoffice guide docs`). pdf: qualquer arquivo .md/.html/.docx/.xlsx/.pptx, impresso pelo renderizador do GenOffice',
    },
    {
      name: 'header',
      description:
        'xlsx: trata a linha 1 de cada planilha como cabeçalho — congela-a e aplica um AutoFilter à área usada',
    },
    {
      name: 'decimal',
      value: ', or .',
      description:
        'xlsx --from .csv: o separador decimal usado no arquivo (padrão "."); com "," um "." é lido como separador de milhar',
    },
    {
      name: 'no-generator-fields',
      description:
        'docx: remove os campos geradores Application / AppVersion de docProps/app.xml (o campo Template padrão do Word permanece)',
    },
    { name: 'out', value: 'path', description: 'arquivo de saída (obrigatório)' },
    { name: 'force', description: 'substituir um arquivo de saída existente' },
    {
      name: 'render',
      value: 'dir',
      description:
        'pptx: após gravar, renderiza um PNG por slide em <dir> (padrão: <output name>-previews ao lado do arquivo) e os lista em detail.previews; inicia um processo oculto do GenOffice por alguns segundos',
    },
    {
      name: 'audit',
      description:
        'pptx: executa a auditoria de geometria na apresentação gravada e a relata em detail.audit (constatações informativas, com sugestões de setTransform)',
    },
  ],
  async run(args, ctx) {
    const type = flagString(args, 'type')?.toLowerCase()
    if (!type)
      throw new CliError(EXIT.usage, 'falta --type <type>', undefined, {
        reason: 'missing_argument',
      })
    if (!(TYPES as readonly string[]).includes(type)) {
      throw new CliError(
        EXIT.usage,
        `ainda não é possível criar .${type}`,
        { supported: [...TYPES] },
        { reason: 'unsupported', suggestion: 'escolha um tipo em detail.supported' },
      )
    }
    if (type !== 'pptx' && (flagBool(args, 'audit') || args.flags['render'] !== undefined)) {
      throw new CliError(
        EXIT.usage,
        '--render e --audit pré-visualizam uma apresentação construída; usam --type pptx',
        undefined,
        { reason: 'invalid_argument' },
      )
    }
    const output = resolveOutput(flagString(args, 'out'), ctx, {
      force: flagBool(args, 'force'),
      fresh: true,
    })
    if (type === 'pdf') {
      const from = flagString(args, 'from')
      if (!from)
        throw new CliError(EXIT.usage, 'pdf precisa de --from <document>', undefined, {
          reason: 'missing_argument',
        })
      const source = resolveInput(from, ctx)
      const r = await exportViaApp(source, 'pdf', output, { env: ctx.env, log: ctx.log })
      return {
        summary: `criado ${basename(output)}`,
        outputPath: output,
        detail: { via: 'genoffice --headless-export', summary: r.summary },
      }
    }
    if (type === 'docx') {
      const r = await createDocx(args, ctx)
      writeOutput(output, r.bytes)
      return { summary: `criado ${basename(output)}`, outputPath: output, detail: r.detail }
    }
    if (type === 'xlsx') {
      const r = await createXlsx(args, ctx, output)
      return {
        summary: `criado ${basename(output)} (${r.sheets} planilhas)`,
        outputPath: output,
        detail: r,
      }
    }
    const result =
      flagString(args, 'spec') !== undefined
        ? await createPptxFromSpec(args, ctx)
        : await createPptx(args, ctx)
    writeOutput(output, result.bytes)
    const detail: Record<string, unknown> = { ...result.detail }
    // The deck is already on disk: a preview or audit failure must not turn the
    // command into an error that hides outputPath.
    if (flagBool(args, 'audit')) {
      try {
        detail.audit = deckAuditDetail(await auditDeckBytes(result.bytes))
      } catch (err) {
        ctx.warn({
          code: 'audit_failed',
          message: `apresentação gravada; a auditoria falhou: ${errorMessage(err)}`,
        })
      }
    }
    const renderFlag = args.flags['render']
    if (renderFlag !== undefined) {
      try {
        const dir =
          typeof renderFlag === 'string'
            ? outputDirectory(renderFlag, ctx)
            : assertAllowed(
                join(dirname(output), `${basename(output, extname(output))}-previews`),
                ctx.env,
                'write',
              )
        const files = await renderToPngs(output, ctx, { outDir: dir, scale: 1, log: ctx.log })
        detail.previews = files.map(({ page, ...f }) => ({ slide: page, ...f }))
      } catch (err) {
        ctx.warn({
          code: 'render_failed',
          message: `apresentação gravada; falha ao renderizar a pré-visualização: ${errorMessage(err)}`,
        })
      }
    }
    return {
      summary: `criado ${basename(output)} (${result.slides} slides)`,
      outputPath: output,
      detail,
    }
  },
}

async function createPptx(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<{ bytes: Uint8Array; slides: number; detail: CommandResult['detail'] }> {
  const { text, source } = readOpsInput(args, ctx)
  const ops = inlineLocalFiles(parseOps(text, source), ctx)
  const opened = await blankDeck()
  // One transaction per op: the executor validates a batch against the
  // pre-transaction deck, so slides added earlier in the list would not yet
  // exist for later ops. Nothing is written unless every op applies.
  const records: NonNullable<TxnResult['records']> = []
  for (const [index, op] of ops.entries()) {
    const r = applyOps(opened, [op], { isolation: 'atomic' })
    if (!r.applied) throw txnFailure(r, index)
    records.push(...(r.records ?? []))
  }
  return {
    bytes: await saveDeck(opened),
    slides: opened.deck.slides.length,
    detail: txnDetail({ applied: true, records }),
  }
}

async function createPptxFromSpec(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<{ bytes: Uint8Array; slides: number; detail: CommandResult['detail'] }> {
  const spec = flagString(args, 'spec')!
  const next =
    'genoffice slides audit <file> para a auditoria de geometria; genoffice slides render <file> --out <dir> para PNGs; genoffice slides replace <file> --slide n --spec <page.json> para reconstruir uma página'
  const outlinePath = flagString(args, 'outline')
  const dir = spec === '-' ? null : specDirectory(spec, ctx)
  if (dir) {
    const stage = stageContext(dir, ctx, outlinePath ? resolveInput(outlinePath, ctx) : undefined)
    const built = await buildDeckFromDir(dir, ctx, stage)
    return {
      bytes: built.bytes,
      slides: built.pages,
      detail: {
        via: 'diretório de especificação de apresentação',
        files: built.files.map((f) => basename(f)),
        issues: built.issues,
        imageFailures: built.imageFailures,
        outline: stage.outline
          ? { file: stage.outline.file, findings: built.outlineFindings }
          : null,
        style: stage.style ? { file: stage.style.file, offPalette: built.offPalette } : null,
        notes: stage.notes,
        next,
      },
    }
  }
  if (outlinePath)
    throw new CliError(EXIT.usage, '--outline acompanha --spec <directory>', undefined, {
      reason: 'invalid_argument',
    })
  const source = spec === '-' ? 'stdin' : resolveInput(spec, ctx)
  const text = spec === '-' ? readFileSync(0, 'utf-8') : readFileSync(source, 'utf-8')
  const built = await buildDeckFromSpec(
    text,
    source,
    ctx,
    spec === '-' ? undefined : dirname(source),
  )
  return {
    bytes: built.bytes,
    slides: built.pages,
    detail: { via: 'especificação de apresentação', issues: built.issues, imageFailures: built.imageFailures, next },
  }
}

function specDirectory(spec: string, ctx: CommandContext): string | null {
  const abs = isAbsolute(spec) ? spec : resolve(ctx.cwd, spec)
  if (!existsSync(abs) || !statSync(abs).isDirectory()) return null
  return assertAllowed(abs, ctx.env, 'read')
}

async function createXlsx(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
  output: string,
): Promise<Record<string, unknown>> {
  const from = flagString(args, 'from')
  if (!from)
    throw new CliError(EXIT.usage, 'xlsx precisa de --from <data.csv|table.json>', undefined, {
      reason: 'missing_argument',
    })
  const source = resolveInput(from, ctx)
  const decimal = flagString(args, 'decimal')
  if (decimal !== undefined && decimal !== ',' && decimal !== '.') {
    throw new CliError(EXIT.usage, `--decimal deve ser "," ou ".", recebido "${decimal}"`, undefined, {
      reason: 'invalid_argument',
    })
  }
  const tables = readTables(source, basename(source, extname(source)), { decimal })
  const [first, ...rest] = tables
  if (!first)
    throw new CliError(EXIT.usage, `${from}: sem planilhas`, undefined, { reason: 'invalid_argument' })
  const edits = tables.flatMap((t) => cellEditsFromTable(t.name, t.rows, t.formats))
  let r = await writeWorkbook(await blankWorkbook(first.name), edits, output, {
    plan: additionPlan(
      first.name,
      rest.map((t) => t.name),
    ),
  })
  if (flagBool(args, 'header')) {
    const ops = tables.flatMap((t) => headerOps(t))
    const dsl = await runWorkbookDsl(readFileSync(output), ops, first.name, ctx, {
      sourcePath: output,
    })
    const w = await writeWorkbook(readFileSync(output), dsl.edits, output, {
      plan: dsl.sheetPlan,
      structuralOps: dsl.structuralOps,
      renames: dsl.renames,
      gateway: dsl.gateway,
    })
    r = {
      ...r,
      warning: r.warning ?? w.warning,
      formulaError: r.formulaError ?? w.formulaError,
      dependentsWarning: r.dependentsWarning ?? w.dependentsWarning,
    }
    for (const note of w.notes ?? []) ctx.warn(note)
  }

  if (r.warning) ctx.warn({ code: 'formulas_not_cached', message: r.warning })
  if (r.formulaError) ctx.warn({ code: 'formula_errors', message: r.formulaError })
  if (r.dependentsWarning)
    ctx.warn({ code: 'dependents_not_recalculated', message: r.dependentsWarning })
  return {
    sheets: tables.length,
    cells: r.cells,
    formulas: r.formulas,
    cached_values: r.cachedValues,
  }
}

/** Freeze row 1 and filter the used range: what Excel's own import wizard does for a header row. */
function headerOps(table: TableInput): Record<string, unknown>[] {
  const width = table.rows.reduce((max, row) => Math.max(max, row.length), 0)
  if (table.rows.length < 1 || width < 1) return []
  const range = `A1:${columnLabel(width - 1)}${Math.max(table.rows.length, 2)}`
  return [
    { op: 'set_freeze', sheet: table.name, rows: 1, columns: 0 },
    { op: 'set_filter', sheet: table.name, range },
  ]
}

function readTables(
  source: string,
  fallbackName: string,
  opts: CsvTableOptions = {},
): TableInput[] {
  const sheetName = sheetNameFromStem(fallbackName)
  if (extname(source).toLowerCase() === '.csv') {
    return [csvTable(readFileSync(source), sheetName, opts)]
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(readFileSync(source, 'utf-8'))
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${source}: JSON inválido (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  if (Array.isArray(parsed)) return [{ name: sheetName, rows: parsed as TableInput['rows'] }]
  const sheets = (parsed as { sheets?: unknown })?.sheets
  if (!Array.isArray(sheets) || sheets.length === 0) {
    throw new CliError(
      EXIT.usage,
      `${source}: esperava uma matriz 2D ou { "sheets": [{ "name", "rows" }] }`,
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  return sheets.map((s, i) => {
    const t = s as Partial<TableInput>
    if (typeof t.name !== 'string' || !Array.isArray(t.rows)) {
      throw new CliError(EXIT.usage, `${source}: sheets[${i}] precisa de "name" e "rows"`)
    }
    return { name: t.name, rows: t.rows }
  })
}

async function createDocx(
  args: Parameters<CommandDef['run']>[0],
  ctx: CommandContext,
): Promise<{ bytes: Uint8Array; detail: Record<string, unknown> }> {
  const from = flagString(args, 'from')
  if (!from)
    throw new CliError(EXIT.usage, 'docx precisa de --from <content.md|fragment.html>', undefined, {
      reason: 'missing_argument',
    })
  const blankOptions = flagBool(args, 'no-generator-fields')
    ? { omitGeneratorFields: true }
    : undefined
  const source = resolveInput(from, ctx)
  const text = readFileSync(source, 'utf-8')
  const ext = extname(source).toLowerCase()
  if (ext === '.md' || ext === '.markdown') {
    return {
      bytes: await markdownToDocx(text, { file: source, ctx }, blankOptions),
      detail: { source: 'markdown' },
    }
  }
  if (ext === '.html' || ext === '.htm') {
    const doc = await blankDocument(blankOptions)
    try {
      const blocks = fillFromHtml(doc, text)
      return { bytes: await saveDocument(doc), detail: { source: 'html', blocks } }
    } finally {
      closeDocument(doc)
    }
  }
  throw new CliError(
    EXIT.usage,
    `docx --from precisa de um arquivo .md ou .html, recebido ${ext || 'sem extensão'}`,
    undefined,
    { reason: 'unsupported' },
  )
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err)
}
