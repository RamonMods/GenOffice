import { basename, dirname } from 'node:path'
import { DEFAULT_COMPARE_AUTHOR } from '@genoffice/docx-engine'
import { flagBool, flagString } from '../args'
import {
  PREVIEW_CHARS,
  applyDocOps,
  type DocOpsOutcome,
  blockRangeHtml,
  closeDocument,
  compareDocxFiles,
  describeDocument,
  headerFooterState,
  listComments,
  listFields,
  listNotes,
  listProperties,
  listTableStyles,
  listRevisions,
  countRevisionsElsewhere,
  listStyles,
  openDocument,
  saveDocument,
} from '../formats/docx'
import { previewChars } from '../preview'
import { checkDocument, DOCS_CHECKS } from '../formats/docx-check'
import { listSections } from '../formats/docx-sections'
import { readInput, resolveInput, resolveOutput, writeOutput } from '../fs'
import { readOpsInput } from '../ops-input'
import type { CommandContext, CommandDef } from '../registry'
import { checkResult } from '../check'
import { CliError, EXIT, type CommandResult } from '../result'
import { BATCH_OPTIONS, batchCounts, batchMode, batchResult, failedBatch } from '../batch'
import { opSuggestion } from '../op-errors'

export const docsCommand: CommandDef = {
  name: 'docs',
  summary: 'Lê ou edita um .docx com as mesmas operações da IA integrada ao aplicativo.',
  usage:
    'docs <read|apply|check> <file.docx> [options] | docs compare <original.docx> <revised.docx> --out <path>',
  options: [
    { name: 'range', value: 'a-b', description: 'read: intervalo de índices de blocos (padrão: todos)' },
    { name: 'html', description: 'read: também retorna o intervalo como HTML restrito' },
    { name: 'full', description: 'read: texto integral do bloco em vez de uma pré-visualização de 200 caracteres' },
    {
      name: 'max-chars',
      value: 'n',
      description: 'read: tamanho da pré-visualização por bloco (padrão 200); texto truncado termina em …(+n chars)',
    },
    {
      name: 'comments',
      description:
        'read: lista todos os comentários (id, autor, data, id pai, bloco de ancoragem e texto, estado de resolução) para reply_comment / resolve_comment / delete_comment',
    },
    {
      name: 'revisions',
      description:
        'read: lista as alterações controladas pendentes (ids para accept_changes / reject_changes, tipo, autor, índice do bloco, texto)',
    },
    {
      name: 'styles',
      description:
        'read: lista os estilos definidos no documento (ids para applyStyle / define_style, tipo, pai, contagem de uso)',
    },
    {
      name: 'header-footer',
      description: 'read: o texto de cabeçalho e rodapé por variante (tokens {PAGE} / {NUMPAGES})',
    },
    {
      name: 'sections',
      description:
        'read: configuração de página por seção (papel, orientação, margens, colunas, intervalo de blocos) para set_page_setup',
    },
    {
      name: 'fields',
      description:
        'read: lista todos os campos (SEQ, REF, DATE, PAGE, TOC …) com seu índice de bloco, instrução, resultado em cache e se o Word o recalcula ao abrir',
    },
    {
      name: 'notes',
      description: 'read: lista notas de rodapé e notas finais (ids para delete_note, índice do bloco de ancoragem)',
    },
    {
      name: 'properties',
      description:
        'read: as propriedades do documento do arquivo (core, app, custom e customTypes — o tipo vt em que cada valor personalizado é armazenado)',
    },
    {
      name: 'ops',
      value: 'file',
      description:
        'apply: array JSON de operações (entradas de apply_ops, mais insert_content / replace_blocks com um campo html, insert_image / insert_picture / insert_text_box, insert_chart / edit_chart, set_header_footer, set_page_setup / insert_section_break, set_watermark, define_style, add_comment / reply_comment / resolve_comment / delete_comment, accept_changes / reject_changes, insert_footnote / insert_endnote / delete_note, wrap_content_control / unwrap_content_control, remove_parts (remove do pacote as partes customXml / web-extension para limpeza de metadados)); "-" lê da stdin. Veja `genoffice guide docs`.',
    },
    {
      name: 'track',
      description:
        'apply: registra edições de conteúdo como alterações controladas (autor: --author, padrão "AI Assistant")',
    },
    {
      name: 'author',
      value: 'name',
      description:
        'apply --track / compare: autor das alterações controladas (padrão na comparação: GenOffice)',
    },
    {
      name: 'date',
      value: 'iso',
      description: 'compare: data e hora registradas em cada alteração (padrão: agora)',
    },
    {
      name: 'granularity',
      value: 'word|char',
      description: 'compare: granularidade usada ao comparar um parágrafo alterado (padrão: word)',
    },
    {
      name: 'ignore',
      value: 'list',
      description: 'compare: lista separada por vírgulas do que ignorar (formatting, case, whitespace)',
    },
    {
      name: 'dry-run',
      description: 'apply: executa o lote em memória e relata cada etapa sem gravar',
    },
    ...BATCH_OPTIONS,
    {
      name: 'out',
      value: 'path',
      description:
        'apply: grava aqui em vez de no próprio arquivo; compare: o documento com alterações marcadas a gravar (obrigatório)',
    },
    {
      name: 'force',
      description:
        'apply / compare: substitui um arquivo --out existente, ou grava enquanto o GenOffice mantém o arquivo aberto',
    },
  ],
  async run(args, ctx) {
    const [verb, file, second] = args.positionals
    switch (verb) {
      case 'read':
        return read(file, args, ctx)
      case 'apply':
        return apply(file, args, ctx)
      case 'check':
        return check(file, ctx)
      case 'compare':
        return compare(file, second, args, ctx)
      default:
        throw new CliError(
          EXIT.usage,
          'esperava "docs read <file>", "docs apply <file>", "docs check <file>" ou "docs compare <original> <revised>"',
          undefined,
          {
            reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
            suggestion: 'execute `genoffice help docs`',
          },
        )
    }
  },
}

type Args = Parameters<CommandDef['run']>[0]

function parseRange(spec: string | undefined, count: number): [number, number] | undefined {
  if (spec === undefined) return undefined
  const m = /^(\d+)(?:-(\d+))?$/.exec(spec)
  if (!m) {
    throw new CliError(EXIT.usage, '--range deve ter o formato 3 ou 3-10 (índices de blocos)', undefined, {
      reason: 'invalid_argument',
    })
  }
  const start = Number(m[1])
  const end = m[2] === undefined ? start : Number(m[2])
  if (start > end || end >= count) {
    throw new CliError(
      EXIT.usage,
      `--range fora do intervalo (0-${count - 1})`,
      { valid_range: [0, count - 1] },
      { reason: 'out_of_range', suggestion: `use índices de bloco entre 0 e ${count - 1}` },
    )
  }
  return [start, end]
}

async function read(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const doc = await openDocument(readInput(path))
  try {
    const count = doc.editor.state.doc.childCount
    const range = parseRange(flagString(args, 'range'), count)
    const blocks = describeDocument(doc, range, previewChars(args, PREVIEW_CHARS))
    const detail: Record<string, unknown> = {
      blocks: count,
      range: range ? `${range[0]}-${range[1]}` : `0-${count - 1}`,
      items: blocks,
      units: 'os índices de bloco são posições baseadas em 0 no corpo; as operações os usam como alvo',
    }
    if (blocks.some((b) => b.table)) detail.tableStyles = listTableStyles(doc)
    if (flagBool(args, 'html')) {
      const [s, e] = range ?? [0, count - 1]
      detail.html = blockRangeHtml(doc, s, e)
    }
    if (flagBool(args, 'comments')) detail.comments = listComments(doc)
    if (flagBool(args, 'revisions')) {
      detail.revisions = listRevisions(doc)
      const elsewhere = await countRevisionsElsewhere(doc)
      if (elsewhere)
        detail.revisionsElsewhere = {
          count: elsewhere,
          note: 'alterações controladas em cabeçalhos, rodapés, notas de rodapé ou notas finais (não listadas); accept_changes / reject_changes com all: true ou author / type / before as incluem',
        }
    }
    if (flagBool(args, 'styles')) detail.styles = listStyles(doc)
    if (flagBool(args, 'header-footer')) detail.headerFooter = headerFooterState(doc)
    if (flagBool(args, 'sections')) detail.sections = listSections(doc)
    if (flagBool(args, 'fields')) detail.fields = listFields(doc)
    if (flagBool(args, 'notes')) detail.notes = listNotes(doc)
    if (flagBool(args, 'properties')) detail.properties = listProperties(doc)
    return { summary: `${basename(path)}: ${count} blocos`, detail }
  } finally {
    closeDocument(doc)
  }
}

async function check(file: string | undefined, ctx: CommandContext): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const doc = await openDocument(readInput(path))
  try {
    return checkResult(path, checkDocument(doc), DOCS_CHECKS)
  } finally {
    closeDocument(doc)
  }
}

function trackFlag(args: Args): string | undefined {
  if (!flagBool(args, 'track')) return undefined
  return flagString(args, 'author')?.trim() || 'AI Assistant'
}

async function apply(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  const { text, source } = readOpsInput(args, ctx)
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch (err) {
    throw new CliError(
      EXIT.usage,
      `${source}: JSON inválido (${(err as Error).message})`,
      undefined,
      { reason: 'invalid_json' },
    )
  }
  const ops = Array.isArray(parsed)
    ? parsed
    : parsed && typeof parsed === 'object' && Array.isArray((parsed as { ops?: unknown }).ops)
      ? (parsed as { ops: unknown[] }).ops
      : null
  if (!ops?.length)
    throw new CliError(EXIT.usage, `${source}: esperava um array não vazio de operações`, undefined, {
      reason: 'invalid_argument',
    })
  const dryRun = flagBool(args, 'dry-run')
  const mode = batchMode(args)
  const trackAuthor = trackFlag(args)
  const doc = await openDocument(readInput(path))
  try {
    let outcome: DocOpsOutcome
    try {
      outcome = await applyDocOps(doc, ops as Record<string, unknown>[], {
        ctx,
        mode,
        baseDir: source === 'stdin' ? undefined : dirname(source),
        ...(trackAuthor ? { trackAuthor } : {}),
      })
    } catch (err) {
      throw err instanceof CliError ? failedBatch(err, ops.length) : err
    }
    const { results, failures } = outcome
    const counts = batchCounts(ops.length, results.length, failures.length)
    if (!results.length) {
      const first = failures[0]!
      throw new CliError(
        EXIT.usage,
        `nenhuma operação foi aplicada: op ${first.index} (${first.op}) rejeitada: ${first.error}`,
        { failures, batch: counts },
        { reason: first.reason, suggestion: opSuggestion(first, 'docs') },
      )
    }
    if (dryRun) {
      return batchResult(
        {
          summary: `simulação: ${results.length} de ${ops.length} operações validadas, nada gravado`,
          detail: {
            plan: results.map((r) => `${r.index}: ${r.op}: ${r.output.split('\n')[0]}`),
          },
        },
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
    writeOutput(output, await saveDocument(doc))
    return batchResult(
      {
        summary: `aplicadas ${results.length} de ${ops.length} operações a ${basename(output)}`,
        outputPath: output,
        detail: { results, blocks: doc.editor.state.doc.childCount },
      },
      counts,
      failures,
    )
  } finally {
    closeDocument(doc)
  }
}

const COMPARE_GRANULARITIES = ['word', 'char'] as const
const COMPARE_IGNORES = ['formatting', 'case', 'whitespace'] as const
/** an ISO 8601 timestamp, the shape Word writes in w:date */
const ISO_DATE =
  /^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/

function parseGranularity(spec: string | undefined): (typeof COMPARE_GRANULARITIES)[number] {
  if (spec === undefined) return 'word'
  const value = spec.trim().toLowerCase()
  if (!COMPARE_GRANULARITIES.includes(value as (typeof COMPARE_GRANULARITIES)[number])) {
    throw new CliError(
      EXIT.usage,
      `--granularity deve ser um de ${COMPARE_GRANULARITIES.join(', ')}`,
      { valid: [...COMPARE_GRANULARITIES] },
      { reason: 'invalid_argument', suggestion: 'use --granularity word' },
    )
  }
  return value as (typeof COMPARE_GRANULARITIES)[number]
}

function parseIgnore(spec: string | undefined): (typeof COMPARE_IGNORES)[number][] {
  if (spec === undefined) return []
  const out: (typeof COMPARE_IGNORES)[number][] = []
  for (const part of spec.split(',')) {
    const value = part.trim().toLowerCase()
    if (!value) continue
    if (!COMPARE_IGNORES.includes(value as (typeof COMPARE_IGNORES)[number])) {
      throw new CliError(
        EXIT.usage,
        `--ignore aceita uma lista separada por vírgulas de ${COMPARE_IGNORES.join(', ')}`,
        { valid: [...COMPARE_IGNORES] },
        { reason: 'invalid_argument', suggestion: 'use --ignore case,whitespace' },
      )
    }
    out.push(value as (typeof COMPARE_IGNORES)[number])
  }
  return out
}

function parseDate(spec: string | undefined): string {
  if (spec === undefined) return new Date().toISOString()
  const value = spec.trim()
  if (!ISO_DATE.test(value) || Number.isNaN(Date.parse(value))) {
    throw new CliError(
      EXIT.usage,
      '--date deve ser um timestamp ISO, ex.: 2026-10-16T00:00:00Z',
      undefined,
      { reason: 'invalid_argument', suggestion: 'use --date 2026-10-16T00:00:00Z' },
    )
  }
  return value
}

/**
 * The redline of two documents: the revised one with the differences marked as
 * tracked insertions and deletions, the way Word's Review > Compare writes it.
 */
async function compare(
  originalFile: string | undefined,
  revisedFile: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const originalPath = resolveInput(originalFile, ctx)
  const revisedPath = resolveInput(revisedFile, ctx)
  const granularity = parseGranularity(flagString(args, 'granularity'))
  const ignore = parseIgnore(flagString(args, 'ignore'))
  const date = parseDate(flagString(args, 'date'))
  const author = flagString(args, 'author')?.trim() || DEFAULT_COMPARE_AUTHOR
  const output = resolveOutput(flagString(args, 'out'), ctx, {
    force: flagBool(args, 'force'),
    fresh: true,
  })
  const { bytes, summary } = await compareDocxFiles(
    readInput(originalPath),
    readInput(revisedPath),
    { author, date, granularity, ignore },
  )
  writeOutput(output, bytes)
  return {
    summary:
      `${basename(originalPath)} vs ${basename(revisedPath)}: ` +
      `${summary.insertedParagraphs} inseridos, ${summary.deletedParagraphs} excluídos, ` +
      `${summary.changedParagraphs} parágrafo(s) alterado(s) em ${basename(output)}`,
    outputPath: output,
    detail: {
      original: originalPath,
      revised: revisedPath,
      output,
      insertions: summary.insertions,
      deletions: summary.deletions,
      inserted_paragraphs: summary.insertedParagraphs,
      deleted_paragraphs: summary.deletedParagraphs,
      changed_paragraphs: summary.changedParagraphs,
      unchanged_paragraphs: summary.unchangedParagraphs,
      author,
      date,
      granularity,
      ignored: ignore,
    },
  }
}
