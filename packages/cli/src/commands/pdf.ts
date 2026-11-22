import { basename } from 'node:path'
import { PdfLoadError } from '@genoffice/pdf2docx'
import { flagBool, flagString } from '../args'
import { PDF_DEFAULT_PAGES, PDF_PAGE_PREVIEW_CHARS, pdfText } from '../formats/pdf'
import { extension, readInput, resolveInput } from '../fs'
import { previewChars } from '../preview'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'

export const pdfCommand: CommandDef = {
  name: 'pdf',
  summary: 'Lê a camada de texto de um PDF página por página, sem nenhum processo de aplicativo.',
  usage: 'pdf read <file.pdf> [--page n | --range a-b] [--password <pw>] [--full | --max-chars n]',
  options: [
    { name: 'page', value: 'n', description: 'read: apenas esta página (base 1)' },
    {
      name: 'range',
      value: 'a-b',
      description: `read: intervalo de páginas (base 1; padrão: as primeiras ${PDF_DEFAULT_PAGES} páginas)`,
    },
    { name: 'password', value: 'pw', description: 'senha para um PDF criptografado' },
    {
      name: 'full',
      description: 'read: todas as páginas com o texto completo, em vez de uma prévia cortada',
    },
    {
      name: 'max-chars',
      value: 'n',
      description: `read: tamanho da prévia por página (padrão ${PDF_PAGE_PREVIEW_CHARS}); texto cortado termina em …(+n caracteres)`,
    },
  ],
  async run(args, ctx) {
    const [verb, file] = args.positionals
    if (verb !== 'read') {
      throw new CliError(EXIT.usage, 'esperado "pdf read <file.pdf>"', undefined, {
        reason: verb === undefined ? 'missing_argument' : 'invalid_argument',
        suggestion: 'execute `genoffice help pdf`',
      })
    }
    return read(file, args, ctx)
  },
}

type Args = Parameters<CommandDef['run']>[0]

async function read(
  file: string | undefined,
  args: Args,
  ctx: CommandContext,
): Promise<CommandResult> {
  const path = resolveInput(file, ctx)
  if (extension(path) !== 'pdf') {
    throw new CliError(EXIT.usage, `não é um PDF: ${path}`, undefined, {
      reason: 'unsupported',
      suggestion: 'use docs read, sheet read ou slides read para arquivos Office',
    })
  }
  const page = flagString(args, 'page')
  const range = flagString(args, 'range')
  if (page !== undefined && range !== undefined) {
    throw new CliError(EXIT.usage, '--page e --range são exclusivos', undefined, {
      reason: 'invalid_argument',
    })
  }
  const full = flagBool(args, 'full')
  const password = flagString(args, 'password')
  let windowed = false
  const select = (count: number): [number, number] => {
    if (page !== undefined) return parseSpan(page, count, '--page deve ser um número de página (base 1)')
    if (range !== undefined) return parseSpan(range, count, '--range deve ter o formato 3 ou 3-10')
    if (full) return [1, count]
    windowed = count > PDF_DEFAULT_PAGES
    return [1, Math.min(PDF_DEFAULT_PAGES, count)]
  }
  let doc
  try {
    doc = await pdfText(readInput(path), {
      password,
      maxChars: previewChars(args, PDF_PAGE_PREVIEW_CHARS),
      select,
    })
  } catch (err) {
    throw loadError(err, password)
  }
  const first = doc.pages_read[0]?.page
  const last = doc.pages_read.at(-1)?.page
  const detail: Record<string, unknown> = { ...doc }
  if (first !== undefined && last !== undefined) detail.range = `${first}-${last}`
  if (windowed && last !== undefined) {
    const next = `${last + 1}-${Math.min(last + PDF_DEFAULT_PAGES, doc.pages)}`
    detail.truncated = `páginas ${last + 1}-${doc.pages} não lidas; use --range ${next} para as próximas ${PDF_DEFAULT_PAGES}, --page n para uma, ou --full para todas`
  }
  const read = detail.range ? `, lidas ${detail.range}` : ''
  return { summary: `${basename(path)}: ${doc.pages} páginas${read}`, detail }
}

function parseSpan(spec: string, count: number, usage: string): [number, number] {
  const m = /^(\d+)(?:-(\d+))?$/.exec(spec)
  if (!m) throw new CliError(EXIT.usage, usage, undefined, { reason: 'invalid_argument' })
  const start = Number(m[1])
  const end = m[2] === undefined ? start : Number(m[2])
  if (start < 1 || start > end || end > count) {
    throw new CliError(
      EXIT.usage,
      `intervalo de páginas ${spec} fora dos limites (1-${count})`,
      { valid_range: [1, count] },
      { reason: 'out_of_range', suggestion: `use páginas entre 1 e ${count}` },
    )
  }
  return [start, end]
}

function loadError(err: unknown, password: string | undefined): unknown {
  if (!(err instanceof PdfLoadError)) return err
  if (err.code === 'password-required') {
    return password === undefined
      ? new CliError(EXIT.usage, 'o PDF está criptografado', undefined, {
          reason: 'missing_argument',
          suggestion: 'use --password <pw>',
        })
      : new CliError(EXIT.usage, 'senha incorreta para este PDF', undefined, {
          reason: 'invalid_argument',
          suggestion: 'verifique o valor de --password',
        })
  }
  return new CliError(EXIT.file, `não é um PDF legível (${err.code})`, undefined, {
    reason: 'unsupported',
  })
}
