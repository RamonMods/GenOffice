import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js'
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js'
import type { BatchCounts } from '../batch'
import { defaultRegistry, VERSION } from '../cli'
import type { OpFailure } from '../op-errors'
import type { CommandRegistry } from '../registry'
import type { JsonOk } from '../result'
import { DECK_TOOLS, runDeckTool } from './deck'
import { loadTypedSchemas, type TypedSchemas } from './op-schemas'
import { attachOutputs, defaultOut, materializeInputs, urlErrorOutcome } from './remote'
import {
  createContext,
  disposeContext,
  imageBlocks,
  runJson,
  runWithInline,
  type McpContext,
  type Outcome,
} from './run'
import { buildArgv, resolveTools, toolShape, type ResolvedTool } from './tools'

const ABOUT =
  'GenOffice: crie, leia, converta, edite e renderize documentos do Office localmente (docx, xlsx, pptx, pdf, md, html, csv). O aplicativo não precisa estar em execução; render, convert-to-pdf e create_pdf iniciam um processo oculto do GenOffice por alguns segundos.'
const WORKFLOW = [
  'Edição: leia o arquivo com a ferramenta *_read correspondente, escreva as ops com a referência de guide (ou os recursos genoffice://guide/*) e então *_apply. Uma op rejeitada informa seu índice e motivo; corrija essa op e reenvie o lote inteiro.',
  'Uma nova apresentação: um deck curto (até cerca de 5 slides, ou conteúdo concreto sem um briefing de design) é uma única chamada create_pptx com ops ou uma spec, seguida de slides_render. Decks mais longos ou sensíveis ao design que uma pessoa vai apresentar: deck_start (style sheet + outline), deck_page uma vez por página em ordem, deck_build, depois slides_render para conferir e slides_audit para a geometria, deck_replace para corrigir uma página. Edições em um deck existente: slides_read + slides_apply, mantendo o design original.',
]

/** What the client shows the model about this server before any tool is called. */
export const INSTRUCTIONS = [
  ABOUT,
  'Caminhos são absolutos ou relativos ao diretório de trabalho no qual o servidor foi iniciado. Apenas search, image e media enviam dados para fora da máquina, ao provedor configurado no GenOffice.',
  ...WORKFLOW,
  'Um arquivo aberto em uma aba do GenOffice não é gravado a menos que force esteja definido. Não chame open a menos que o usuário peça para ver o arquivo.',
].join('\n')

/** The http-mode variant: the client is on another machine, so files travel as URLs and result content. */
export function remoteInstructions(baseUrl: string): string {
  const base = baseUrl.replace(/\/$/, '')
  return [
    ABOUT,
    `Este servidor roda em outra máquina: os caminhos que você conhece não são visíveis para ele. Para trabalhar em um arquivo que você tem, envie-o primeiro (curl -T report.docx ${base}/files/ — a resposta traz sua url) e passe essa url onde quer que uma ferramenta aceite um arquivo; qualquer outra URL http(s) que o servidor consiga acessar também funciona. Caminhos relativos e pastas de deck ficam em um diretório privado de rascunho desta sessão.`,
    'Omita out: o arquivo que uma ferramenta grava volta no resultado como output_url (baixe com curl -o) e, quando pequeno, como um recurso incorporado com os bytes. Apenas search, image e media enviam dados ao provedor em nuvem configurado no GenOffice.',
    ...WORKFLOW,
  ].join('\n')
}

const GUIDES: { uri: string; name: string; argv: string[]; description: string }[] = [
  {
    uri: 'genoffice://guide/docs',
    name: 'Referência de ops do Word',
    argv: ['guide', 'docs'],
    description: 'todas as ops de docs_apply com seus campos e as regras de HTML restrito',
  },
  {
    uri: 'genoffice://guide/sheets',
    name: 'Referência de ops do Excel',
    argv: ['guide', 'sheets'],
    description: 'todas as ops da DSL de sheet_apply com seus campos',
  },
  {
    uri: 'genoffice://guide/slides',
    name: 'Referência de ops do PowerPoint',
    argv: ['guide', 'slides'],
    description: 'os grupos de ops e o vocabulário de slides_apply',
  },
  {
    uri: 'genoffice://guide/slides/design',
    name: 'Guia de design de decks',
    argv: ['guide', 'slides', 'design'],
    description: 'o fluxo de decks em etapas: style sheet, outline, uma página por vez, build, QC',
  },
  {
    uri: 'genoffice://guide/slides/spec',
    name: 'Formato de spec de decks',
    argv: ['guide', 'slides', 'spec'],
    description: 'o outline e a spec JSON de uma página que as ferramentas deck_* recebem',
  },
]

export interface ServerOptions {
  registry?: CommandRegistry
  /** advertise ops / cells / data as untyped `array | object` (GENOFFICE_MCP_COMPACT_SCHEMAS=1, --compact-schemas) */
  compactSchemas?: boolean
}

let typedSchemas: Promise<TypedSchemas> | undefined

export async function createMcpServer(
  ctx: McpContext,
  opts: ServerOptions = {},
): Promise<McpServer> {
  const registry = opts.registry ?? defaultRegistry()
  const remote = ctx.mode === 'http'
  const typed = opts.compactSchemas ? undefined : await (typedSchemas ??= loadTypedSchemas())
  const server = new McpServer(
    { name: 'genoffice', version: VERSION },
    { instructions: remote ? remoteInstructions(ctx.baseUrl ?? '') : INSTRUCTIONS },
  )

  for (const tool of resolveTools(registry)) {
    // there is no GenOffice window in front of a remote client
    if (remote && tool.localOnly) continue
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: toolShape(tool, ctx.mode, typed),
        annotations: {
          readOnlyHint: tool.readOnly === true,
          openWorldHint: tool.openWorld === true,
        },
      },
      async (args) => {
        const outcome = await runTool(tool, { ...(args as Record<string, unknown>) }, ctx)
        return toResult(outcome, { images: tool.images, plainText: tool.plainText, ctx })
      },
    )
  }

  for (const tool of DECK_TOOLS) {
    server.registerTool(
      tool.name,
      {
        description: tool.description,
        inputSchema: tool.shape,
        annotations: { readOnlyHint: tool.readOnly === true, openWorldHint: false },
      },
      async (args) =>
        toResult(await runDeckTool(tool, args as Record<string, unknown>, ctx), { ctx }),
    )
  }

  for (const g of GUIDES) {
    server.registerResource(
      g.name,
      g.uri,
      { description: g.description, mimeType: 'text/plain' },
      async (uri) => {
        const r = await runJson(g.argv, ctx)
        const text = r.ok ? r.ok.summary : `${r.error.message}`
        return { contents: [{ uri: uri.href, mimeType: 'text/plain', text }] }
      },
    )
  }
  return server
}

async function runTool(
  tool: ResolvedTool,
  args: Record<string, unknown>,
  ctx: McpContext,
): Promise<Outcome> {
  if (ctx.mode === 'http') {
    try {
      await materializeInputs(tool, args, ctx)
    } catch (err) {
      return urlErrorOutcome(tool, err)
    }
    defaultOut(tool, args, ctx)
  }
  const { argv, inline } = buildArgv(tool, args)
  return runWithInline(argv, inline, ctx)
}

function toResult(
  outcome: Outcome,
  opts: { images?: boolean; plainText?: boolean; ctx: McpContext },
): CallToolResult {
  if (outcome.error) {
    return { isError: true, content: [{ type: 'text', text: JSON.stringify(outcome.error) }] }
  }
  const ok = outcome.ok
  const outputs = opts.ctx.mode === 'http' ? attachOutputs(ok, opts.ctx) : []
  const content: CallToolResult['content'] = [
    ...(ok.status === 'partial' ? [{ type: 'text' as const, text: partialSummary(ok) }] : []),
    { type: 'text', text: opts.plainText ? ok.summary : JSON.stringify(ok) },
    ...outputs,
  ]
  if (opts.images) {
    const { images, omitted } = imageBlocks(ok)
    content.push(...images)
    if (omitted)
      content.push({
        type: 'text',
        text: `${omitted} página(s) adicional(is) gravada(s) em disco, mas não anexada(s); use render uma por vez`,
      })
  }
  return { content }
}

/** A partial batch is not an error, so the model needs the lost ops spelled out ahead of the envelope. */
function partialSummary(ok: JsonOk): string {
  const batch = ok.detail?.batch as BatchCounts | undefined
  const failures = (ok.detail?.failures as OpFailure[] | undefined) ?? []
  const head = batch
    ? `parcial: ${batch.applied} de ${batch.total} ops aplicadas, ${batch.failed} falharam${batch.skipped ? `, ${batch.skipped} ignoradas` : ''}`
    : `parcial: ${ok.summary}`
  const lines = failures.map((f) => `  op ${f.index} (${f.op}): ${f.error}`)
  return [head, ...lines].join('\n')
}

export interface ServeOptions extends ServerOptions {
  cwd: string
  env: NodeJS.ProcessEnv
  log: (message: string) => void
}

/** Serves on stdin/stdout until the client closes the transport. */
export async function serveStdio(opts: ServeOptions): Promise<void> {
  const ctx = createContext(opts)
  const transport = new StdioServerTransport()
  try {
    const server = await createMcpServer(ctx, opts)
    await new Promise<void>((resolve, reject) => {
      server.server.onclose = () => resolve()
      server.connect(transport).catch(reject)
    })
  } finally {
    disposeContext(ctx)
  }
}
