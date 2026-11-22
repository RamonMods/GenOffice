import type { AgentToolDef } from './types'
import type { AgentSkill, ExecutedToolCall } from './skill'
import { kbCiteHref } from './knowledge-base-skill'

/**
 * The bundled `genoffice` command line as an agent skill: the same engine the
 * CLI, `genoffice mcp` and the app's own MCP server expose, driven here by the
 * in-app loop. The host owns the one dangerous operation (spawning the CLI):
 * it fixes the executable, confines reads and writes to the roots it hands
 * out, and stages the files the model writes into a per-conversation work
 * directory. This file only shapes the tools and the workflow prompt.
 */

export const RUN_GENOFFICE_TOOL = 'run_genoffice'
export const LOAD_GUIDE_TOOL = 'load_guide'

export const GUIDE_DOMAINS = ['docs', 'sheets', 'slides'] as const
export type GuideDomain = (typeof GUIDE_DOMAINS)[number]

/** one file the model writes before a command runs (relative to the work dir) */
export interface CliStagedFile {
  name: string
  content: string
}

export interface CliRunRequest {
  args: string[]
  files?: CliStagedFile[]
}

export interface CliRunResult {
  ok: boolean
  /** process exit code; -1 when the CLI never ran */
  code: number
  /** the CLI's `--json` payload when it printed one */
  json?: unknown
  stdout: string
  stderr: string
}

export interface CliContext {
  /** scratch directory for ops, specs and staged inputs; relative paths resolve here */
  workDir: string
  /** where finished documents go, the folder the home screen lists */
  outDir: string
}

export interface CliSkillDeps {
  /** null until the host has resolved the directories for this conversation */
  context(): CliContext | null
  run(request: CliRunRequest): Promise<CliRunResult>
}

export interface CliSkill extends AgentSkill {
  /** absolute paths the CLI reported as output_path during the session */
  createdPaths(): readonly string[]
}

const MAX_OUTPUT_CHARS = 24_000

const CLI_SYSTEM_PROMPT = `## Criando documentos com genoffice
run_genoffice executa a linha de comando genoffice incluída (sempre com --json) e
retorna seu resultado em JSON. Use-o sempre que o usuário quiser um arquivo real criado a partir
do material dele: um relatório no Word com títulos e tabelas, uma planilha com
fórmulas e gráficos dinâmicos, uma apresentação de slides. Texto simples que o
usuário só quer salvar pode ir por create_document; qualquer coisa com estrutura,
fórmulas, vários arquivos como fonte, ou um resultado .xlsx / .pptx vai por run_genoffice.
- Caminhos: o contexto informa um diretório de trabalho e um diretório de saída. Escreva
  ops, specs, markdown e outras entradas com o argumento files (nomes relativos
  ao diretório de trabalho), referencie-os por esse nome relativo em args e coloque
  todo documento pronto no diretório de saída com --out, usando um nome de arquivo
  curto e descritivo no idioma do usuário. Arquivos de origem são os caminhos da base
  de conhecimento no contexto; leia-os com "docs read", "sheet read" ou
  "pdf read" (adicione --full ou --range quando a prévia não for suficiente).
- Um arquivo dentro do diretório de saída pode ser editado no lugar ("apply" /
  "replace" sem --out; o original é copiado antes). Um arquivo da base de
  conhecimento em qualquer outro lugar é somente leitura: edite-o com --out para um
  arquivo novo no diretório de saída. --force é recusado; escolha um nome que ainda
  não exista.
- "<command> --help" mostra as opções de um comando quando o guia não mostra.
- Antes de escrever ops ou uma spec de apresentação, chame load_guide para esse domínio uma vez
  (docs, sheets ou slides; um grupo ou nome de opção reduz o escopo). Siga o
  pipeline de apresentação em etapas descrito no guia de design de slides: style.md, outline.json,
  um arquivo de página por vez, verificado com "slides check", e então "create --type
  pptx --spec pages --outline outline.json --out ... --audit". Uma página de
  apresentação que precisa de uma tabela ou gráfico nativo é criada com ops ("create
  --type pptx --ops" com addTable / addChart): uma spec de página não tem elemento
  de tabela, e uma grade desenhada com formas não é uma tabela.
- Word: "create --type docx --from report.md" cria um documento a partir de Markdown
  (títulos, listas, tabelas, negrito); depois "docs apply --ops ... --out" para estilos
  de tabela, campos ou configuração de página. Excel: "create --type xlsx --from table.json"
  e então "sheet apply --ops" para fórmulas, formatações, regras condicionais e gráficos;
  escreva fórmulas como strings que começam com "=".
- Leia o JSON: status "ok" com output_path significa que o arquivo existe; um erro
  traz error, message e suggestion, então corrija a entrada e reenvie o
  lote inteiro. Após uma criação, "render <file> --out shots" gera um PNG por página quando
  você precisa conferir o layout.
- Relate o resultado de forma breve e vincule cada arquivo pronto como
  [nome do arquivo](link de citação), copiando exatamente o link de citação retornado pela ferramenta.
  Só afirme que um arquivo foi criado quando run_genoffice retornou status "ok".`

// save-only phrasing, as in create-document-skill: "created"/"generated" also describe corpus facts
const CLAIMS_BUILT =
  /\bI(?:'ve| have)? (?:saved|exported|built|written) (?:it|this|that|the|a|an|your)\b|\b(?:has been|have been|is now) (?:saved|exported|built)\b|\bsaved (?:it |the (?:file|document|workbook|deck) )?(?:as|to|under)\b/i

function clip(text: string): string {
  return text.length > MAX_OUTPUT_CHARS
    ? `${text.slice(0, MAX_OUTPUT_CHARS)}\n…(+${text.length - MAX_OUTPUT_CHARS} caracteres)`
    : text
}

function stringList(value: unknown): string[] | null {
  if (!Array.isArray(value) || value.length === 0) return null
  const out: string[] = []
  for (const item of value) {
    if (typeof item !== 'string') return null
    out.push(item)
  }
  return out
}

function stagedFiles(value: unknown): CliStagedFile[] | null {
  if (value === undefined || value === null) return []
  if (!Array.isArray(value)) return null
  const out: CliStagedFile[] = []
  for (const item of value) {
    if (!item || typeof item !== 'object') return null
    const { name, content } = item as Record<string, unknown>
    if (typeof name !== 'string' || !name.trim() || typeof content !== 'string') return null
    out.push({ name, content })
  }
  return out
}

/** the CLI's own one-line summary, when the payload has one */
function summaryOf(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined
  const { summary, message } = json as { summary?: unknown; message?: unknown }
  if (typeof summary === 'string') return summary
  if (typeof message === 'string') return message
  return undefined
}

function outputPathOf(json: unknown): string | undefined {
  if (!json || typeof json !== 'object') return undefined
  const { status, output_path: out } = json as { status?: unknown; output_path?: unknown }
  return status === 'ok' && typeof out === 'string' && out ? out : undefined
}

export function createCliSkill(deps: CliSkillDeps): CliSkill {
  const created: string[] = []

  const runTool: AgentToolDef = {
    name: RUN_GENOFFICE_TOOL,
    description:
      'Executa a linha de comando genoffice incluída e retorna seu resultado em JSON. args é o comando ' +
      'e seus argumentos (--json é adicionado), ex.: ["create","--type","docx","--from","report.md","--out","<output dir>/Report.docx"]. ' +
      'files grava as entradas no diretório de trabalho antes da execução (nomes relativos), para que ops, specs e markdown ' +
      'não precisem ser escapados na linha de comando. Documentos prontos vão para o diretório de saída com --out.',
    inputSchema: {
      type: 'object',
      properties: {
        args: {
          type: 'array',
          items: { type: 'string' },
          description: 'o comando genoffice e seus argumentos, uma string por item',
        },
        files: {
          type: 'array',
          description: 'arquivos a gravar no diretório de trabalho antes da execução do comando',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'caminho relativo dentro do diretório de trabalho' },
              content: { type: 'string', description: 'o arquivo inteiro como texto' },
            },
            required: ['name', 'content'],
          },
        },
      },
      required: ['args'],
    },
  }

  const guideTool: AgentToolDef = {
    name: LOAD_GUIDE_TOOL,
    description:
      'As referências de operações e os guias de design de um domínio (o que "genoffice guide" imprime). ' +
      'Chame antes de escrever ops ou uma spec de apresentação; topic reduz o escopo a um grupo, uma operação, "design" ou "spec" (slides).',
    inputSchema: {
      type: 'object',
      properties: {
        domain: { type: 'string', enum: [...GUIDE_DOMAINS] },
        topic: {
          type: 'string',
          description: 'opcional: um nome de grupo, um nome de operação ou, para slides, "design" / "spec"',
        },
      },
      required: ['domain'],
    },
  }

  const execute = async (request: CliRunRequest, summaryLabel: string, plain = false) => {
    const result = await deps.run(request)
    const out = outputPathOf(result.json)
    if (out) created.push(out)
    // a guide is prose: hand it over as text, not as one JSON string with escaped newlines
    const body =
      plain && result.ok
        ? (summaryOf(result.json) ?? result.stdout)
        : result.json
          ? JSON.stringify(result.json, null, 1)
          : result.stdout
    const lines = [`saída ${result.code}`]
    const summary = summaryOf(result.json)
    if (summary && !plain) lines.push(`resumo: ${summary}`)
    if (out) lines.push(`output_path: ${out}`, `cite: ${kbCiteHref(out)}`)
    lines.push(clip(body))
    if (result.stderr.trim()) lines.push(`[stderr]\n${clip(result.stderr)}`)
    return {
      output: lines.join('\n'),
      isError: !result.ok,
      mutated: false,
      summary: summaryLabel,
    }
  }

  return {
    id: 'genoffice-cli',
    systemPrompt: CLI_SYSTEM_PROMPT,
    tools: [runTool, guideTool],
    createdPaths: () => created,
    buildContext: () => {
      const ctx = deps.context()
      if (!ctx) return ''
      return `diretórios do genoffice:\n  diretório de trabalho: ${ctx.workDir}\n  diretório de saída: ${ctx.outDir}`
    },
    executeTool: async (call) => {
      if (call.name === LOAD_GUIDE_TOOL) {
        const domain = call.input.domain
        if (typeof domain !== 'string' || !(GUIDE_DOMAINS as readonly string[]).includes(domain)) {
          return {
            output: `domain deve ser um de ${GUIDE_DOMAINS.join(', ')}`,
            isError: true,
            summary: 'domínio inválido',
          }
        }
        const topic = typeof call.input.topic === 'string' ? call.input.topic.trim() : ''
        const args = ['guide', domain, ...(topic ? [topic] : [])]
        return execute({ args }, `guide ${domain}${topic ? ` ${topic}` : ''}`, true)
      }
      if (call.name !== RUN_GENOFFICE_TOOL) {
        return { output: `ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
      }
      const args = stringList(call.input.args)
      if (!args) {
        return {
          output: 'args deve ser um array não vazio de strings',
          isError: true,
          summary: 'args inválidos',
        }
      }
      const files = stagedFiles(call.input.files)
      if (!files) {
        return {
          output: 'files deve ser um array de { name, content } com campos do tipo string',
          isError: true,
          summary: 'files inválidos',
        }
      }
      try {
        return await execute({ args, files }, `genoffice ${args.slice(0, 3).join(' ')}`)
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        return {
          output: `não foi possível executar o genoffice: ${message}`,
          isError: true,
          summary: 'falha na CLI',
        }
      }
    },
    verifyResponse: (finalText, executed: readonly ExecutedToolCall[]) => {
      const built = executed.some((c) => c.name === RUN_GENOFFICE_TOOL && c.ok)
      if (built || !CLAIMS_BUILT.test(finalText)) return null
      return (
        'A resposta diz que um arquivo foi criado ou salvo, mas nenhuma chamada run_genoffice teve sucesso nesta execução. ' +
        'Diga claramente que nada foi gravado ou crie o arquivo primeiro.'
      )
    },
  }
}
