import { z } from 'zod'
import type { CommandRegistry, OptionDef } from '../registry'
import type { McpMode } from './run'
import { READ_UNITS } from '../length-units'
import type { JsonSchema } from '../op-catalog'
import type { TypedKey, TypedSchemas } from './op-schemas'

/**
 * How a tool parameter reaches the command: `inline` values are written to a
 * scratch file whose path is passed to the option, so an MCP client without a
 * filesystem can still send ops, specs and Markdown.
 */
export type ParamKind =
  'string' | 'boolean' | 'integer' | 'number' | 'inline-json' | 'inline-md' | 'inline-html'

export interface ParamSpec {
  /** parameter name in the tool schema */
  key: string
  /** the command option it becomes (`--<option>`); absent when key and option match */
  option?: string
  kind?: ParamKind
  /** closed set of string values; the schema becomes an enum */
  choices?: readonly string[]
  /** inline-json parameters: the typed schema advertised to clients (op-schemas.ts); runtime stays permissive */
  typed?: TypedKey
  description?: string
  required?: boolean
}

export interface PositionalSpec {
  key: string
  description: string
  optional?: boolean
}

export interface ToolSpec {
  name: string
  command: string
  verb?: string
  description: string
  positionals?: PositionalSpec[]
  /** argv appended to every call (`--type docx`) */
  fixed?: string[]
  /** a bare name takes its description from the command's OptionDef */
  options?: (string | ParamSpec)[]
  /** attach the PNGs listed in detail.files as image content */
  images?: boolean
  /** the text content is the result's summary, not the JSON envelope (guides) */
  plainText?: boolean
  readOnly?: boolean
  /** the call reaches the configured cloud provider (search, image, media) */
  openWorld?: boolean
  /** needs the GenOffice window in front of the client, so not registered in http mode */
  localOnly?: boolean
}

export interface ResolvedParam extends Required<Pick<ParamSpec, 'key' | 'option' | 'kind'>> {
  description: string
  required: boolean
  choices?: readonly string[]
  typed?: TypedKey
}

export interface ResolvedTool extends ToolSpec {
  params: ResolvedParam[]
}

const INTEGER_OPTIONS = new Set([
  'slide',
  'page',
  'max',
  'max-chars',
  'max-rows',
  'cols',
  'pad',
  'tile',
  'block',
])
const NUMBER_OPTIONS = new Set(['scale'])

const OPS_JSON = 'inline-json' as const

const READ_THEN_APPLY =
  'Leia primeiro o arquivo com a ferramenta *_read correspondente: as ops apontam para os ids e índices que ela lista. '
const GUI_OPEN =
  'Um arquivo aberto em uma aba do GenOffice é recusado sem force (o editor sobrescreveria a edição no próximo salvamento).'

export const TOOLS: ToolSpec[] = [
  {
    name: 'info',
    command: 'info',
    readOnly: true,
    description:
      'Metadados e um resumo da estrutura de um documento (docx, xlsx, pptx, pdf, md, html, csv): número de páginas ou slides, planilhas, quantidade de blocos, tamanho. Executa localmente.',
    positionals: [{ key: 'file', description: 'caminho do documento' }],
    options: ['password'],
  },
  {
    name: 'convert',
    command: 'convert',
    description:
      'Converta um documento para outro formato com os motores do GenOffice (pdf, docx, xlsx, pptx, md, html, csv; a descrição de `to` lista as rotas). Destinos que precisam de layout (para pdf, docx para html, html para docx) iniciam um processo oculto do GenOffice por 1-6 s. Por padrão, a saída usa o nome de entrada com a nova extensão.',
    positionals: [{ key: 'file', description: 'caminho do documento de origem' }],
    options: ['to', 'out', 'force', 'password', 'sheet'],
  },
  {
    name: 'create_docx',
    command: 'create',
    description:
      'Crie um novo documento do Word a partir de Markdown (GFM; títulos, listas, tabelas, imagens por caminho relativo, $formulas$) ou de um fragmento de HTML restrito (tags conforme guide docs). Informe exatamente um entre markdown, html ou from. Edições em um .docx existente devem usar docs_apply.',
    fixed: ['--type', 'docx'],
    options: [
      {
        key: 'markdown',
        option: 'from',
        kind: 'inline-md',
        description: 'o documento como texto Markdown',
      },
      {
        key: 'html',
        option: 'from',
        kind: 'inline-html',
        description: 'o documento como um fragmento de HTML restrito',
      },
      { key: 'from', description: 'caminho de um arquivo .md ou .html existente para usar como base' },
      { key: 'out', description: 'caminho de saída .docx (required)', required: true },
      'force',
    ],
  },
  {
    name: 'create_xlsx',
    command: 'create',
    description:
      'Crie uma nova planilha a partir de dados: uma matriz 2D de valores de células, { "sheets": [{ "name", "rows" }] } para várias planilhas, ou um arquivo .csv/.json. Strings que começam com "=" são fórmulas. Informe exatamente um entre data ou from. Formatação, gráficos e mais planilhas depois: sheet_apply.',
    fixed: ['--type', 'xlsx'],
    options: [
      {
        key: 'data',
        option: 'from',
        kind: OPS_JSON,
        typed: 'xlsx-data',
        description: 'linhas como uma matriz 2D, ou { "sheets": [{ "name", "rows" }] }',
      },
      { key: 'from', description: 'caminho de um arquivo .csv ou .json para usar como base' },
      'header',
      'decimal',
      { key: 'out', description: 'caminho de saída .xlsx (required)', required: true },
      'force',
    ],
  },
  {
    name: 'create_pdf',
    command: 'create',
    description:
      'Imprima um documento (md, html, docx, xlsx, pptx) em um novo PDF com o renderizador do GenOffice, em um processo oculto do GenOffice.',
    fixed: ['--type', 'pdf'],
    options: [
      { key: 'from', description: 'caminho do documento a imprimir', required: true },
      { key: 'out', description: 'caminho de saída .pdf (required)', required: true },
      'force',
    ],
  },
  {
    name: 'create_pptx',
    command: 'create',
    images: true,
    description:
      'Crie um .pptx em uma única chamada a partir de ops aplicadas a um deck em branco de um slide, ou a partir de uma spec de deck (páginas com texto, formas e imagens posicionadas em px em uma tela de 1280x720; guide slides spec). Este é o caminho para um deck curto (até cerca de 5 slides) e sempre que o usuário der conteúdo concreto e sem briefing de design: escreva as ops ou a spec diretamente e depois slides_render para conferir. Defina render=true para receber os PNGs dos slides na mesma chamada e audit=true para os achados de geometria; do contrário, execute slides_render e slides_audit depois. O fluxo em etapas deck_start / deck_page / deck_build é para apresentações mais longas ou sensíveis ao design que uma pessoa vai apresentar. Informe exatamente um entre ops ou spec.',
    fixed: ['--type', 'pptx'],
    options: [
      {
        key: 'ops',
        kind: OPS_JSON,
        typed: 'slides-op-names',
        description: 'array JSON de ops de slides, os mesmos objetos que slides_apply recebe (guide slides)',
      },
      {
        key: 'spec',
        kind: OPS_JSON,
        typed: 'object',
        description: 'um objeto de spec de deck { "pages": [...] } (guide slides spec)',
      },
      {
        key: 'spec_dir',
        option: 'spec',
        description:
          'caminho de um diretório de arquivos de spec de uma página, lidos em ordem de nome (as ferramentas deck_* gravam um)',
      },
      'outline',
      { key: 'out', description: 'caminho de saída .pptx (required)', required: true },
      'force',
      {
        key: 'render',
        kind: 'boolean',
        description:
          'renderize um PNG por slide após gravar e retorne-os como imagens (detail.previews); inicia um processo oculto do GenOffice por alguns segundos',
      },
      {
        key: 'audit',
        kind: 'boolean',
        description: 'inclua a auditoria de geometria do deck gerado em detail.audit',
      },
    ],
  },
  {
    name: 'docs_read',
    command: 'docs',
    verb: 'read',
    readOnly: true,
    description:
      'Estrutura e texto de um .docx em blocos indexados a partir de 0 (parágrafos, títulos, tabelas, imagens, gráficos) com uma prévia de 200 caracteres cada; os índices de bloco são o que as ops de docs_apply apontam. As flags adicionam comentários, alterações rastreadas, estilos, seções, campos, notas ou texto de cabeçalho/rodapé. Use full para o texto inteiro do bloco, range para restringir.',
    positionals: [{ key: 'file', description: 'caminho do .docx' }],
    options: [
      'range',
      'html',
      'full',
      'max-chars',
      'comments',
      'revisions',
      'styles',
      'header-footer',
      'sections',
      'fields',
      'notes',
      'properties',
    ],
  },
  {
    name: 'docs_apply',
    command: 'docs',
    verb: 'apply',
    description:
      'Edite um .docx com um lote de ops (a referência de ops: guide docs ou o recurso genoffice://guide/docs). ' +
      READ_THEN_APPLY +
      'Os índices de bloco dentro de um lote são dinâmicos: após uma inserção adicionar dois blocos, os índices seguintes se deslocam em dois. Atômico por padrão: uma op rejeitada deixa o arquivo intacto e retorna com seu índice e motivo; corrija essa op e reenvie o lote inteiro. dry_run valida sem gravar. ' +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'caminho do .docx' }],
    options: [
      {
        key: 'ops',
        kind: OPS_JSON,
        typed: 'docs-ops',
        description:
          'array JSON de ops, cada entrada com "op"; target é { "blockIndexes": [n] } ou { "nodeType": "docHeading"|"docParagraph"|"docListItem"|"image"|"table", "headingLevel"? }',
        required: true,
      },
      'track',
      'author',
      'dry-run',
      'best-effort',
      'stop-on-error',
      'out',
      'force',
    ],
  },
  {
    name: 'docs_check',
    command: 'docs',
    verb: 'check',
    readOnly: true,
    description:
      'Verificações de consistência em um .docx: campos não avaliados, referências quebradas, TOC desatualizado, imagens ausentes, gráficos vazios, títulos vazios; um achado por problema.',
    positionals: [{ key: 'file', description: 'caminho do .docx' }],
  },
  {
    name: 'merge',
    command: 'merge',
    description:
      'Preencha os placeholders {{key}} em um modelo .docx, .pptx ou .xlsx com valores de um objeto JSON e grave o resultado em out (objetos aninhados são achatados em chaves com ponto; espaços dentro das chaves são tolerados; um placeholder que ocupa a célula inteira no xlsx assume o tipo do valor). O resultado lista used_keys, unused_keys e unresolved_placeholders com sua localização (bloco / slide + elemento / planilha + célula) e motivo (no_key; split_placeholder quando as chaves estão onde uma correspondência de texto plano não consegue alcançá-las, como através de uma quebra de linha ou de um limite de bloco; unreachable_nested dentro de um grupo aninhado do pptx); strict transforma um placeholder não resolvido em erro. ' +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'caminho do modelo (.docx, .pptx ou .xlsx)' }],
    options: [
      {
        key: 'data',
        kind: OPS_JSON,
        typed: 'object',
        description: 'os valores como um objeto JSON, ex.: { "name": "Ada", "amount": 12 }',
        required: true,
      },
      { key: 'out', description: 'caminho de saída (required)', required: true },
      'force',
      'strict',
    ],
  },
  {
    name: 'pdf_read',
    command: 'pdf',
    verb: 'read',
    readOnly: true,
    description:
      'Camada de texto de um PDF página a página (page a partir de 1, width_pt/height_pt, texto limitado a 4000 caracteres com truncated: true) mais a contagem de páginas, o sinalizador de criptografia e os metadados title/author/producer; extraído localmente com pdfium, sem processo do aplicativo. Lê as páginas 1-20 por padrão e informa como obter o restante: page ou range restringem, full lê todas as páginas sem limite. Páginas digitalizadas retornam com texto vazio (convert para docx executa OCR).',
    positionals: [{ key: 'file', description: 'caminho do .pdf' }],
    options: ['page', 'range', 'password', 'full', 'max-chars'],
  },
  {
    name: 'sheet_read',
    command: 'sheet',
    verb: 'read',
    readOnly: true,
    description:
      'Células de uma planilha como valores brutos (0.25, números seriais de data) com fórmulas indexadas por endereço A1, mais recursos da planilha (painéis, filtro, gráficos, tabelas). Por padrão, usa a planilha ativa e até 500x100 células; o resultado informa quando houve truncamento. stats para contagens e a lista de planilhas, formats para estilos, larguras de coluna e alturas de linha, where para listar apenas células de um tipo.',
    positionals: [{ key: 'file', description: 'caminho do .xlsx' }],
    options: ['sheet', 'range', 'cols', 'max-rows', 'where', 'stats', 'formats'],
  },
  {
    name: 'sheet_apply',
    command: 'sheet',
    verb: 'apply',
    description:
      'Edite um .xlsx: cells define valores, fórmulas e estilos por endereço; ops executa a DSL da planilha (formatação, gráficos, imagens, tabelas, congelar painéis, filtros, formatação condicional, validação, hyperlinks, notas, configuração de página, proteção, nomes definidos, ordem das planilhas; guide sheets). Informe exatamente um entre cells ou ops. ' +
      READ_THEN_APPLY +
      'As ops executam em ordem, cada uma vendo o resultado anterior; ops estruturais (linhas, colunas, planilhas) precisam de um lote próprio. ' +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'caminho do .xlsx' }],
    options: [
      'sheet',
      {
        key: 'cells',
        kind: OPS_JSON,
        typed: 'cells',
        description:
          'array JSON de { "cell": "B2", "sheet"?, "value"? | "formula"?, "style"? }; style recebe os campos de formato da op format_range',
      },
      {
        key: 'ops',
        kind: OPS_JSON,
        typed: 'sheets-ops',
        description: 'array JSON de ops da DSL da planilha',
      },
      'dry-run',
      'best-effort',
      'stop-on-error',
      'out',
      'force',
    ],
  },
  {
    name: 'sheet_check',
    command: 'sheet',
    verb: 'check',
    readOnly: true,
    description:
      'Verificações de consistência em um .xlsx: erros de fórmula, fórmulas não avaliadas, referências de planilha ausentes, nomes definidos quebrados, referências de gráfico, estouro de número; um achado por problema.',
    positionals: [{ key: 'file', description: 'caminho do .xlsx' }],
  },
  {
    name: 'slides_read',
    command: 'slides',
    verb: 'read',
    readOnly: true,
    description:
      'Estrutura de um .pptx: slides s_<n> com seus elementos e_* (tipo, prévia de texto, geometria EMU), tabelas e, com full, o texto inteiro e as notas do palestrante. Os ids são o que as ops de slides_apply apontam; os ids de elementos editados ou criados mudam, então leia novamente antes de um segundo lote. layouts lista os layouts mestres para addSlideWithLayout.',
    positionals: [{ key: 'file', description: 'caminho do .pptx' }],
    options: ['slide', 'full', 'layouts', 'max-chars', { key: 'units', choices: READ_UNITS }],
  },
  {
    name: 'slides_apply',
    command: 'slides',
    verb: 'apply',
    description:
      'Edite um .pptx com um lote de ops (setText, addElement, addPicture, deleteElement, setTransform, addSlide ...; guide slides). ' +
      READ_THEN_APPLY +
      "As execuções de setText herdam o estilo da primeira execução do elemento, a menos que um campo seja definido. As ops executam em ordem; uma op posterior pode apontar para um slide adicionado por uma op anterior. Atômico por padrão. Após a edição, use slides_audit para transbordamento e slides_render para conferir. " +
      GUI_OPEN,
    positionals: [{ key: 'file', description: 'caminho do .pptx' }],
    options: [
      {
        key: 'ops',
        kind: OPS_JSON,
        typed: 'slides-ops',
        description: 'array JSON de ops; target é { "slide": <index|"s_n">, "el"?: "e_*" }',
        required: true,
      },
      'dry-run',
      'best-effort',
      'stop-on-error',
      'isolation',
      'out',
      'force',
    ],
  },
  {
    name: 'slides_audit',
    command: 'slides',
    verb: 'audit',
    readOnly: true,
    description:
      'Auditoria de layout de um .pptx: transbordamento de texto, elementos fora dos limites ou totalmente fora do slide, sobreposições e imagens esticadas (proporção da caixa diferente da proporção da origem recortada em mais de 5%) por slide, com os ids dos elementos e uma op setTransform sugerida para slides_apply. Larguras de glifo heurísticas; confirme com slides_render.',
    positionals: [{ key: 'file', description: 'caminho do .pptx' }],
    options: ['slide'],
  },
  {
    name: 'slides_render',
    command: 'slides',
    verb: 'render',
    readOnly: true,
    images: true,
    description:
      'Um PNG por slide (960x540 na escala 1 para 16:9), gravado em um diretório e retornado como imagens para você conferir as páginas. Inicia um processo oculto do GenOffice (alguns segundos). Use após montar ou editar um deck; slide renderiza uma página.',
    positionals: [{ key: 'file', description: 'caminho do .pptx' }],
    options: [
      { key: 'out', description: 'diretório para os PNGs (required)', required: true },
      'slide',
      'scale',
    ],
  },
  {
    name: 'slides_check',
    command: 'slides',
    verb: 'check',
    readOnly: true,
    description:
      'Verifique um arquivo do fluxo de decks em disco: um outline.json (gancho central e plano de páginas) ou uma spec de página, que é conferida em relação ao outline.json e ao style.md ao lado ou um nível acima. As ferramentas deck_* fazem isso por você; use diretamente quando os arquivos foram gravados por outros meios.',
    positionals: [{ key: 'file', description: 'caminho de outline.json ou de um .json de página' }],
    options: ['outline', 'page'],
  },
  {
    name: 'slides_replace',
    command: 'slides',
    verb: 'replace',
    description:
      'Reconstrua um slide de um .pptx a partir de um arquivo de spec de uma página em disco (um deck construído pelo fluxo em etapas); os outros slides mantêm seus ids e conteúdo. deck_replace faz o mesmo a partir de uma página inline.',
    positionals: [{ key: 'file', description: 'caminho do .pptx' }],
    options: [
      { key: 'slide', kind: 'integer', description: 'slide a reconstruir, indexado a partir de 0', required: true },
      { key: 'spec', description: 'caminho do arquivo de spec de uma página', required: true },
      'outline',
      'out',
      'force',
    ],
  },
  {
    name: 'render',
    command: 'render',
    readOnly: true,
    images: true,
    description:
      'Um PNG por página de um documento (docx, xlsx, pptx, pdf, md, html) conforme o renderizador do GenOffice a dispõe, gravado em um diretório e retornado como imagens. Inicia um processo oculto do GenOffice. page renderiza uma página; grid adiciona uma folha de contato com todas as páginas.',
    positionals: [{ key: 'file', description: 'caminho do documento' }],
    options: [
      { key: 'out', description: 'diretório para os PNGs (required)', required: true },
      'page',
      'scale',
      'el',
      'pad',
      'grid',
      'cols',
      'tile',
    ],
  },
  {
    name: 'guide',
    command: 'guide',
    readOnly: true,
    plainText: true,
    description:
      'A referência de ops que um agente lê antes de escrever ops: guide("docs"), guide("sheets") ou guide("slides") listam os grupos; topic restringe a um grupo ou a uma op. guide("slides", "design") é o fluxo de decks em etapas e guide("slides", "spec") o formato de outline e spec de página (deck_start retorna ambos). Também disponível como recursos genoffice://guide/*.',
    positionals: [
      { key: 'domain', description: 'slides | docs | sheets' },
      {
        key: 'topic',
        description: 'um grupo, um nome de op ou, para slides: design | spec',
        optional: true,
      },
    ],
    options: ['index'],
  },
  {
    name: 'capabilities',
    command: 'capabilities',
    readOnly: true,
    description:
      'Quais recursos em nuvem estão configurados no GenOffice: busca na web, busca de imagens, geração de imagens, compreensão de mídia. Verifique uma vez antes de planejar fotos ou fatos em tempo real; todo o resto executa localmente.',
  },
  {
    name: 'search',
    command: 'search',
    openWorld: true,
    readOnly: true,
    description:
      'Busca na web, ou busca de imagens com images=true, pelo provedor configurado no GenOffice (a consulta sai da máquina). Os resultados de imagem trazem imageUrl, largura e altura para addPicture / insert_image.',
    positionals: [{ key: 'query', description: 'a consulta de busca' }],
    options: ['images', 'max'],
  },
  {
    name: 'image',
    command: 'image',
    openWorld: true,
    description:
      'Gere uma imagem a partir de um prompt com o provedor configurado no GenOffice e salve-a; o resultado informa o formato real. Imagens ref orientam edições (remoção de fundo, upscale).',
    positionals: [{ key: 'prompt', description: 'o que desenhar' }],
    options: ['out', 'aspect', 'size', 'ref', 'model', 'force'],
  },
  {
    name: 'media',
    command: 'media',
    openWorld: true,
    readOnly: true,
    description:
      'Descreva ou extraia de um arquivo de imagem, áudio ou vídeo (ou URL) com o provedor configurado; o arquivo sai da máquina.',
    positionals: [{ key: 'source', description: 'caminho ou URL da mídia' }],
    options: ['ask'],
  },
  {
    name: 'open',
    command: 'open',
    localOnly: true,
    description:
      'Abra um documento no aplicativo GenOffice para o usuário (inicia o aplicativo se necessário), selecionando opcionalmente um slide, elemento, bloco, intervalo ou página. Apenas quando o usuário pedir para ver o arquivo: uma aba aberta faz com que chamadas *_apply posteriores se recusem a gravar.',
    positionals: [{ key: 'file', description: 'caminho do documento' }],
    options: ['slide', 'el', 'block', 'range', 'sheet', 'page'],
  },
  {
    name: 'selection',
    command: 'selection',
    localOnly: true,
    readOnly: true,
    description:
      "O que o usuário tem selecionado no momento no editor do GenOffice que mostra este arquivo: ids de slide + elemento, um intervalo de blocos com seu texto, um intervalo de planilha ou uma página de pdf. A própria indicação do usuário para 'este' / 'aqui'; requer o arquivo aberto no aplicativo.",
    positionals: [{ key: 'file', description: 'caminho do documento aberto no GenOffice' }],
  },
]

export function resolveTools(registry: CommandRegistry): ResolvedTool[] {
  return TOOLS.map((tool) => resolveTool(tool, registry))
}

export function resolveTool(tool: ToolSpec, registry: CommandRegistry): ResolvedTool {
  const def = registry.get(tool.command)
  if (!def) throw new Error(`mcp tool ${tool.name}: unknown command ${tool.command}`)
  const defs = new Map((def.options ?? []).map((o) => [o.name, o]))
  const params = (tool.options ?? []).map((entry): ResolvedParam => {
    const spec: ParamSpec = typeof entry === 'string' ? { key: entry } : entry
    const option = spec.option ?? spec.key
    const opt = defs.get(option)
    if (!opt) throw new Error(`mcp tool ${tool.name}: ${tool.command} has no --${option}`)
    return {
      key: spec.key.replace(/-/g, '_'),
      option,
      kind: spec.kind ?? defaultKind(opt),
      description: spec.description ?? stripVerbPrefix(opt.description),
      required: spec.required ?? false,
      ...(spec.choices ? { choices: spec.choices } : {}),
      ...(spec.typed ? { typed: spec.typed } : {}),
    }
  })
  return { ...tool, params }
}

function defaultKind(opt: OptionDef): ParamKind {
  if (!opt.value) return 'boolean'
  if (INTEGER_OPTIONS.has(opt.name)) return 'integer'
  if (NUMBER_OPTIONS.has(opt.name)) return 'number'
  return 'string'
}

/** Option descriptions start with the verb or type they apply to ("read: ...", "pptx --spec <dir>: ..."); the tool already says. */
export function stripVerbPrefix(description: string): string {
  const m = /^([a-z][\w .<>/-]{0,28}): /.exec(description)
  return m ? description.slice(m[0].length) : description
}

export type ZodShape = Record<string, z.ZodTypeAny>

/** Parameters that name a document on disk; a remote client passes URLs here instead. */
export const PATH_KEYS = new Set(['file', 'from', 'outline'])

const REMOTE_PATH_NOTE = '; ou uma URL http(s), como a retornada por POST /files para um envio'
const REMOTE_OUT_NOTE = '; omita-o e o arquivo volta no resultado como uma URL de download'

/**
 * `typed` carries the op schemas to advertise; without it (compact mode) inline JSON stays
 * `array | object`. The typed schema is metadata on a permissive zod type, so the SDK
 * accepts any array and the CLI's own op validation produces the structured error.
 */
export function toolShape(
  tool: ResolvedTool,
  mode: McpMode = 'stdio',
  typed?: TypedSchemas,
): ZodShape {
  const remote = mode === 'http'
  const shape: ZodShape = {}
  for (const p of tool.positionals ?? []) {
    const note = remote && PATH_KEYS.has(p.key) ? REMOTE_PATH_NOTE : ''
    const s = z.string().describe(p.description + note)
    shape[p.key] = p.optional ? s.optional() : s
  }
  for (const p of tool.params) {
    let description = p.description
    let required = p.required
    if (remote && p.key === 'out') {
      description = description.replace(/ \(required\)$/, '') + REMOTE_OUT_NOTE
      required = false
    } else if (remote && p.kind === 'string' && PATH_KEYS.has(p.option)) {
      description += REMOTE_PATH_NOTE
    }
    const typedSchema = typed && p.typed ? typed[p.typed] : undefined
    const base = p.choices
      ? z.enum(p.choices)
      : typedSchema
        ? inlineJsonSchema(typedSchema)
        : kindSchema(p.kind)
    const s = base.describe(description)
    shape[p.key] = required ? s : s.optional()
  }
  return shape
}

function kindSchema(kind: ParamKind): z.ZodTypeAny {
  switch (kind) {
    case 'boolean':
      return z.boolean()
    case 'integer':
      return z.number().int().nonnegative()
    case 'number':
      return z.number()
    case 'inline-json':
      return z.union([z.array(z.unknown()), z.record(z.string(), z.unknown())])
    default:
      return z.string()
  }
}

/** The zod type matches the JSON schema's outer shape; `.meta()` merges the typed body into the advertised schema. */
function inlineJsonSchema(schema: JsonSchema): z.ZodTypeAny {
  if (schema.type === 'array') return z.array(z.unknown()).meta(schema)
  if (schema.type === 'object') return z.record(z.string(), z.unknown()).meta(schema)
  return z.union([z.array(z.unknown()), z.record(z.string(), z.unknown())]).meta(schema)
}

export interface InlineFile {
  option: string
  ext: string
  text: string
}

export interface BuiltCall {
  /** argv without the trailing inline-file paths; `inline[i]` is appended as `--<option> <path>` once written */
  argv: string[]
  inline: InlineFile[]
}

/** The command line a tool call becomes; inline parameters are returned separately so the caller can place them on disk. */
export function buildArgv(tool: ResolvedTool, args: Record<string, unknown>): BuiltCall {
  const argv: string[] = [tool.command, ...(tool.verb ? [tool.verb] : []), ...(tool.fixed ?? [])]
  for (const p of tool.positionals ?? []) {
    const v = args[p.key]
    if (v === undefined || v === null || v === '') {
      if (!p.optional) throw new Error(`missing ${p.key}`)
      continue
    }
    argv.push(String(v))
  }
  const inline: InlineFile[] = []
  for (const p of tool.params) {
    const v = args[p.key]
    if (v === undefined || v === null) continue
    if (p.kind === 'boolean') {
      if (v === true) argv.push(`--${p.option}`)
      continue
    }
    if (p.kind === 'inline-json') {
      inline.push({ option: p.option, ext: '.json', text: JSON.stringify(v) })
      continue
    }
    if (p.kind === 'inline-md' || p.kind === 'inline-html') {
      inline.push({
        option: p.option,
        ext: p.kind === 'inline-md' ? '.md' : '.html',
        text: String(v),
      })
      continue
    }
    argv.push(`--${p.option}`, String(v))
  }
  return { argv, inline }
}
