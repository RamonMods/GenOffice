import { z } from 'zod'
import type { AgentToolCall, AgentToolDef } from '@genoffice/agent-core'
import {
  copyTargetBounds,
  describeOperationErrors,
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  columnLabel,
  parseRange,
  rangeCellCount,
  formatAddress,
  type RangeBounds,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type {
  ApplyOutcome,
  CellFormatState,
  CellScalar,
  ChangePlan,
} from '@genoffice/xlsx-gateway/domain/workbook.types'
import { t } from '../i18n/locale'
import { formatRangeAggregate, type RangeAggregate } from './aggregate'
import { guideCatalogSummary, loadGuides } from './guides'
import { modelCellValueAt, NO_REDACTIONS, type RedactionIndex } from './redact'
import { redactGuardForOps } from './redact-guard'

/**
 * The workbook DSL as an AgentSkill tool set: read-only context/reader tools
 * and one propose tool. Mirrors the docx skill's read-before-write discipline
 * (get_document_context / read_blocks / replace_blocks), but the mutating
 * tool never touches the workbook directly — it only computes a ChangePlan
 * and hands it to the SAME plan/apply path the manual flow uses, which now
 * auto-applies immediately (undo via ⌘Z / inline button covers everything;
 * the preview card only remains as a manual fallback when apply fails).
 */

/** raw shape the model sends for one operation; validated against workbookOperationSchema */
export type ProposedOperation = Record<string, unknown>

export interface SheetRef {
  readonly id: string
  readonly name: string
  /** Data extent (from the xlsx dimension or known cells); may drift slightly
   * after structural changes within the session */
  readonly rows?: number
  readonly columns?: number
  /** Worksheet XML above the gateway save-patch cap: readable, never editable. */
  readonly readOnlyOversized?: boolean
}

export interface ChartRef {
  readonly path: string
  readonly title: string
  readonly types: string
  readonly sheetId: string
}

/**
 * The user's selection captured when they sent the message. The grid selection
 * is live: users keep clicking around while the AI works, so reading it at
 * tool-call time would silently retarget "this column" mid-run.
 */
export interface FrozenSelection {
  /** A1 notation on the sheet it was taken from, clamped to the data extent so
   *  a whole-column click is not reported as a million rows */
  readonly a1: string
  readonly sheetId: string
  /** header names when the selection covers whole columns */
  readonly columns?: readonly string[]
}

export interface ActiveSheetInfo {
  readonly mode: 'demo' | 'lazy' | 'none'
  /** lazy mode only: the file is too large for a full load — cached values
   * stream in per viewport, and the live formula engine never sees the whole
   * data (formula writes over the file's sheets are gated) */
  readonly streaming?: boolean | undefined
  readonly sheetId: string
  readonly sheetName: string
  /** demo mode only: current revision, needed for the CAS-checked plan() call */
  readonly revision?: number
  /** non-empty cell addresses known to the caller without an extra read */
  readonly knownAddresses: readonly string[]
  /** lazy mode only: the viewport-backed range currently present in Univer */
  readonly loadedRange?: string | undefined
  /** every sheet in the workbook, active one included */
  readonly sheets: readonly SheetRef[]
  /** the selection to interpret "this column / these rows" against, in A1
   * notation (sheet-qualified when it is not the active sheet) */
  readonly selection?: string | undefined
  /** the selection above is the send-time snapshot rather than a live read */
  readonly selectionFrozen?: boolean | undefined
  /** header names of the columns the selection covers, when it covers whole
   * ones — what the user means by "this column" */
  readonly selectionColumns?: readonly string[] | undefined
  /** merged ranges on the active sheet (A1 notation) */
  readonly merges?: readonly string[] | undefined
  /** charts in the workbook (imported files only) */
  readonly charts?: readonly ChartRef[] | undefined
}

export interface FindCellsOptions {
  /** empty only when errorsOnly is set */
  readonly query: string
  readonly regex: boolean
  readonly lookIn: 'values' | 'formulas' | 'both'
  readonly errorsOnly: boolean
  readonly sheetId?: string | undefined
  readonly maxResults: number
}

export interface FindCellsMatch {
  readonly sheetName: string
  readonly address: string
  readonly value: CellScalar
  readonly formula?: string | undefined
}

export interface FindCellsOutcome {
  readonly matches: readonly FindCellsMatch[]
  /** the scan stopped at maxResults or the scan budget; more matches may exist */
  readonly truncated: boolean
  /** sheets whose background indexing hasn't finished — results there may be partial */
  readonly incompleteSheets: readonly string[]
  readonly error?: string
}

export interface SelectRangeOutcome {
  readonly ok: boolean
  readonly sheetName?: string
  readonly error?: string
}

export interface TraceCellSample {
  readonly address: string
  readonly value: CellScalar
  readonly formula?: string | undefined
}

export interface TraceRefInfo {
  /** display label like "B2:B9" or "Data!A1" */
  readonly label: string
  readonly cellCount: number
  readonly samples: readonly TraceCellSample[]
  /** some sampled value is a formula error (#REF!, #DIV/0!, …) */
  readonly hasError: boolean
  /** qualifier didn't match any sheet — external workbook or unresolvable */
  readonly external?: boolean
}

export interface TracePrecedentsOutcome {
  readonly formula?: string | undefined
  readonly value?: CellScalar
  readonly refs: readonly TraceRefInfo[]
  /** the formula had more references than the tracer reports */
  readonly truncatedRefs?: boolean
  /** formula also uses defined names / identifiers the tracer cannot resolve */
  readonly usesNames?: boolean
  readonly error?: string
}

export interface TraceDependentInfo {
  readonly sheetName: string
  readonly address: string
  readonly formula: string
  readonly value: CellScalar
}

export interface TraceDependentsOutcome {
  readonly dependents: readonly TraceDependentInfo[]
  readonly truncated: boolean
  readonly incompleteSheets: readonly string[]
  readonly error?: string
}

/** every file type create_document can produce */
export type CreateDocumentFileType = 'xlsx' | 'csv' | 'docx' | 'pdf' | 'md' | 'html'

/** create_document request handed to the App: xlsx/csv name a worksheet to
 * export; docx/pdf/md/html carry AI-authored content (routed to the docs flow).
 * Members keep singleton discriminants so the type narrows properly. */
export type CreateDocumentToolRequest =
  | { type: 'xlsx'; sheetId?: string | undefined; title?: string | undefined }
  | { type: 'csv'; sheetId?: string | undefined; title?: string | undefined }
  | { type: 'docx'; title: string; content: string }
  | { type: 'pdf'; title: string; content: string }
  | { type: 'md'; title: string; content: string }
  | { type: 'html'; title: string; content: string }

export type CreateDocumentToolOutcome =
  | {
      ok: true
      /** final file name including extension (title may default to the sheet name) */
      name: string
      /** absolute path when the file was written directly (docx opens a tab that saves itself) */
      path?: string
      /** xlsx/csv: the exported worksheet's name */
      sheetName?: string
      /** xlsx/csv: the sheet holds formulas — the file keeps computed values only */
      hadFormulas?: boolean
    }
  | { ok: false; error: string }

export interface SheetsSkillDeps {
  getActiveSheetInfo(): ActiveSheetInfo
  /** Ensure a lazy workbook range is present in Univer before reading it. */
  ensureRangeLoaded?(range: RangeBounds, sheetId?: string): boolean | Promise<boolean>
  /** cell values/formulas; reads the active sheet unless sheetId targets another */
  readCells(
    addresses: readonly string[],
    sheetId?: string,
  ): Record<string, { value: CellScalar; formula?: string }>
  /** per-cell explicit formatting; cells with no explicit format are omitted */
  readFormats(addresses: readonly string[], sheetId?: string): Record<string, CellFormatState>
  /** formatted report of a sheet's feature state (filters, CF, DV, names, visuals, …) */
  readSheetFeatures(sheetId?: string): string
  /** workbook-wide value/formula search (ai/workbook-search.ts) */
  findCells(options: FindCellsOptions): FindCellsOutcome | Promise<FindCellsOutcome>
  /** activate a sheet, select a range, and scroll it into view */
  selectRange(
    sheetId: string | undefined,
    bounds: RangeBounds,
  ): SelectRangeOutcome | Promise<SelectRangeOutcome>
  /** formula audit: what a formula reads (ai/formula-audit.ts) */
  tracePrecedents(
    sheetId: string | undefined,
    address: string,
  ): TracePrecedentsOutcome | Promise<TracePrecedentsOutcome>
  /** formula audit: which formulas read a cell (ai/formula-audit.ts) */
  traceDependents(
    sheetId: string | undefined,
    address: string,
  ): TraceDependentsOutcome | Promise<TraceDependentsOutcome>
  /** Batched statistics over a large range (lazy mode streams it through the
   * sidecar without loading the grid) — the supported path for distinct
   * counts / frequency questions that must never become COUNTIF formulas. */
  aggregateRange?(
    sheetId: string | undefined,
    range: RangeBounds,
  ): Promise<{ ok: true; aggregate: RangeAggregate } | { ok: false; error: string }>
  /** `applied` resolves with the real apply result (the lazy path applies async);
   * the tool awaits it so the model never hears "applied" for a batch that failed */
  proposeOperations(
    operations: readonly WorkbookOperation[],
    summary: string,
  ): { ok: true; plan: ChangePlan; applied?: Promise<ApplyOutcome> } | { ok: false; error: string }
  /**
   * The cells the reader withheld from the model, resolved per call so the
   * index never goes stale behind a ref. Optional, and defaulting to "nothing
   * withheld", so a test or a caller without a session pays nothing.
   */
  redactions?(): RedactionIndex
  /** AI create_document: write a new standalone file (xlsx/csv from a
   * worksheet; docx/pdf/md from content) into the default save folder and
   * open it in a new tab (ai/create-document.ts). */
  createDocument?(request: CreateDocumentToolRequest): Promise<CreateDocumentToolOutcome>
}

const MAX_READ_ADDRESSES = 100
/** Max cells per streamed block; the App's ensureRangeLoaded enforces it too. */
export const MAX_READ_RANGE_CELLS = 2000
const MAX_AGGREGATE_CELLS = 1_000_000
const MAX_AGGREGATE_TOP_VALUES = 50
const DEFAULT_AGGREGATE_TOP_VALUES = 10
/** Read-back after write: max number of formula cells whose results are read back */
const MAX_READBACK_FORMULAS = 10
/** Read-back after write: wait time (ms) for Univer's async formula recalc */
const FORMULA_RECALC_DELAY_MS = 300
const MAX_READ_FORMAT_CELLS = 200
const MAX_FIND_RESULTS = 200
const DEFAULT_FIND_RESULTS = 50

export const WORKBOOK_TOOLS: AgentToolDef[] = [
  {
    name: 'get_workbook_context',
    description:
      'Obtenha uma visão geral da pasta de trabalho: todas as planilhas (id/nome/extensão de dados linhas-colunas), planilha ativa, seleção atual, endereços de células não vazias conhecidos. ' +
      'Para perguntas sobre tamanho dos dados (quantas linhas / quanto dados), responda pela extensão de dados aqui em vez de ler bloco por bloco; use read_range ou read_cells quando forem necessários valores concretos.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'read_range',
    description:
      'Leia valores/fórmulas atuais por intervalo retangular, retornando uma grade com números de linha e letras de coluna. ' +
      'O intervalo solicitado não é a extensão de dados da planilha: nunca infira o total de linhas ou registros a partir da última linha; use get_workbook_context. ' +
      'Esta é a forma preferida de ler dados; máximo de 2000 células — leia regiões maiores em várias chamadas.',
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Intervalo como "A1:D20"; uma única célula como "B2" também é aceita',
        },
        sheetId: {
          type: 'string',
          description:
            'Planilha de origem (id obtido em get_workbook_context); lê a planilha ativa quando omitida',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'aggregate_range',
    description:
      'Calcule estatísticas para um intervalo sem lê-lo nem modificá-lo célula por célula: contagem de não vazias, contagem de valores distintos, ' +
      'soma/média/mín/máx numéricos e os valores mais frequentes. Lida com intervalos muito grandes (até 1.000.000 de células) com eficiência. ' +
      'SEMPRE use isto para perguntas como "quantos fornecedores/clientes distintos", distribuições de valores ou totais de coluna em planilhas grandes — ' +
      'nunca repita read_range sobre muitos dados e nunca escreva fórmulas de contagem distinta com COUNTIF/SUMPRODUCT (são rejeitadas por serem caras demais). ' +
      'Agrege uma coluna por vez para obter contagens distintas significativas.',
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Intervalo como "D2:D88588" (normalmente uma coluna, sem o cabeçalho)',
        },
        sheetId: {
          type: 'string',
          description: 'Id da planilha de destino; a planilha ativa quando omitida',
        },
        topValues: {
          type: 'number',
          description: 'Quantos dos valores mais frequentes retornar (0-50, padrão 10)',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'load_guide',
    description:
      'Carregue para o contexto os documentos de guia de operações (definições de campos, convenções, erros comuns). Exceto para leituras/gravações de uma única célula bem básicas, carregue os guias relevantes antes de gerar propose_operations; vários podem ser carregados de uma vez. ' +
      `Guias disponíveis: ${guideCatalogSummary()}`,
    inputSchema: {
      type: 'object',
      properties: {
        guides: {
          type: 'array',
          items: { type: 'string' },
          description: 'Nomes dos guias a carregar, ex.: ["writing","formatting"]',
        },
      },
      required: ['guides'],
    },
  },
  {
    name: 'read_formats',
    description:
      'Leia os formatos explícitos das células em um intervalo (negrito/itálico/sublinhado/cores/formato de número/alinhamento/bordas); apenas células formatadas são retornadas. ' +
      'Use quando precisar "reaproveitar a formatação de outro lugar" ou inspecionar a formatação atual; máximo de 200 células.',
    inputSchema: {
      type: 'object',
      properties: {
        range: { type: 'string', description: 'Intervalo como "A1:D20"' },
        sheetId: {
          type: 'string',
          description:
            'Planilha de origem (id obtido em get_workbook_context); lê a planilha ativa quando omitida',
        },
      },
      required: ['range'],
    },
  },
  {
    name: 'read_sheet_features',
    description:
      "Leia o estado dos recursos de uma planilha: AutoFilter (intervalo e critérios de coluna), regras de formatação condicional, regras de validação de dados, nomes definidos, " +
      'painéis congelados, estado oculta/protegida, formas e imagens, e configuração de página pendente de salvamento nesta sessão. ' +
      'Leia o estado atual antes de modificar ou limpar qualquer uma dessas configurações existentes — nunca as altere às cegas.',
    inputSchema: {
      type: 'object',
      properties: {
        sheetId: {
          type: 'string',
          description: 'Id da planilha de destino; lê a planilha ativa quando omitida',
        },
      },
      required: [],
    },
  },
  {
    name: 'read_cells',
    description:
      'Leia valores/fórmulas atuais de células específicas e espalhadas (use read_range para regiões contíguas). Sempre leia as células afetadas antes de gravar — nunca presuma seu conteúdo.',
    inputSchema: {
      type: 'object',
      properties: {
        addresses: {
          type: 'array',
          items: { type: 'string' },
          description: 'Lista de endereços de células, ex.: ["A1","B2"], máximo 100',
        },
        sheetId: {
          type: 'string',
          description:
            'Planilha de origem (id obtido em get_workbook_context); lê a planilha ativa quando omitida',
        },
      },
      required: ['addresses'],
    },
  },
  {
    name: 'find_cells',
    description:
      'Pesquise em toda a pasta de trabalho (ou em uma planilha) células cujo valor ou fórmula corresponda a uma consulta; retorna Planilha!Endereço com o conteúdo da célula. ' +
      'Texto simples corresponde como substring sem diferenciar maiúsculas/minúsculas; regex=true trata a consulta como uma regex JavaScript sem diferenciar maiúsculas/minúsculas; ' +
      'errors_only=true encontra células com erro de fórmula (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A, #NUM!, #NULL!) e, nesse caso, query pode ser omitida. ' +
      'Em células de fórmula, o valor correspondido é o último valor calculado. Prefira isto a percorrer read_range em páginas ao localizar dados ou auditar erros.',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Texto ou regex a encontrar; pode ser omitida somente quando errors_only=true',
        },
        regex: {
          type: 'boolean',
          description: 'Tratar query como uma regex JavaScript (padrão false)',
        },
        look_in: {
          type: 'string',
          enum: ['values', 'formulas', 'both'],
          default: 'both',
          description: 'O que comparar (padrão both)',
        },
        sheetId: {
          type: 'string',
          description: 'Restringir a busca a uma planilha; busca em todas quando omitido',
        },
        errors_only: {
          type: 'boolean',
          description: 'Encontrar células cujo valor calculado é um erro de fórmula (padrão false)',
        },
        max_results: {
          type: 'integer',
          description: `Número máximo de correspondências retornadas, padrão ${DEFAULT_FIND_RESULTS}, máximo ${MAX_FIND_RESULTS}`,
        },
      },
      required: [],
    },
  },
  {
    name: 'select_range',
    description:
      "Selecione um intervalo na grade e role a visualização do usuário até ele, ativando sua planilha. " +
      'É apenas navegação de visualização — não altera dados, mas substitui o que o usuário havia selecionado. ' +
      'Use somente quando ele pedir para ser levado até lá ("me leve lá", "selecione essas linhas"); para apenas apontar um ' +
      'local, cite-o como [C42](sheetnav://C42) na sua resposta e deixe que ele clique.',
    inputSchema: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'Intervalo como "A1:D20"; uma única célula também é aceita',
        },
        sheetId: { type: 'string', description: 'Id da planilha de destino; usa a ativa por padrão' },
      },
      required: ['range'],
    },
  },
  {
    name: 'trace_precedents',
    description:
      'Liste as células/intervalos que uma fórmula lê (suas antecedentes) com os valores atuais, sinalizando qualquer antecedente que contenha um valor de erro — ' +
      'o primeiro passo ao diagnosticar uma fórmula quebrada ("por que C10 está #DIV/0!?"). Profundidade 1; chame novamente em um antecedente suspeito para avançar na cadeia. ' +
      'Nomes definidos não são expandidos — resolva-os com read_sheet_features e rastreie seus intervalos diretamente.',
    inputSchema: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'Célula de fórmula a auditar, ex.: "C10"' },
        sheetId: {
          type: 'string',
          description: 'Planilha onde a célula está; usa a ativa por padrão',
        },
      },
      required: ['address'],
    },
  },
  {
    name: 'trace_dependents',
    description:
      'Encontre todas as fórmulas da pasta de trabalho (todas as planilhas) que leem uma determinada célula — suas dependentes. ' +
      'Use antes de alterar ou excluir uma célula para ver o que quebraria, ou para acompanhar como um valor de erro se propaga adiante. ' +
      'Dependentes que chegam à célula somente por um nome definido não são detectadas.',
    inputSchema: {
      type: 'object',
      properties: {
        address: { type: 'string', description: 'Célula de destino, ex.: "B2"' },
        sheetId: {
          type: 'string',
          description: 'Planilha onde a célula está; usa a ativa por padrão',
        },
      },
      required: ['address'],
    },
  },
  {
    name: 'propose_operations',
    description:
      'Envie um lote de operações de alteração, aplicadas à pasta de trabalho imediatamente (o usuário pode reverter com o botão [Undo] ou ⌘Z). Operações básicas: ' +
      '{op:"set_cell",sheetId,address,value} | {op:"set_formula",sheetId,address,formula(starts with =)} | ' +
      '{op:"clear_cell",sheetId,address} | {op:"rename_sheet",sheetId,name}. ' +
      'As definições de campos das demais operações estão nos guias — use load_guide antes de empregá-las: ' +
      'writing(set_range/fill_range/copy_range/move_range/convert_to_values/clear_range/find_replace) | formatting(format_range) | ' +
      'layout(sort_range/merge_cells/unmerge_cells/set_row_height/set_col_width/set_rows_hidden/set_cols_hidden/set_freeze/set_page_setup) | ' +
      'structure(insert_rows/delete_rows/insert_cols/delete_cols/add_sheet/delete_sheet/' +
      'duplicate_sheet/set_sheet_hidden/move_sheet/protect_sheet) | ' +
      'charts(add_chart/edit_chart/delete_visual/add_sparkline/add_shape/edit_shape/add_image) | ' +
      'pivot(add_pivot/refresh_pivot) | ' +
      'table(add_table/add_table_row/add_table_column/delete_table_row/delete_table_column/delete_table) | ' +
      'data(set_hyperlink/set_filter/clear_filter/set_filter_criteria/add_conditional_format/' +
      'clear_conditional_formats/set_data_validation/set_note/add_defined_name/delete_defined_name). ' +
      'Limites: operações estruturais (inserção/exclusão de linhas e colunas, adicionar/excluir/duplicar/mover/ocultar planilha) não podem compartilhar um lote com outras classes; no máximo 2000 alterações de célula expandidas — ' +
      'exceto as operações em massa por intervalo fill_range / copy_range / move_range / convert_to_values / clear_range / find_replace / sort_range / format_range, que tratam até 200.000 células em uma única operação ' +
      '(use fill_range para preencher uma fórmula ou padrão ao longo de uma coluna inteira em vez de lotes enormes de set_range, ' +
      'copy_range para duplicar um bloco grande de uma vez, move_range para mover/recortar um bloco — também para outra planilha — em vez de redigitá-lo célula por célula, convert_to_values para congelar fórmulas em seus valores calculados); ' +
      'sheetId deve ser um id retornado por get_workbook_context.',
    inputSchema: {
      type: 'object',
      properties: {
        operations: {
          type: 'array',
          items: { type: 'object' },
          description: 'Array de operações no formato de união discriminada da DSL da pasta de trabalho',
        },
        summary: { type: 'string', description: 'Resumo em uma frase deste lote de alterações' },
      },
      required: ['operations', 'summary'],
    },
  },
  {
    name: 'create_document',
    description:
      'Crie um NOVO arquivo independente na pasta de salvamento padrão e abra-o em uma nova aba; a pasta de trabalho atual não é modificada. ' +
      "Os tipos 'xlsx' (padrão) e 'csv' exportam UMA planilha desta pasta de trabalho: informe sheetId (por padrão, a planilha ativa); o arquivo recebe os valores exibidos atualmente pela planilha (resultados de fórmulas; fórmulas e formatação não são transportadas) e content deve ser omitido. " +
      'Para dividir uma pasta de trabalho em arquivos separados, chame uma vez por planilha. Para exportar dados que ainda não estão em uma planilha, grave-os em uma nova planilha primeiro (add_sheet + set_range) e depois exporte essa planilha. ' +
      "Os tipos 'docx' e 'pdf' recebem HTML simples em content (<h1>-<h6>, <p>, <ul>/<ol>/<li>, <table>, <pre>, <blockquote>; inline <strong>/<em>/<u>/<s>); o tipo 'md' recebe código Markdown; o tipo 'html' recebe uma página HTML completa e independente (abre no editor HTML) — use estes quando o usuário quiser um relatório/resumo como documento próprio. " +
      'title se torna o nome do arquivo; para xlsx/csv, o padrão é o nome da planilha.',
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          enum: ['xlsx', 'csv', 'docx', 'pdf', 'md', 'html'],
          description: "tipo de arquivo de destino (padrão 'xlsx')",
        },
        sheetId: {
          type: 'string',
          description:
            'somente xlsx/csv: a planilha a exportar (id obtido em get_workbook_context); por padrão, a ativa',
        },
        title: {
          type: 'string',
          description:
            'nome do arquivo sem extensão (obrigatório para docx/pdf/md; para xlsx/csv, o padrão é o nome da planilha)',
        },
        content: {
          type: 'string',
          description:
            'docx/pdf: HTML simples; md: código Markdown. Proibido para xlsx/csv — seus dados vêm da planilha',
        },
      },
      required: [],
    },
  },
]

export interface ToolExecution {
  output: string
  isError?: boolean
  /** true when propose_operations auto-applied a batch of changes */
  mutated: boolean
  summary: string
}

const fail = (summary: string, output: string): ToolExecution => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

/** Optional sheetId input shared by the read tools: validated against the
 * workbook's sheet list so a bad id fails with a clear message instead of
 * silently reading empty cells. */
function parseReadSheetId(
  input: Record<string, unknown>,
  info: ActiveSheetInfo,
  summary: string,
): { sheetId: string | undefined } | { fail: ToolExecution } {
  const raw = input.sheetId
  if (raw === undefined || raw === null) return { sheetId: undefined }
  if (typeof raw !== 'string' || !raw.trim()) {
    return { fail: fail(summary, 'sheetId deve ser uma string não vazia quando informado') }
  }
  const sheetId = raw.trim()
  if (!info.sheets.some((sheet) => sheet.id === sheetId)) {
    return {
      fail: fail(summary, `Planilha desconhecida: ${sheetId} (use um id de get_workbook_context)`),
    }
  }
  return { sheetId }
}

/** Clamp a range to the sheet's data extent before a streaming ensure; null
 * when the range lies entirely outside it (nothing to stream — cells there
 * are empty by definition). Sheets without a known extent pass through. */
function clampToExtent(bounds: RangeBounds, sheet: SheetRef | undefined): RangeBounds | null {
  if (sheet?.rows === undefined || sheet.columns === undefined) return bounds
  const endRow = Math.min(bounds.endRow, sheet.rows - 1)
  const endColumn = Math.min(bounds.endColumn, sheet.columns - 1)
  if (endRow < bounds.startRow || endColumn < bounds.startColumn) return null
  return { startRow: bounds.startRow, startColumn: bounds.startColumn, endRow, endColumn }
}

const RANGE_NOT_LOADED =
  'As células solicitadas não puderam ser carregadas totalmente; tente novamente após a conclusão da indexação da pasta de trabalho.'

/** mirror of the docs/pdf tool-echo guard: reject tool-protocol output pasted
 * as document content (create_document docx/pdf) */
function contentEchoError(content: string): string | null {
  if (/<\/?tool_response>/i.test(content)) {
    return 'content contém uma tag literal <tool_response> — isso é saída do protocolo de ferramentas, não conteúdo de documento; tente novamente com o HTML real do documento'
  }
  if (/"index"\s*:\s*\d+\s*,\s*"type"\s*:\s*"/.test(content)) {
    return 'content contém um despejo bruto de JSON, não um fragmento HTML; tente novamente com HTML simples (ex.: <p>…</p>)'
  }
  return null
}

/** Shared input validation for the two formula-audit tools. */
/**
 * A reader that hands out a cell's value as the model may see it.
 *
 * The trace tools print cell values, and a cell the reader withheld keeps its
 * real value on the sheet — that is the point of the feature. So every value
 * they print goes through here, and a workbook with no marks pays one property
 * read per cell.
 */
function cellProjector(
  deps: SheetsSkillDeps,
  sheetId: string | undefined,
): (address: string, value: CellScalar) => CellScalar {
  const index = deps.redactions?.() ?? NO_REDACTIONS
  const target = sheetId ?? deps.getActiveSheetInfo().sheetId
  return (address, value) => modelCellValueAt(index, target, address, value)
}

function parseAuditAddress(
  input: Record<string, unknown>,
  summary: string,
): { address: string; sheetId: string | undefined } | { fail: ToolExecution } {
  const raw = input.address
  if (typeof raw !== 'string' || !raw.trim()) {
    return { fail: fail(summary, 'address deve ser uma string não vazia') }
  }
  const address = raw.trim().toUpperCase().replace(/\$/g, '')
  if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(address)) {
    return {
      fail: fail(summary, `Não foi possível interpretar o endereço da célula: ${raw} (esperado um único endereço como "C10")`),
    }
  }
  const sheetIdRaw = input.sheetId
  return {
    address,
    sheetId: typeof sheetIdRaw === 'string' && sheetIdRaw.trim() ? sheetIdRaw.trim() : undefined,
  }
}

export function buildWorkbookContext(deps: SheetsSkillDeps): string {
  const info = deps.getActiveSheetInfo()
  if (info.mode === 'none') return 'Nenhuma planilha está aberta no momento.'
  const dims = (sheet: SheetRef): string => {
    if (sheet.rows === undefined || sheet.columns === undefined) return ''
    if (sheet.rows === 0 || sheet.columns === 0) return ', sem dados (planilha vazia)'
    return `, extensão de dados de cerca de ${sheet.rows} linhas × ${sheet.columns} colunas`
  }
  const active = info.sheets.find((sheet) => sheet.id === info.sheetId)
  const lines = [
    `Planilha ativa: ${info.sheetName} (id=${info.sheetId}${active ? dims(active) : ''})`,
    info.mode === 'demo'
      ? `Modo: pasta de trabalho de demonstração, revisão atual=${info.revision}`
      : info.streaming
        ? 'Modo: arquivo importado em MODO DE STREAMING DE ARQUIVO GRANDE — o arquivo é grande demais para carregá-lo por completo, então a grade ' +
          'materializa a região visualizada sob demanda. Consequências: (1) as fórmulas que você escreve permanecem ativas — suas ' +
          'células referenciadas no arquivo são carregadas no motor automaticamente — mas apenas dentro de um orçamento por sessão de ' +
          '~50.000 células referenciadas; lotes além disso são rejeitados: calcule com aggregate_range e grave valores ' +
          'estáticos em vez disso; (2) para extrair/dividir linhas, use copy_range com filterColumn/filterValues (copia as linhas ' +
          'correspondentes como valores estáticos, lendo os dados reais do arquivo) para uma planilha criada com add_sheet com ' +
          'linhas/colunas dimensionadas para os dados esperados (aggregate_range fornece contagens por valor) — nunca use ' +
          'derramamento de fórmulas FILTER; ' +
          '(3) add_pivot, refresh_pivot, duplicate_sheet e edições de filtro não estão disponíveis nas planilhas ' +
          "do arquivo; sort_range funciona, mas intervalos de 2000 células ou menos devem passar por read_range (carregamento) primeiro. " +
          'Planilhas adicionadas nesta sessão são totalmente ativas e isentas de tudo isso.'
        : 'Modo: arquivo xlsx real importado (algumas regiões ainda podem estar sendo carregadas em streaming)',
  ]
  if (active?.rows && active.columns) {
    lines.push(
      `Área de dados da planilha ativa: A1:${columnLabel(active.columns - 1)}${active.rows}` +
        ' (responda perguntas sobre o tamanho dos dados diretamente por aqui — não some bloco por bloco com read_range)',
    )
  }
  if (info.sheets.length > 1) {
    lines.push(
      `Todas as planilhas: ${info.sheets.map((sheet) => `${sheet.name} (id=${sheet.id}${dims(sheet)})`).join(', ')}`,
    )
  }
  const oversized = info.sheets.filter((sheet) => sheet.readOnlyOversized)
  if (oversized.length > 0) {
    lines.push(
      `Planilhas SOMENTE LEITURA: ${oversized.map((sheet) => sheet.name).join(', ')} — o XML da planilha está acima ` +
        'do limite de patch de salvamento, então edições ali nunca podem ser salvas e toda operação de alteração nelas é rejeitada. ' +
        'Leia-as livremente (como origem em read_range/aggregate_range/copy_range) e grave os resultados em outras planilhas.',
    )
  }
  if (info.selection) {
    const named = info.selectionColumns ?? []
    // The range alone leaves the model to re-derive which column the user meant
    // from the header row; name it here so "this column" resolves by meaning.
    const columns = named.length
      ? ` (${named.length > 1 ? 'as colunas inteiras' : 'a coluna inteira'} ${named.map((name) => `"${name}"`).join(', ')})`
      : ''
    lines.push(
      info.selectionFrozen
        ? `Seleção do usuário: ${info.selection}${columns} — capturada quando ele enviou esta mensagem, então é ` +
            'o que "esta coluna / estas linhas / a parte selecionada" significa. Ela permanece fixa durante ' +
            'toda a execução, mesmo que o usuário clique em outro lugar enquanto você trabalha.'
        : `Seleção atual: ${info.selection}${columns}`,
    )
  }
  if (info.loadedRange) {
    lines.push(`Viewport carregado no momento: ${info.loadedRange} (não é a extensão de dados da planilha)`)
  }
  if (info.merges && info.merges.length > 0) {
    lines.push(`Intervalos mesclados na planilha ativa: ${info.merges.slice(0, 50).join(', ')}`)
  }
  if (info.charts && info.charts.length > 0) {
    lines.push(
      'Gráficos na pasta de trabalho (use o caminho abaixo com edit_chart para editar um gráfico existente; use add_chart para criar um):',
    )
    for (const chart of info.charts) {
      lines.push(
        `- ${chart.path} | título: ${chart.title || '(nenhum)'} | tipo: ${chart.types} | sheetId: ${chart.sheetId}`,
      )
    }
  }
  if (info.knownAddresses.length > 0) {
    lines.push(
      `Células não vazias conhecidas (podem estar incompletas): ${info.knownAddresses.slice(0, 200).join(', ')}`,
    )
  } else {
    lines.push(
      'Ainda não há informações sobre células não vazias; leia sob demanda com read_range/read_cells.',
    )
  }
  return lines.join('\n')
}

function describeFormatState(format: CellFormatState): string {
  const parts: string[] = []
  if (format.bold) parts.push('negrito')
  if (format.italic) parts.push('itálico')
  if (format.underline) parts.push('sublinhado')
  if (format.strikethrough) parts.push('tachado')
  if (format.fontFamily) parts.push(`fonte ${format.fontFamily}`)
  if (format.fontSize) parts.push(`tamanho ${format.fontSize}`)
  if (format.fontColor) parts.push(`cor da fonte ${format.fontColor}`)
  if (format.fillColor) parts.push(`preenchimento ${format.fillColor}`)
  if (format.numberFormat) parts.push(`formato de número ${format.numberFormat}`)
  if (format.horizontalAlign) parts.push(`alinhamento ${format.horizontalAlign}`)
  if (format.verticalAlign) parts.push(`alinhamento vertical ${format.verticalAlign}`)
  if (format.wrapText) parts.push('quebra de linha')
  if (format.border) {
    parts.push(
      `borda ${format.border.type}${format.border.color ? ` ${format.border.color}` : ''}`,
    )
  }
  return parts.join(', ') || '(nenhum)'
}

// Cell text is emitted into tab/newline-delimited tool output (read_range grid,
// read_cells lists, plan summaries), where raw control characters would tear the
// line/column structure apart and scramble the model's view of the grid. Escape
// them — and backslash itself, so the encoding stays unambiguous. Univer streams
// in-cell paragraph breaks as \r (the file model uses \n; see edit-journal's
// dataStream conversion), so CR/CRLF are normalized to \n first: the model sees
// a single line-break representation, and echoing the same `\n` escape inside
// JSON string values of write operations round-trips into real line breaks.
function escapeCellText(text: string): string {
  return text
    .replace(/\\/g, '\\\\')
    .replace(/\t/g, '\\t')
    .replace(/\r\n?/g, '\n')
    .replace(/\n/g, '\\n')
}

function formatCellScalar(cell: { value: CellScalar; formula?: string | undefined }): string {
  // Formula cells prefer the value (the AI reasons from computed values, with
  // the formula as provenance); when the value isn't computed yet, give only the
  // formula
  if (cell.formula) {
    return cell.value === null || cell.value === undefined
      ? escapeCellText(cell.formula)
      : `${escapeCellText(String(cell.value))} (${escapeCellText(cell.formula)})`
  }
  if (cell.value === null) return '(vazia)'
  return escapeCellText(String(cell.value))
}

function formatPlanSummary(plan: ChangePlan): string {
  const parts: string[] = []
  if (plan.structuralChanges.length > 0) {
    parts.push(plan.structuralChanges.map((change) => change.label).join('; '))
  }
  if (plan.formatChanges.length > 0) {
    parts.push(plan.formatChanges.map((change) => change.label).join('; '))
  }
  if (plan.cellChanges.length > 0) {
    const shown = plan.cellChanges.slice(0, 20)
    const rest = plan.cellChanges.length - shown.length
    parts.push(
      shown
        .map((c) => `${c.address}: ${formatCellScalar(c.before)} → ${formatCellScalar(c.after)}`)
        .join('; ') + (rest > 0 ? `; …${rest} células a mais` : ''),
    )
  }
  if (plan.sheetRenames.length > 0) {
    parts.push(plan.sheetRenames.map((r) => `planilha ${r.before} → ${r.after}`).join('; '))
  }
  return parts.join(' | ') || '(nenhuma alteração)'
}

export function executeWorkbookTool(
  call: AgentToolCall,
  deps: SheetsSkillDeps,
): ToolExecution | Promise<ToolExecution> {
  switch (call.name) {
    case 'get_workbook_context':
      return {
        output: buildWorkbookContext(deps),
        mutated: false,
        summary: t('aiToolWorkbookContext'),
      }

    case 'read_range': {
      const raw = call.input.range
      if (typeof raw !== 'string' || !raw.trim())
        return fail(t('aiToolReadRange'), 'range deve ser uma string não vazia')
      let bounds
      try {
        bounds = parseRange(raw.trim().toUpperCase())
      } catch {
        return fail(t('aiToolReadRange'), `Não foi possível interpretar o intervalo: ${raw}`)
      }
      if (rangeCellCount(bounds) > MAX_READ_RANGE_CELLS) {
        return fail(
          t('aiToolReadRange'),
          `O intervalo contém mais de ${MAX_READ_RANGE_CELLS} células; leia-o em várias chamadas`,
        )
      }
      const info = deps.getActiveSheetInfo()
      const parsedSheet = parseReadSheetId(call.input, info, t('aiToolReadRange'))
      if ('fail' in parsedSheet) return parsedSheet.fail
      const sheetId = parsedSheet.sheetId
      const target = info.sheets.find((sheet) => sheet.id === (sheetId ?? info.sheetId))
      if (target?.rows !== undefined && target.columns !== undefined) {
        if (target.rows === 0 || target.columns === 0) {
          return fail(t('aiToolReadRange'), 'A planilha não tem dados (extensão vazia).')
        }
        if (bounds.endRow >= target.rows || bounds.endColumn >= target.columns) {
          return fail(
            t('aiToolReadRange'),
            `O intervalo solicitado está fora da extensão de dados da planilha A1:${columnLabel(target.columns - 1)}${target.rows}.`,
          )
        }
      }
      const executeRead = (): ToolExecution => {
        const normalizedRange = `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`
        const metadata =
          target?.rows && target.columns
            ? `Metadados de leitura: intervalo solicitado ${normalizedRange}; extensão de dados autoritativa da planilha A1:${columnLabel(target.columns - 1)}${target.rows} (${target.rows} linhas da planilha, incluindo eventual cabeçalho). Não infira o total de linhas ou registros a partir do intervalo solicitado.`
            : `Metadados de leitura: intervalo solicitado ${normalizedRange}; a extensão de dados da planilha é desconhecida. Não infira o total de linhas ou registros a partir do intervalo solicitado.`
        const addresses: string[] = []
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            addresses.push(formatAddress(row, column))
          }
        }
        const cells = deps.readCells(addresses, sheetId)
        const header = [
          '',
          ...Array.from({ length: bounds.endColumn - bounds.startColumn + 1 }, (_, offset) =>
            columnLabel(bounds.startColumn + offset),
          ),
        ].join('\t')
        const rows: string[] = [metadata, header]
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          const columns: string[] = [String(row + 1)]
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            const cell = cells[formatAddress(row, column)]
            columns.push(
              cell
                ? cell.value === null
                  ? escapeCellText(cell.formula ?? '')
                  : formatCellScalar(cell)
                : '',
            )
          }
          rows.push(columns.join('\t'))
        }
        return {
          output: rows.join('\n'),
          mutated: false,
          summary: t('aiToolReadRangeOf', { range: raw.trim().toUpperCase() }),
        }
      }
      const loading = deps.ensureRangeLoaded?.(bounds, sheetId)
      if (loading instanceof Promise) {
        return loading.then((loaded) =>
          loaded
            ? executeRead()
            : fail(
                t('aiToolReadRange'),
                'O intervalo solicitado não pôde ser carregado totalmente; tente novamente após a conclusão da indexação da pasta de trabalho.',
              ),
        )
      }
      if (loading === false) {
        return fail(
          t('aiToolReadRange'),
          'O intervalo solicitado não pôde ser carregado totalmente; tente novamente após a conclusão da indexação da pasta de trabalho.',
        )
      }
      return executeRead()
    }

    case 'aggregate_range': {
      const raw = call.input.range
      if (typeof raw !== 'string' || !raw.trim())
        return fail(t('aiToolAggregate'), 'range deve ser uma string não vazia')
      const rangeLabel = raw.trim().toUpperCase()
      let bounds
      try {
        bounds = parseRange(rangeLabel)
      } catch {
        return fail(t('aiToolAggregate'), `Não foi possível interpretar o intervalo: ${raw}`)
      }
      if (rangeCellCount(bounds) > MAX_AGGREGATE_CELLS) {
        return fail(
          t('aiToolAggregate'),
          `O intervalo contém mais de ${MAX_AGGREGATE_CELLS.toLocaleString('en-US')} células; agregue uma coluna (ou um bloco menor) por vez.`,
        )
      }
      if (!deps.aggregateRange) {
        return fail(t('aiToolAggregate'), 'aggregate_range não está disponível neste contexto.')
      }
      const parsedSheet = parseReadSheetId(
        call.input,
        deps.getActiveSheetInfo(),
        t('aiToolAggregate'),
      )
      if ('fail' in parsedSheet) return parsedSheet.fail
      const sheetId = parsedSheet.sheetId
      const topRaw = call.input.topValues
      const topValues =
        typeof topRaw === 'number' && Number.isFinite(topRaw)
          ? Math.min(Math.max(Math.floor(topRaw), 0), MAX_AGGREGATE_TOP_VALUES)
          : DEFAULT_AGGREGATE_TOP_VALUES
      return deps.aggregateRange(sheetId, bounds).then((outcome) =>
        outcome.ok
          ? {
              output: formatRangeAggregate(rangeLabel, outcome.aggregate, topValues),
              mutated: false,
              summary: t('aiToolAggregateOf', { range: rangeLabel }),
            }
          : fail(t('aiToolAggregate'), outcome.error),
      )
    }

    case 'load_guide': {
      const raw = call.input.guides
      if (!Array.isArray(raw) || raw.length === 0)
        return fail(t('aiToolLoadGuide'), 'guides deve ser um array não vazio')
      const outcome = loadGuides(raw.map(String))
      if (!outcome.ok) return fail(t('aiToolLoadGuide'), outcome.error)
      return {
        output: outcome.content,
        mutated: false,
        summary: t('aiToolLoadGuideOf', { names: raw.join(', ') }),
      }
    }

    case 'read_formats': {
      const raw = call.input.range
      if (typeof raw !== 'string' || !raw.trim())
        return fail(t('aiToolReadFormats'), 'range deve ser uma string não vazia')
      let bounds
      try {
        bounds = parseRange(raw.trim().toUpperCase())
      } catch {
        return fail(t('aiToolReadFormats'), `Não foi possível interpretar o intervalo: ${raw}`)
      }
      if (rangeCellCount(bounds) > MAX_READ_FORMAT_CELLS) {
        return fail(
          t('aiToolReadFormats'),
          `O intervalo contém mais de ${MAX_READ_FORMAT_CELLS} células; leia-o em várias chamadas`,
        )
      }
      const info = deps.getActiveSheetInfo()
      const parsedSheet = parseReadSheetId(call.input, info, t('aiToolReadFormats'))
      if ('fail' in parsedSheet) return parsedSheet.fail
      const sheetId = parsedSheet.sheetId
      const executeRead = (): ToolExecution => {
        const addresses: string[] = []
        for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
          for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
            addresses.push(formatAddress(row, column))
          }
        }
        const formats = deps.readFormats(addresses, sheetId)
        const lines = Object.entries(formats).map(
          ([address, format]) => `${address}: ${describeFormatState(format)}`,
        )
        return {
          output: lines.length > 0 ? lines.join('\n') : 'Não há formatos explícitos neste intervalo.',
          mutated: false,
          summary: t('aiToolReadFormatsOf', { range: raw.trim().toUpperCase() }),
        }
      }
      const target = info.sheets.find((sheet) => sheet.id === (sheetId ?? info.sheetId))
      const ensureBounds = clampToExtent(bounds, target)
      const loading = ensureBounds ? deps.ensureRangeLoaded?.(ensureBounds, sheetId) : true
      if (loading instanceof Promise) {
        return loading.then((loaded) =>
          loaded ? executeRead() : fail(t('aiToolReadFormats'), RANGE_NOT_LOADED),
        )
      }
      if (loading === false) return fail(t('aiToolReadFormats'), RANGE_NOT_LOADED)
      return executeRead()
    }

    case 'read_sheet_features': {
      const info = deps.getActiveSheetInfo()
      const parsedSheet = parseReadSheetId(call.input, info, t('aiToolSheetFeatures'))
      if ('fail' in parsedSheet) return parsedSheet.fail
      const sheetId = parsedSheet.sheetId
      const executeRead = (): ToolExecution => ({
        output: deps.readSheetFeatures(sheetId),
        mutated: false,
        summary: t('aiToolSheetFeatures'),
      })
      // Filter/CF/validation/freeze models install on a sheet's first
      // stream-in; reading a never-loaded sheet would report false "none"s.
      const loading = deps.ensureRangeLoaded?.(
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        sheetId,
      )
      if (loading instanceof Promise) {
        return loading.then((loaded) =>
          loaded ? executeRead() : fail(t('aiToolSheetFeatures'), RANGE_NOT_LOADED),
        )
      }
      if (loading === false) return fail(t('aiToolSheetFeatures'), RANGE_NOT_LOADED)
      return executeRead()
    }

    case 'read_cells': {
      const raw = call.input.addresses
      if (!Array.isArray(raw) || raw.length === 0)
        return fail(t('aiToolReadCells'), 'addresses deve ser um array não vazio')
      const info = deps.getActiveSheetInfo()
      const parsedSheet = parseReadSheetId(call.input, info, t('aiToolReadCells'))
      if ('fail' in parsedSheet) return parsedSheet.fail
      const sheetId = parsedSheet.sheetId
      const addresses = raw.slice(0, MAX_READ_ADDRESSES).map(String)
      const executeRead = (): ToolExecution => {
        const cells = deps.readCells(addresses, sheetId)
        const lines = addresses.map((addr) => {
          const cell = cells[addr]
          return `${addr}: ${cell ? formatCellScalar(cell) : '(desconhecida)'}`
        })
        return {
          output: lines.join('\n'),
          mutated: false,
          summary: t('aiToolReadCellsCount', { count: addresses.length }),
        }
      }
      const ensure = deps.ensureRangeLoaded
      if (!ensure) return executeRead()
      // Unstreamed cells read as empty, so the covering box of the addresses
      // is loaded first. Per-address fallback loads are NOT an option: the
      // loaded-range bookkeeping is a single rectangle per sheet, so each
      // load would evict the previous one and just-read cells would drop out
      // of the editable region again.
      let box: RangeBounds | null = null
      for (const addr of addresses) {
        let cell: RangeBounds
        try {
          cell = parseRange(addr.trim().toUpperCase())
        } catch {
          continue
        }
        box =
          box === null
            ? cell
            : {
                startRow: Math.min(box.startRow, cell.startRow),
                endRow: Math.max(box.endRow, cell.endRow),
                startColumn: Math.min(box.startColumn, cell.startColumn),
                endColumn: Math.max(box.endColumn, cell.endColumn),
              }
      }
      const target = info.sheets.find((sheet) => sheet.id === (sheetId ?? info.sheetId))
      const boundsBox = box === null ? null : clampToExtent(box, target)
      if (boundsBox === null) return executeRead()
      // Whether the box needs streaming at all (demo workbook, fully preloaded
      // import, session-added sheet) is the callback's call — it also rejects
      // boxes too large to stream as one block.
      return (async (): Promise<ToolExecution> => {
        if (!(await ensure(boundsBox, sheetId))) {
          return fail(
            t('aiToolReadCells'),
            'As células solicitadas não puderam ser carregadas totalmente — tente novamente após a conclusão da indexação da pasta de trabalho; ' +
              `em pastas de trabalho transmitidas por streaming, uma dispersão que abranja mais de ${MAX_READ_RANGE_CELLS} células deve ser dividida em chamadas read_cells mais próximas.`,
          )
        }
        return executeRead()
      })()
    }

    case 'find_cells': {
      const errorsOnly = call.input.errors_only === true
      const query = typeof call.input.query === 'string' ? call.input.query.trim() : ''
      if (!errorsOnly && !query) {
        return fail(
          t('aiToolFindCells'),
          'query deve ser uma string não vazia (ou defina errors_only=true)',
        )
      }
      const lookInRaw = call.input.look_in
      const maxRaw = Number(call.input.max_results)
      const sheetIdRaw = call.input.sheetId
      const options: FindCellsOptions = {
        query,
        regex: call.input.regex === true,
        lookIn: lookInRaw === 'values' || lookInRaw === 'formulas' ? lookInRaw : 'both',
        errorsOnly,
        sheetId:
          typeof sheetIdRaw === 'string' && sheetIdRaw.trim() ? sheetIdRaw.trim() : undefined,
        maxResults:
          Number.isFinite(maxRaw) && maxRaw >= 1
            ? Math.min(Math.floor(maxRaw), MAX_FIND_RESULTS)
            : DEFAULT_FIND_RESULTS,
      }
      const finish = (result: FindCellsOutcome): ToolExecution => {
        if (result.error) return fail(t('aiToolFindCells'), result.error)
        const lines = result.matches.map(
          (match) => `${match.sheetName}!${match.address}: ${formatCellScalar(match)}`,
        )
        const header =
          result.matches.length === 0
            ? result.truncated
              ? 'Nenhuma célula correspondente encontrada na região verificada, mas a busca parou no limite de varredura antes de cobrir toda a pasta de trabalho — NÃO conclua que não há correspondências; restrinja a consulta ou o escopo com sheetId e tente novamente.'
              : 'Nenhuma célula correspondente encontrada.'
            : `${result.matches.length} célula(s) correspondente(s)` +
              (result.truncated
                ? ' (a busca parou no limite — pode haver mais; restrinja a consulta ou o escopo com sheetId):'
                : ':')
        if (result.incompleteSheets.length > 0) {
          lines.push(
            `Observação: a indexação em segundo plano não terminou em ${result.incompleteSheets.join(', ')} — podem faltar correspondências lá; tente novamente depois se algo esperado estiver ausente.`,
          )
        }
        return {
          output: [header, ...lines].join('\n'),
          mutated: false,
          summary: errorsOnly
            ? t('aiToolFindErrors', { count: result.matches.length })
            : t('aiToolFindCellsOf', { query, count: result.matches.length }),
        }
      }
      const outcome = deps.findCells(options)
      return outcome instanceof Promise ? outcome.then(finish) : finish(outcome)
    }

    case 'select_range': {
      const raw = call.input.range
      if (typeof raw !== 'string' || !raw.trim())
        return fail(t('aiToolSelectRange'), 'range deve ser uma string não vazia')
      let bounds: RangeBounds
      try {
        bounds = parseRange(raw.trim().toUpperCase().replace(/\$/g, ''))
      } catch {
        return fail(t('aiToolSelectRange'), `Não foi possível interpretar o intervalo: ${raw}`)
      }
      const sheetIdRaw = call.input.sheetId
      const sheetId =
        typeof sheetIdRaw === 'string' && sheetIdRaw.trim() ? sheetIdRaw.trim() : undefined
      const normalized =
        bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn
          ? formatAddress(bounds.startRow, bounds.startColumn)
          : `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`
      const finish = (result: SelectRangeOutcome): ToolExecution => {
        if (!result.ok) return fail(t('aiToolSelectRange'), result.error ?? 'Falha na seleção')
        const label = `${result.sheetName ?? ''}!${normalized}`
        return {
          output: `Selecionado ${label} e a visualização foi rolada até lá.`,
          mutated: false,
          summary: t('aiToolSelectRangeOf', { range: label }),
        }
      }
      const outcome = deps.selectRange(sheetId, bounds)
      return outcome instanceof Promise ? outcome.then(finish) : finish(outcome)
    }

    case 'trace_precedents': {
      const parsed = parseAuditAddress(call.input, t('aiToolTracePrecedents'))
      if ('fail' in parsed) return parsed.fail
      const project = cellProjector(deps, parsed.sheetId)
      const finish = (result: TracePrecedentsOutcome): ToolExecution => {
        if (result.error) return fail(t('aiToolTracePrecedents'), result.error)
        const summary = t('aiToolTracePrecedentsOf', { address: parsed.address })
        if (!result.formula) {
          return {
            output: `${parsed.address} não é uma célula de fórmula; valor: ${formatCellScalar({ value: project(parsed.address, result.value ?? null) })}. Nada a rastrear acima — use trace_dependents para ver o que a lê.`,
            mutated: false,
            summary,
          }
        }
        // A formula's own value is its precedents' values added up, so a trace
        // that reaches a withheld cell has to withhold the result too — showing
        // `C10 = 42` when `B2` is the withheld number gives the number away
        // just as plainly as printing `B2`.
        const refValues = result.refs.flatMap((ref) => ref.samples)
        const aPrecedentIsWithheld = refValues.some((s) => project(s.address, s.value) !== s.value)
        const lines = [
          `${parsed.address} = ${formatCellScalar({
            value: aPrecedentIsWithheld
              ? `(${parsed.address} lê uma célula retida do modelo; valor não exibido)`
              : project(parsed.address, result.value ?? null),
            formula: result.formula,
          })}`,
          `Lê ${result.refs.length}${result.truncatedRefs ? '+' : ''} referência(s):`,
        ]
        for (const ref of result.refs) {
          if (ref.external) {
            lines.push(
              `- ${ref.label}: referência externa/não resolvida (outra pasta de trabalho ou planilha desconhecida)`,
            )
            continue
          }
          const shown = ref.samples
            .map(
              (sample) =>
                `${sample.address}=${formatCellScalar({ ...sample, value: project(sample.address, sample.value) })}`,
            )
            .join('; ')
          const rest = ref.cellCount - ref.samples.length
          lines.push(
            `- ${ref.label} (${ref.cellCount} célula(s))${ref.hasError ? ' ⚠️ contém valores de erro' : ''}: ${shown}${rest > 0 ? `; …${rest} a mais` : ''}`,
          )
        }
        if (result.truncatedRefs) {
          lines.push('Observação: a fórmula tem mais referências do que as mostradas aqui.')
        }
        if (result.usesNames) {
          lines.push(
            'Observação: a fórmula também usa nomes definidos — liste-os com read_sheet_features e rastreie seus intervalos diretamente.',
          )
        }
        return { output: lines.join('\n'), mutated: false, summary }
      }
      const outcome = deps.tracePrecedents(parsed.sheetId, parsed.address)
      return outcome instanceof Promise ? outcome.then(finish) : finish(outcome)
    }

    case 'trace_dependents': {
      const parsed = parseAuditAddress(call.input, t('aiToolTraceDependents'))
      if ('fail' in parsed) return parsed.fail
      const project = cellProjector(deps, parsed.sheetId)
      // the cell being traced is withheld, so everything downstream of it is
      // too: `=B2` beside a withheld B2 shows B2's value
      const tracedIsWithheld = project(parsed.address, null) !== null
      const finish = (result: TraceDependentsOutcome): ToolExecution => {
        if (result.error) return fail(t('aiToolTraceDependents'), result.error)
        const summary = t('aiToolTraceDependentsOf', {
          address: parsed.address,
          count: result.dependents.length,
        })
        const lines =
          result.dependents.length === 0
            ? [
                `Nenhuma fórmula lê ${parsed.address}. (Dependentes que chegam a ela somente por um nome definido não são detectadas.)`,
              ]
            : [
                `${result.dependents.length}${result.truncated ? '+' : ''} célula(s) de fórmula leem ${parsed.address}:`,
                ...result.dependents.map((dep) => {
                  const sheetId = deps
                    .getActiveSheetInfo()
                    .sheets.find((sh) => sh.name === dep.sheetName)?.id
                  const withheld =
                    tracedIsWithheld ||
                    modelCellValueAt(
                      deps.redactions?.() ?? NO_REDACTIONS,
                      sheetId,
                      dep.address,
                      dep.value,
                    ) !== dep.value
                  return `- ${dep.sheetName}!${dep.address} = ${formatCellScalar({
                    value: withheld
                      ? '(lê uma célula retida do modelo; valor não exibido)'
                      : dep.value,
                    formula: dep.formula,
                  })}`
                }),
              ]
        if (result.truncated) {
          lines.push('Observação: parou no limite de resultados — existem mais dependentes.')
        }
        if (result.incompleteSheets.length > 0) {
          lines.push(
            `Observação: a indexação em segundo plano não terminou em ${result.incompleteSheets.join(', ')} — podem faltar dependentes lá.`,
          )
        }
        return { output: lines.join('\n'), mutated: false, summary }
      }
      const outcome = deps.traceDependents(parsed.sheetId, parsed.address)
      return outcome instanceof Promise ? outcome.then(finish) : finish(outcome)
    }

    case 'propose_operations': {
      const rawOps = call.input.operations
      const summaryInput = call.input.summary
      if (!Array.isArray(rawOps) || rawOps.length === 0) {
        return fail(t('aiToolPropose'), 'operations deve ser um array não vazio')
      }
      if (typeof summaryInput !== 'string' || !summaryInput.trim()) {
        return fail(t('aiToolPropose'), 'summary não deve estar vazio')
      }
      const parsedOps = z.array(workbookOperationSchema).safeParse(rawOps)
      if (!parsedOps.success) {
        return fail(t('aiToolPropose'), describeOperationErrors(rawOps, parsedOps.error))
      }
      const operations: WorkbookOperation[] = parsedOps.data
      // A cell the reader withheld still holds real data the model never saw,
      // so a write landing on one destroys something the user chose to keep —
      // silently, with the file still opening afterwards. Checked here, before
      // the batch is planned or applied, so a refusal changes nothing.
      const refusal = redactGuardForOps(
        operations,
        deps.redactions?.() ?? NO_REDACTIONS,
        (sheetId) => deps.getActiveSheetInfo().sheets.find((sheet) => sheet.id === sheetId)?.name,
        (sheetName) =>
          deps.getActiveSheetInfo().sheets.find((sheet) => sheet.name === sheetName)?.id,
      )
      if (refusal) {
        return fail(
          t('aiToolPropose'),
          `Rejeitado — nenhuma das ${operations.length} operação(ões) foi aplicada (um lote é tudo ou nada): ${refusal.reason}`,
        )
      }
      const outcome = deps.proposeOperations(operations, summaryInput.trim())
      if (!outcome.ok) {
        return fail(
          t('aiToolPropose'),
          `Rejeitado — nenhuma das ${operations.length} operação(ões) foi aplicada (um lote é tudo ou nada): ${outcome.error}`,
        )
      }
      const summary = summaryInput.trim()
      const finish = (
        appliedNotices: readonly string[] = [],
      ): ToolExecution | Promise<ToolExecution> => {
        const notes = [...outcome.plan.warnings, ...appliedNotices]
        const warnings = notes.length > 0 ? `\nObservação: ${notes.join('; ')}` : ''
        const opCount =
          outcome.plan.cellChanges.length +
          outcome.plan.formatChanges.length +
          outcome.plan.sheetRenames.length +
          outcome.plan.structuralChanges.length
        const base = `Aplicadas automaticamente ${opCount} alteração(ões) (desfazer pelo botão [Undo] no painel lateral ou ⌘Z): ${formatPlanSummary(outcome.plan)}${warnings}`
        // Read-back after write (write → verify): formula cells fetch their
        // computed values after the async recalc, so the AI sees real results and
        // errors like #REF!/#DIV/0! instead of just what it wrote.
        const formulaCells: { sheetId: string; address: string }[] = outcome.plan.cellChanges
          .filter((c) => c.after.formula)
          .map((c) => ({ sheetId: c.sheetId, address: c.address }))
        // fill_range / copy_range / move_range apply as range-level bulk writes (no
        // per-cell plan entries), so read back each target's corners to
        // confirm the write actually landed and its shifted formulas compute.
        for (const op of operations) {
          if (op.op !== 'fill_range' && op.op !== 'copy_range' && op.op !== 'move_range') continue
          const bounds = op.op === 'fill_range' ? parseRange(op.target) : copyTargetBounds(op)
          const first = formatAddress(bounds.startRow, bounds.startColumn)
          formulaCells.push({ sheetId: op.sheetId, address: first })
          // A filtered copy's real extent is smaller than the worst-case
          // bounds (and it writes no formulas) — read back only the anchor.
          if (op.op === 'copy_range' && op.filterColumn !== undefined) continue
          const last = formatAddress(bounds.endRow, bounds.endColumn)
          if (last !== first) formulaCells.push({ sheetId: op.sheetId, address: last })
        }
        if (formulaCells.length === 0) {
          return { output: base, mutated: true, summary }
        }
        return (async (): Promise<ToolExecution> => {
          await new Promise((resolve) => setTimeout(resolve, FORMULA_RECALC_DELAY_MS))
          const shown = formulaCells.slice(0, MAX_READBACK_FORMULAS)
          // Cells are read per target sheet (operations may span sheets);
          // addresses are prefixed with the sheet name only when they do.
          const bySheet = new Map<string, string[]>()
          for (const cell of shown) {
            bySheet.set(cell.sheetId, [...(bySheet.get(cell.sheetId) ?? []), cell.address])
          }
          const sheetNames = new Map(
            deps.getActiveSheetInfo().sheets.map((sheet) => [sheet.id, sheet.name]),
          )
          const lines: string[] = []
          const redactions = deps.redactions?.() ?? NO_REDACTIONS
          for (const [cellSheetId, addresses] of bySheet) {
            const cells = deps.readCells(addresses, cellSheetId)
            const prefix = bySheet.size > 1 ? `${sheetNames.get(cellSheetId) ?? cellSheetId}!` : ''
            for (const addr of addresses) {
              const raw = cells[addr]?.value
              const v =
                raw === null || raw === undefined
                  ? raw
                  : modelCellValueAt(redactions, cellSheetId, addr, raw)
              lines.push(
                `${prefix}${addr} = ${v === null || v === undefined ? '(ainda calculando; verifique com read_cells)' : String(v)}`,
              )
            }
          }
          const rest = formulaCells.length - shown.length
          const hasError = lines.some((l) =>
            /#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!)/.test(l),
          )
          return {
            output:
              `${base}\nResultados das fórmulas: ${lines.join('; ')}${rest > 0 ? `; …${rest} células de fórmula a mais` : ''}` +
              (hasError
                ? '\n⚠️ Há valores de erro em fórmulas — verifique referências/divisores e corrija-os.'
                : ''),
            mutated: true,
            summary,
          }
        })()
      }
      if (!outcome.applied) return finish()
      return outcome.applied.then((applied) => {
        if (applied.ok) return finish(applied.notices)
        const reason = applied.reason ?? 'motivo desconhecido'
        return fail(
          t('aiToolPropose'),
          applied.partiallyApplied
            ? `Falha na aplicação NO MEIO DO LOTE — as operações anteriores à que falhou já foram confirmadas: ${reason}. ` +
                'Leia os intervalos afetados para ver o estado atual antes de continuar; ' +
                (applied.undoDropped
                  ? 'a parte confirmada é grande demais para o histórico de desfazer — ⌘Z NÃO a reverterá.'
                  : 'todo o lote parcial é uma única etapa de desfazer (⌘Z / [Undo]).')
            : `Falha na aplicação — a pasta de trabalho está INALTERADA: ${reason}. ` +
                'Não diga ao usuário que as alterações foram feitas; ajuste as operações e tente novamente, ou explique a falha.',
        )
      })
    }

    case 'create_document': {
      const summary = t('aiToolCreateDocument')
      const create = deps.createDocument
      if (!create) return fail(summary, 'create_document não está disponível neste contexto.')
      const typeRaw = call.input.type === undefined ? 'xlsx' : String(call.input.type)
      if (
        typeRaw !== 'xlsx' &&
        typeRaw !== 'csv' &&
        typeRaw !== 'docx' &&
        typeRaw !== 'pdf' &&
        typeRaw !== 'md' &&
        typeRaw !== 'html'
      ) {
        return fail(summary, 'type deve ser um de xlsx/csv/docx/pdf/md/html')
      }
      const title = typeof call.input.title === 'string' ? call.input.title.trim() : ''
      if (typeRaw === 'xlsx' || typeRaw === 'csv') {
        if (typeof call.input.content === 'string' && call.input.content.trim() !== '') {
          return fail(
            summary,
            'content é proibido para xlsx/csv — eles exportam uma planilha. Para criar um arquivo a partir de novos dados, ' +
              'grave-os em uma planilha primeiro (add_sheet + set_range via propose_operations) e depois exporte essa planilha.',
          )
        }
        const parsedSheet = parseReadSheetId(call.input, deps.getActiveSheetInfo(), summary)
        if ('fail' in parsedSheet) return parsedSheet.fail
        return create({
          type: typeRaw,
          sheetId: parsedSheet.sheetId,
          ...(title ? { title } : {}),
        }).then((outcome) => {
          if (!outcome.ok) return fail(summary, outcome.error)
          const note = outcome.hadFormulas
            ? ' Observação: a planilha contém fórmulas — o novo arquivo contém apenas seus valores calculados atuais (sem fórmulas nem formatação).'
            : ''
          return {
            output:
              `Criado ${outcome.name}${outcome.path ? ` em ${outcome.path}` : ''} a partir da planilha ` +
              `"${outcome.sheetName ?? ''}" e aberto em uma nova aba. A pasta de trabalho atual está inalterada.${note}`,
            mutated: false,
            summary: t('aiToolCreatedDocument', { name: outcome.name }),
          }
        })
      }
      if (!title) return fail(summary, 'title não deve estar vazio')
      const content = String(call.input.content ?? '')
      if (!content.trim()) return fail(summary, 'content não deve estar vazio')
      if (typeRaw !== 'md') {
        const echo = contentEchoError(content)
        if (echo) return fail(summary, echo)
      }
      return create({ type: typeRaw, title, content }).then((outcome) => {
        if (!outcome.ok) return fail(summary, outcome.error)
        return {
          output: outcome.path
            ? `Novo documento criado em ${outcome.path} e aberto em uma nova aba.`
            : `Novo documento "${outcome.name}" criado em uma nova aba; ele se salva automaticamente na pasta padrão.`,
          mutated: false,
          summary: t('aiToolCreatedDocument', { name: outcome.name }),
        }
      })
    }

    default:
      return fail(call.name, `Ferramenta desconhecida: ${call.name}`)
  }
}
