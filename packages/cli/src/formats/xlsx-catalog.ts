import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { z } from 'zod'
import {
  renderGroups,
  signatureFromSchema,
  withFingerprint,
  type JsonSchema,
  type OpCatalog,
  type OpEntry,
} from '../op-catalog'
import { REFUSED_DSL_OPS } from './xlsx-gateway-ops'

/** Presentation order of the guide; a test asserts every supported op sits in exactly one group. */
export const SHEET_OP_GROUPS: Record<string, readonly string[]> = {
  content: [
    'set_cell',
    'set_formula',
    'set_range',
    'clear_cell',
    'clear_range',
    'fill_range',
    'copy_range',
    'move_range',
    'find_replace',
    'sort_range',
    'convert_to_values',
  ],
  format: [
    'format_range',
    'add_conditional_format',
    'clear_conditional_formats',
    'set_data_validation',
  ],
  layout: [
    'merge_cells',
    'unmerge_cells',
    'set_row_height',
    'set_col_width',
    'set_rows_hidden',
    'set_cols_hidden',
    'set_freeze',
    'set_filter',
    'set_filter_criteria',
    'clear_filter',
    'set_page_setup',
    'protect_sheet',
  ],
  data: [
    'set_hyperlink',
    'set_note',
    'add_defined_name',
    'delete_defined_name',
    'add_table',
    'add_pivot',
  ],
  visuals: ['add_chart', 'edit_chart', 'add_image', 'add_shape', 'add_sparkline'],
  structure: [
    'insert_rows',
    'delete_rows',
    'insert_cols',
    'delete_cols',
    'add_sheet',
    'delete_sheet',
    'rename_sheet',
    'duplicate_sheet',
    'move_sheet',
    'set_sheet_hidden',
  ],
}

const GROUP_SUMMARIES: Record<string, string> = {
  content: 'valores de célula, fórmulas, preenchimentos, localizar/substituir, classificação',
  format: 'formatação de células, formatação condicional, validação de dados',
  layout: 'mesclagens, tamanhos, linhas/colunas ocultas, painéis, filtros, configuração de página, proteção',
  data: 'hiperlinks, notas, nomes definidos, tabelas do Excel, tabelas dinâmicas',
  visuals: 'gráficos, imagens, formas, minigráficos',
  structure: 'linhas, colunas e planilhas; as operações seguintes no lote usam os endereços deslocados',
  'editor-only': 'aceitas pelo aplicativo, recusadas sem interface',
}

const NOTES: Record<string, string> = {
  set_cell:
    'um valor de texto que começa com "=" é gravado como fórmula; type: "text" mantém essa string como texto literal da célula',
  set_range: 'start é a célula superior esquerda; range, quando informado no lugar, deve corresponder à grade de valores',
  format_range:
    'null limpa uma propriedade; as cores são #RRGGBB ou slots de tema ("accent1", "accent1+40%", "dk2-25%", {theme, tint}) que seguem o tema da pasta de trabalho; fill = {pattern, fg, bg?} ou {gradient: {angle?, stops}}; border type none remove todas as bordas',
  add_conditional_format:
    'adiciona às regras que a planilha já possui (use clear_conditional_formats primeiro para recomeçar); format = {fillColor?, fontColor?, bold?, italic?}',
  clear_conditional_formats:
    'remove todas as regras da planilha; operações add_conditional_format posteriores no lote recriam do zero',
  set_data_validation:
    'adiciona às regras que a planilha já possui; em um intervalo que já tem uma, a substitui; null a remove',
  set_freeze: '0/0 desafixa o congelamento',
  set_filter_criteria: 'precisa de um filtro sem critérios; null limpa a coluna',
  set_page_setup:
    'paperSize 9 = A4, 1 = Letter; scale e fitToWidth/fitToHeight são exclusivos; margins normal|wide|narrow; printArea/printTitles/header/footer com null limpa; rowBreaks/colBreaks substituem as quebras manuais',
  add_pivot:
    'a primeira linha de sourceRange = nomes dos campos; a saída é gravada como células em targetCell (deve estar vazia) mais uma tabela dinâmica nativa que o Excel pode atualizar; targetSheet já deve existir no arquivo; não pode estar em um lote com operações de planilha ou de linha/coluna',
  add_sparkline:
    'um minigráfico por linha de dataRange (no máximo 200 linhas por operação), na coluna à direita, a menos que targetCell; type line|column|stacked',
  move_range:
    'recortar → colar em uma única operação: copia a origem (valores e fórmulas) para o destino e depois limpa as células de origem; sourceSheet para outra planilha; origem e destino não podem se sobrepor',
  convert_to_values:
    'substitui as fórmulas em range pelos valores que o mecanismo da pasta de trabalho calcula; células que resultam em erro mantêm a fórmula',
  set_col_width: 'px',
  set_row_height: 'pt',
  add_defined_name: 'ref como "Sheet1!$A$2:$A$9"; não pode estar em um lote com operações de planilha ou de linha/coluna',
  delete_defined_name: 'não pode estar em um lote com operações de planilha ou de linha/coluna',
  add_table:
    'a primeira linha = nomes de coluna únicos e não vazios; não pode estar em um lote com operações de linha/coluna, mesclagem ou tamanho na mesma planilha',
  add_chart:
    'a linha de cabeçalho e uma coluna de categoria inicial são detectadas; o gráfico fica duas colunas à direita dos dados, a menos que anchorCell',
  edit_chart:
    '`genoffice sheet read` lista os caminhos dos gráficos em features.charts; redirecionar uma série para outras células requer o aplicativo',
  add_image: 'path é um arquivo local (absoluto ou relativo) ou uma URL https',
  add_shape: 'shapeType "textbox" para uma caixa de texto',
  sort_range: 'hasHeader mantém a primeira linha no lugar',
  duplicate_sheet: 'a cópia é clonada do arquivo; edições de célula do mesmo lote não estão nela',
}

const HEADER = [
  'Operações do Excel (genoffice sheet apply --ops): um array JSON; cada entrada tem "op". Enderece uma planilha com',
  '"sheet": "<name>" (omitido = a planilha ativa; `genoffice sheet read` lista os nomes). Intervalos são A1:D9,',
  'as linhas começam em 1, as colunas são letras. Strings que começam com "=" são fórmulas.',
  'Notação dos campos: sem sufixo = string, n = número, bool = booleano, ? = opcional, a|b = uma dentre.',
]

const FOOTER = [
  'As operações executam em ordem: após inserir/excluir linhas ou colunas, as operações seguintes endereçam a grade deslocada, e uma',
  'planilha adicionada ou renomeada antes no lote é endereçada pelo novo nome. Regras, links, notas, filtros e',
  'elementos visuais devem vir depois de qualquer operação de linha/coluna de sua planilha; operações de nome definido não podem compartilhar um lote com operações de planilha',
  'ou de linha/coluna. Fórmulas gravadas por um lote são avaliadas pelo mecanismo da pasta de trabalho e armazenadas',
  'com seus resultados; funções que o mecanismo não possui ficam para o Excel calcular na abertura (aviso formulas_not_cached),',
  'e uma fórmula que o mecanismo não consegue interpretar é reportada separadamente (aviso formula_errors), para que uma referência inválida não seja',
  'confundida com uma função ausente.',
]

function opSchemas(): Map<string, JsonSchema> {
  const json = z.toJSONSchema(workbookOperationSchema, {
    io: 'input',
    unrepresentable: 'any',
  }) as JsonSchema
  const out = new Map<string, JsonSchema>()
  for (const variant of json.oneOf ?? json.anyOf ?? []) {
    const name = variant.properties?.op?.const
    if (typeof name === 'string') out.set(name, cliSurface(variant))
  }
  return out
}

/** Fields the app's schema accepts but `sheet apply` rejects headless. */
const CLI_UNSUPPORTED_FIELDS: Record<string, readonly string[]> = {
  edit_chart: ['seriesData'],
}

/** The app addresses sheets by id; the CLI takes the name, defaulting to the active sheet. */
function cliSurface(schema: JsonSchema, nested = false): JsonSchema {
  const op = schema.properties?.op?.const
  const dropped = typeof op === 'string' ? (CLI_UNSUPPORTED_FIELDS[op] ?? []) : []
  const properties: Record<string, JsonSchema> = {}
  for (const [key, value] of Object.entries(schema.properties ?? {})) {
    if (dropped.includes(key)) continue
    if (key === 'sheetId') {
      properties.sheet = {
        type: 'string',
        description: nested ? 'nome da planilha' : 'nome da planilha; omitido = a planilha ativa',
      }
    } else if (key === 'sourceSheetId') {
      properties.sourceSheet = {
        type: 'string',
        description: 'planilha do intervalo de origem; omitido = sheet',
      }
    } else if (key === 'targetSheetId') {
      properties.targetSheet = {
        type: 'string',
        description: 'planilha que recebe a saída; omitido = sheet',
      }
    } else if (value.type === 'array' && value.items?.properties) {
      properties[key] = { ...value, items: cliSurface(value.items, true) }
    } else properties[key] = value
  }
  return {
    ...schema,
    properties,
    required: (schema.required ?? [])
      .filter((k) => k !== 'sheetId' && !dropped.includes(k))
      .map((k) =>
        k === 'sourceSheetId' ? 'sourceSheet' : k === 'targetSheetId' ? 'targetSheet' : k,
      ),
  }
}

let cached: OpCatalog | undefined

export function sheetCatalog(): OpCatalog {
  if (cached) return cached
  const schemas = opSchemas()
  const ops: OpEntry[] = []
  const groups = Object.entries(SHEET_OP_GROUPS).map(([name, names]) => {
    for (const op of names) {
      const schema = schemas.get(op)
      if (!schema) throw new Error(`sheet op catalog: ${op} has no schema`)
      ops.push({
        op,
        group: name,
        signature: signatureFromSchema(schema),
        ...(NOTES[op] ? { note: NOTES[op] } : {}),
        schema,
      })
    }
    return { name, summary: GROUP_SUMMARIES[name], ops: [...names] }
  })
  const refused = Object.keys(REFUSED_DSL_OPS)
  for (const op of refused) {
    const schema = schemas.get(op)
    ops.push({
      op,
      group: 'editor-only',
      signature: schema ? signatureFromSchema(schema) : '{}',
      available: false,
      reason: REFUSED_DSL_OPS[op]!,
    })
  }
  groups.push({ name: 'editor-only', summary: GROUP_SUMMARIES['editor-only'], ops: refused })
  cached = withFingerprint({ domain: 'sheets', groups, ops })
  return cached
}

export function sheetGuideText(group?: string): string {
  const lines = [...HEADER, '', ...renderGroups(sheetCatalog(), group)]
  if (!group) lines.push('', ...FOOTER)
  return lines.join('\n')
}
