import writingGuide from './prompts/guides/writing.md?raw'
import formattingGuide from './prompts/guides/formatting.md?raw'
import layoutGuide from './prompts/guides/layout.md?raw'
import structureGuide from './prompts/guides/structure.md?raw'
import chartsGuide from './prompts/guides/charts.md?raw'
import dataGuide from './prompts/guides/data.md?raw'
import financialFormattingGuide from './prompts/guides/financial-formatting.md?raw'
import dataAttributionGuide from './prompts/guides/data-attribution.md?raw'
import tableGuide from './prompts/guides/table.md?raw'
import pivotGuide from './prompts/guides/pivot.md?raw'
import shapeImageGuide from './prompts/guides/shape-image.md?raw'

/**
 * On-demand prompt guides: the base system prompt stays small, the load_guide
 * tool's description carries this catalog (name + one-line summary), and the
 * model loads the domains it actually needs for the task at hand.
 */

export interface GuideEntry {
  readonly description: string
  readonly content: string
}

export const GUIDE_CATALOG: Readonly<Record<string, GuideEntry>> = {
  writing: {
    description:
      'Escrita de conteúdo: definições de campos de set_cell/set_formula/set_range/clear_*, regras de fórmulas, colunas de participação e convenções de valor/unidade',
    content: writingGuide,
  },
  formatting: {
    description:
      'Formatação: todos os campos de format_range (incl. bordas), formatos de número, regras de estilo de tabela (cabeçalhos/alinhamento/linhas de total/paleta financeira)',
    content: formattingGuide,
  },
  layout: {
    description:
      'Ordenação e layout: semântica e limites de sort_range, mesclar/desmesclar, alturas de linha/larguras de coluna, ocultar linhas/colunas',
    content: layoutGuide,
  },
  structure: {
    description:
      'Alterações estruturais: inserção/exclusão de linhas e colunas, adicionar/excluir/duplicar/mover/ocultar planilhas, proteção de planilha, disciplina de dois lotes, reescrita automática de referências de fórmulas e #REF!',
    content: structureGuide,
  },
  charts: {
    description:
      'Gráficos e formas: add_chart para criar gráficos a partir de intervalos de dados, edit_chart para editar gráficos existentes, add_shape para inserir formas/caixas de texto',
    content: chartsGuide,
  },
  data: {
    description:
      'Ferramentas de dados: hiperlinks, AutoFiltro, formatação condicional, validação de dados, nomes definidos',
    content: dataGuide,
  },
  'financial-formatting': {
    description:
      'Formatação de demonstrações financeiras: fixação da localidade da moeda, parênteses vermelhos para negativos, traço para zeros, recuo hierárquico, linhas de total, ordem de rubricas GAAP, paleta contábil',
    content: financialFormattingGuide,
  },
  'data-attribution': {
    description:
      'Atribuição de fontes de dados: três métodos de rastreamento para preencher planilhas com dados externos (busca/API/raspagem) — coluna Fonte / linhas de rodapé / planilha Fontes dedicada — mais a tabela de decisão',
    content: dataAttributionGuide,
  },
  table: {
    description:
      'Tabelas estruturadas: add_table cria objetos Table nativos (estilos/faixas/filtro integrados) + add/delete_table_row/column para editar linhas e colunas',
    content: tableGuide,
  },
  pivot: {
    description:
      'Tabelas dinâmicas: add_pivot cria PivotTables nativas (OOXML real, reconhecidas e atualizáveis no Excel) + refresh_pivot recalcula tabelas dinâmicas existentes',
    content: pivotGuide,
  },
  'shape-image': {
    description:
      'Formas e imagens: add_shape para formas/caixas de texto, SPARKLINE para minigráficos em células, IMAGE/HYPERLINK incorporados',
    content: shapeImageGuide,
  },
}

export const GUIDE_NAMES = Object.keys(GUIDE_CATALOG)

export function guideCatalogSummary(): string {
  return Object.entries(GUIDE_CATALOG)
    .map(([name, entry]) => `${name} — ${entry.description}`)
    .join('; ')
}

export function loadGuides(
  names: readonly string[],
): { ok: true; content: string } | { ok: false; error: string } {
  const unknown = names.filter((name) => !GUIDE_CATALOG[name])
  if (unknown.length > 0) {
    return {
      ok: false,
      error: `Guia(s) desconhecido(s): ${unknown.join(', ')}. Disponíveis: ${GUIDE_NAMES.join(', ')}`,
    }
  }
  const content = names.map((name) => GUIDE_CATALOG[name]!.content.trim()).join('\n\n---\n\n')
  return { ok: true, content }
}
