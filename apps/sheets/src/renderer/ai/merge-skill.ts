/**
 * AI-driven workbook merging: the user
 * attaches spreadsheet files in the chat and asks the assistant to combine
 * them. The heavy lifting stays engine-side — the tool imports every sheet of
 * the chosen attachments into the current workbook through the same pipeline
 * as Data ▸ Merge Workbooks, and the model then works on the merged data with
 * the normal workbook tools (and create_document for new output files).
 */
import type { AgentSkill } from '@genoffice/agent-core'
import type { AttachmentMeta } from '../../shared/desktop-api'
import type { MergeSourcesResult } from '../merge-workbooks'

export const SPREADSHEET_ATTACHMENT_EXTS = new Set(['xlsx', 'xlsm', 'xls', 'csv', 'tsv'])

const MERGE_SYSTEM_PROMPT = `## Mesclando planilhas anexadas
Quando o usuário anexar arquivos de planilha (xlsx/xlsm/xls/csv/tsv) e pedir para mesclar, combinar, consolidar ou processá-los juntos:
- Chame merge_attached_workbooks PRIMEIRO. Ele importa todas as planilhas dos anexos escolhidos para a pasta de trabalho atual no mecanismo (rápido, exato) e informa os novos nomes de planilhas. Nunca reconstrua o conteúdo das planilhas a partir do texto de read_attachment para mesclar — esse caminho tem perdas e é lento.
- Depois disso os dados JÁ estão na pasta de trabalho atual: use get_workbook_context / read_range / aggregate_range nas planilhas informadas e create_document para gerar novos arquivos se o usuário quiser saídas separadas.
- Fórmulas nas origens chegam como seus valores calculados.`

/** attachment indexes that are spreadsheets (the tool's default selection) */
export function spreadsheetAttachmentIndexes(list: readonly AttachmentMeta[]): number[] {
  const indexes: number[] = []
  for (let index = 0; index < list.length; index += 1) {
    const attachment = list[index]
    if (attachment && SPREADSHEET_ATTACHMENT_EXTS.has(attachment.ext)) indexes.push(index)
  }
  return indexes
}

export interface MergeSkillDeps {
  getAttachments(): readonly AttachmentMeta[]
  mergePaths(paths: string[]): Promise<MergeSourcesResult>
}

export function createMergeSkill(deps: MergeSkillDeps): AgentSkill {
  return {
    id: 'workbook-merge',
    systemPrompt: MERGE_SYSTEM_PROMPT,
    tools: [
      {
        name: 'merge_attached_workbooks',
        description:
          'Importa todas as planilhas dos arquivos de planilha anexados (xlsx/xlsm/xls/csv) para a pasta de trabalho atual, no mecanismo. ' +
          'Retorna os nomes das planilhas criadas. Use isto antes de analisar ou reestruturar dados de planilhas anexadas; ' +
          'nunca reconstrua o conteúdo dos anexos célula por célula a partir de texto.',
        inputSchema: {
          type: 'object',
          properties: {
            indexes: {
              type: 'array',
              items: { type: 'integer' },
              description:
                'Índices dos anexos a mesclar (base 0, da lista de anexos). Omita para mesclar todos os anexos de planilha.',
            },
          },
          required: [],
        },
      },
    ],
    executeTool: async (call) => {
      if (call.name !== 'merge_attached_workbooks') {
        return { output: `Ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
      }
      const list = deps.getAttachments()
      const requested = Array.isArray((call.input as { indexes?: unknown }).indexes)
        ? ((call.input as { indexes: unknown[] }).indexes.filter(
            (value) => Number.isInteger(value) && (value as number) >= 0,
          ) as number[])
        : spreadsheetAttachmentIndexes(list)
      const paths: string[] = []
      const skipped: string[] = []
      for (const index of requested) {
        const attachment = list[index]
        if (!attachment) {
          skipped.push(`#${index} (anexo inexistente)`)
        } else if (!SPREADSHEET_ATTACHMENT_EXTS.has(attachment.ext)) {
          skipped.push(`${attachment.name} (não é uma planilha)`)
        } else {
          paths.push(attachment.path)
        }
      }
      if (paths.length === 0) {
        return {
          output:
            'Não há anexos de planilha para mesclar. Peça ao usuário para anexar arquivos xlsx/xlsm/xls/csv primeiro.',
          isError: true,
          summary: 'mesclar: sem fontes',
        }
      }
      try {
        const result = await deps.mergePaths(paths)
        const lines = [
          `${result.importedSheets} planilhas mescladas de ${result.files} arquivos na pasta de trabalho atual.`,
          `Novas planilhas: ${result.sheetNames.join(', ')}`,
          'Fórmulas das origens importadas como valores calculados.',
        ]
        if (skipped.length > 0) lines.push(`Ignorados: ${skipped.join('; ')}`)
        return {
          output: lines.join('\n'),
          summary: `${result.importedSheets} planilhas mescladas`,
        }
      } catch (error: unknown) {
        return {
          output: `Falha na mesclagem: ${error instanceof Error ? error.message : String(error)}`,
          isError: true,
          summary: 'falha na mesclagem',
        }
      }
    },
  }
}
