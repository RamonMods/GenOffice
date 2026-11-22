import type { AgentSkill } from '@genoffice/agent-core'
import type { AttachmentMeta } from '../../shared/desktop-api'
import { ATTACHMENT_IMAGE_EXTS } from '../../shared/desktop-api'
import { t } from '../i18n/locale'

/**
 * Chat-attachment capability as an AgentSkill (same structure as the files-skill
 * in apps/docs and apps/slides): each turn's context lists the attached local
 * files, and read_attachment reads their extracted text in pages (parsing
 * happens in the main process; files never leave the machine).
 */

const READ_CHUNK_CHARS = 24_000

const FILES_SYSTEM_PROMPT = `## Attachments
The user may attach local files to the conversation (see the "attachment list" in each turn's context).
- When the user's request involves attachment content, read it with read_attachment first, then answer or generate; never guess the content from the file name.
- Long files are read in pages: the result reports the total character count and the current range; to continue, set offset to the previous chunk's end position.
- Image attachments (png/jpg/gif/webp) are already sent as images with the user message — just look at them; read_attachment is only for text-like attachments.
- Do not call read_attachment when there are no attachments or they are unrelated to the request.`

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)}KB`
  return `${(bytes / 1024 / 1024).toFixed(1)}MB`
}

export function createFilesSkill(getAttachments: () => readonly AttachmentMeta[]): AgentSkill {
  return {
    id: 'files',
    systemPrompt: FILES_SYSTEM_PROMPT,
    tools: [
      {
        name: 'read_attachment',
        description:
          'Ler o conteúdo de texto de um anexo (analisado localmente). Arquivos longos são paginados: leia com offset=0 primeiro e depois decida se continua com base no total de caracteres retornado.',
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'integer',
              description: 'Índice do anexo (base 0, veja a lista de anexos)',
            },
            offset: { type: 'integer', description: 'Posição inicial do caractere, padrão 0' },
          },
          required: ['index'],
        },
      },
    ],
    buildContext: () => {
      const list = getAttachments()
      if (list.length === 0) return ''
      const lines = list.map((a, i) => `${i} | ${a.name} | .${a.ext} | ${formatSize(a.sizeBytes)}`)
      return `Attachment list (index | file name | type | size):\n${lines.join('\n')}`
    },
    executeTool: async (call) => {
      if (call.name !== 'read_attachment') {
        return { output: `Ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
      }
      const list = getAttachments()
      const index = Number(call.input.index)
      const att = Number.isInteger(index) ? list[index] : undefined
      if (!att) {
        return {
          output: 'Índice de anexo inválido (veja a lista de anexos)',
          isError: true,
          summary: t('aiToolReadAttachment'),
        }
      }
      // Images skip text extraction: already provided as multimodal images with
      // the user message on send
      if (ATTACHMENT_IMAGE_EXTS.has(att.ext)) {
        return {
          output: `${att.name} é um anexo de imagem, já enviado como imagem junto com a mensagem do usuário — basta olhar a imagem na mensagem; não há texto para ler.`,
          mutated: false,
          summary: t('aiToolImageAttachment', { name: att.name }),
        }
      }
      const offset = Math.max(0, Number(call.input.offset) || 0)
      const result = await window.desktopApi.readAttachment(att.path, offset, READ_CHUNK_CHARS)
      if (!result.ok) {
        return {
          output: result.error ?? 'Falha na leitura',
          isError: true,
          summary: t('aiToolReadFile', { name: att.name }),
        }
      }
      const end = (result.offset ?? 0) + (result.text?.length ?? 0)
      const header = `Arquivo ${att.name}, total de caracteres ${result.totalChars}, este trecho ${result.offset}-${end}${
        end < (result.totalChars ?? 0)
          ? ' (não terminou; continue com offset=' + end + ')'
          : ' (fim do arquivo)'
      }`
      return {
        output: `${header}\n---\n${result.text ?? ''}`,
        mutated: false,
        summary: t('aiToolReadAttachmentOf', { name: att.name }),
      }
    },
  }
}
