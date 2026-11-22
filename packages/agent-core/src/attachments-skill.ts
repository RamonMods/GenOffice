import type { AgentSkill } from './skill'
import { formatKbSize, kbCiteHref } from './knowledge-base-skill'

/**
 * Chat attachments as an AgentSkill: the per-turn context lists the files the
 * user attached, and read_attachment pages through their extracted text. The
 * parsing lives in the main process; deps inject it, so the skill is shared by
 * the editors' panels and the home screen's Ask AI without an electron import.
 */

export const ATTACHMENT_IMAGE_EXTS: ReadonlySet<string> = new Set([
  'png',
  'jpg',
  'jpeg',
  'gif',
  'webp',
])

export const READ_ATTACHMENT_TOOL = 'read_attachment'
export const ATTACHMENT_READ_CHUNK_CHARS = 24_000

export interface AttachmentInfo {
  /** absolute local path; the file never leaves the machine */
  path: string
  name: string
  /** lowercased extension without the dot */
  ext: string
  sizeBytes: number
}

export interface AttachmentPage {
  text: string
  totalChars: number
  offset: number
}

export interface AttachmentsSkillDeps {
  list(): readonly AttachmentInfo[]
  /** one slice of the parsed text; rejects with the user-facing reason */
  read(path: string, offset: number, maxChars: number): Promise<AttachmentPage>
  /** list absolute paths and cite links too (home: the CLI tools take paths, answers cite files) */
  pathsInContext?: boolean
  /** tool-row labels shown in the panel */
  labels?: {
    read?: (name: string) => string
    image?: (name: string) => string
    invalid?: string
  }
}

const ATTACHMENTS_SYSTEM_PROMPT = `## Anexos
O usuário pode anexar arquivos locais à conversa (veja a "lista de anexos" no contexto de cada turno).
- Quando a solicitação do usuário envolver o conteúdo de um anexo, leia-o com read_attachment primeiro e só então responda ou escreva; não adivinhe o conteúdo pelo nome do arquivo.
- Arquivos longos são lidos em páginas: o resultado informa a contagem total de caracteres e o intervalo atual; para continuar, defina offset como a posição final da fatia anterior.
- Anexos de imagem (png/jpg/gif/webp) já são enviados como imagens junto com a mensagem do usuário — basta olhá-los; read_attachment serve apenas para anexos de texto.
- Não chame read_attachment quando não houver anexos ou quando eles não tiverem relação com a solicitação.`

const PATHS_PROMPT = `
- Os anexos são listados com seus caminhos absolutos: ferramentas que recebem um caminho de arquivo (conversão, criação a partir de um modelo) podem usá-los diretamente, e uma afirmação extraída de um anexo é citada com seu link de citação, como ocorre com um arquivo da base de conhecimento.`

export function createAttachmentsSkill(deps: AttachmentsSkillDeps): AgentSkill {
  const labels = deps.labels ?? {}
  return {
    id: 'files',
    systemPrompt: deps.pathsInContext
      ? ATTACHMENTS_SYSTEM_PROMPT + PATHS_PROMPT
      : ATTACHMENTS_SYSTEM_PROMPT,
    tools: [
      {
        name: READ_ATTACHMENT_TOOL,
        description:
          'Lê o conteúdo de texto de um anexo (processado localmente). Arquivos longos são paginados: leia com offset=0 primeiro e decida se continua com base na contagem total de caracteres retornada.',
        inputSchema: {
          type: 'object',
          properties: {
            index: {
              type: 'integer',
              description: 'índice do anexo (base 0, veja a lista de anexos)',
            },
            offset: { type: 'integer', description: 'posição inicial do caractere, padrão 0' },
          },
          required: ['index'],
        },
      },
    ],
    buildContext: () => {
      const list = deps.list()
      if (list.length === 0) return ''
      const lines = list.map(
        (a, i) => `${i} | ${a.name} | .${a.ext} | ${formatKbSize(a.sizeBytes)}`,
      )
      const head = `Lista de anexos (índice | nome do arquivo | tipo | tamanho):\n${lines.join('\n')}`
      if (!deps.pathsInContext) return head
      const paths = list.map((a, i) => `${i} = ${a.path}\n    cite: ${kbCiteHref(a.path)}`)
      return `${head}\n\nCaminhos absolutos e links de citação:\n${paths.join('\n')}`
    },
    executeTool: async (call) => {
      if (call.name !== READ_ATTACHMENT_TOOL) {
        return { output: `ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
      }
      const list = deps.list()
      const index = Number(call.input.index)
      const att = Number.isInteger(index) ? list[index] : undefined
      if (!att) {
        return {
          output: 'índice de anexo inválido (veja a lista de anexos)',
          isError: true,
          summary: labels.invalid ?? 'ler anexo',
        }
      }
      if (ATTACHMENT_IMAGE_EXTS.has(att.ext)) {
        return {
          output: `${att.name} é um anexo de imagem já enviado como imagem junto com a mensagem do usuário; basta olhar a imagem na mensagem, não há texto para ler.`,
          mutated: false,
          summary: labels.image?.(att.name) ?? att.name,
        }
      }
      const offset = Math.max(0, Number(call.input.offset) || 0)
      let page: AttachmentPage
      try {
        page = await deps.read(att.path, offset, ATTACHMENT_READ_CHUNK_CHARS)
      } catch (error) {
        return {
          output: error instanceof Error ? error.message : String(error),
          isError: true,
          summary: labels.read?.(att.name) ?? att.name,
        }
      }
      const end = page.offset + page.text.length
      const header = `Arquivo ${att.name}, total de caracteres ${page.totalChars}, esta fatia ${page.offset}-${end}${
        end < page.totalChars
          ? ' (não terminou, continue com offset=' + end + ')'
          : ' (fim do arquivo)'
      }`
      return {
        output: `${header}\n---\n${page.text}`,
        mutated: false,
        summary: labels.read?.(att.name) ?? att.name,
      }
    },
  }
}
