import type { AgentToolDef } from './types'
import type { AgentSkill, ExecutedToolCall } from './skill'
import { kbCiteHref } from './knowledge-base-skill'

/**
 * Lets a chat that has no open document write its result to disk: the model
 * drafts the content as markdown and create_document lands it as a .md or
 * .docx in the user's save folder. The skill is storage-agnostic — the host
 * injects the one write operation — so it pairs with the knowledge-base skill
 * on the home screen today and can back a "save as new file" in the editor
 * panels later.
 *
 * The created path is handed back as a filenav cite link (the same primitive
 * the knowledge-base citations use) so the answer can offer a click-to-open
 * link; createdPaths() lets a sibling skill's citation guard accept it.
 */

export type CreateDocumentFormat = 'md' | 'docx'

export interface CreateDocumentRequest {
  title: string
  format: CreateDocumentFormat
  /** the whole document as markdown */
  content: string
}

export interface CreateDocumentDeps {
  create(request: CreateDocumentRequest): Promise<{ path: string }>
  /** sibling tools whose success also makes "I saved it" a true statement */
  alsoSaves?: readonly string[]
}

export interface CreateDocumentSkill extends AgentSkill {
  /** absolute paths this skill wrote during the session */
  createdPaths(): readonly string[]
}

export const CREATE_DOCUMENT_TOOL = 'create_document'

/** keep in step with the main-process gate, which refuses larger bodies */
export const CREATE_DOCUMENT_MAX_CHARS = 2_000_000

const CREATE_SYSTEM_PROMPT = `## Criando documentos
Quando o usuário pedir um arquivo — um relatório, resumo, memorando, notas, "salve isso",
"exporte isso", "faça um documento / markdown / arquivo do Word" — escreva o conteúdo
completo e chame create_document; não cole o documento no chat e
não descreva o que o arquivo conteria.
- Escreva o conteúdo em Markdown (títulos, listas, tabelas, negrito). format "docx"
  gera um documento do Word, "md" um arquivo Markdown; escolha o que o usuário pediu,
  usando "docx" como padrão quando ele disser apenas "documento" e "md" quando disser
  Markdown.
- Mantenha as citações da base de conhecimento dentro do documento como texto simples (nome
  do arquivo e pasta), não como links de citação.
- A ferramenta retorna o caminho salvo com um link de citação: responda de forma breve e vincule o
  arquivo como [nome do arquivo](link de citação), copiando o link exatamente.
- Só afirme que um arquivo foi salvo quando create_document retornar ok.`

// save-only phrasing: "created" / "generated" also describe corpus facts
// ("the file was created in 2019"), so they are not claims here
const CLAIMS_SAVED =
  /\bI(?:'ve| have)? (?:saved|exported) (?:it|this|that|the|a|an|your)\b|\b(?:has been|have been|is now) saved\b|\bsaved (?:it |the (?:file|document) )?(?:as|to|in) \S+\.(?:md|docx)\b/i

export function createDocumentSkill(deps: CreateDocumentDeps): CreateDocumentSkill {
  const created: string[] = []
  const tool: AgentToolDef = {
    name: CREATE_DOCUMENT_TOOL,
    description:
      'Salva um novo documento na pasta de salvamento do usuário a partir de conteúdo Markdown. ' +
      'format "docx" grava um documento do Word, "md" um arquivo Markdown. ' +
      'Retorna o caminho salvo e seu link de citação; vincule-o na resposta.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'nome do arquivo sem extensão' },
        format: { type: 'string', enum: ['md', 'docx'], description: 'formato de saída' },
        content: { type: 'string', description: 'o documento completo em Markdown' },
      },
      required: ['title', 'format', 'content'],
    },
  }
  return {
    id: 'create-document',
    systemPrompt: CREATE_SYSTEM_PROMPT,
    tools: [tool],
    createdPaths: () => created,
    executeTool: async (call) => {
      if (call.name !== CREATE_DOCUMENT_TOOL) {
        return { output: `ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
      }
      const title = typeof call.input.title === 'string' ? call.input.title.trim() : ''
      const format = call.input.format
      const content = typeof call.input.content === 'string' ? call.input.content : ''
      if (format !== 'md' && format !== 'docx') {
        return { output: 'format deve ser "md" ou "docx"', isError: true, summary: 'formato inválido' }
      }
      if (!content.trim()) {
        return { output: 'content não pode estar vazio', isError: true, summary: 'conteúdo vazio' }
      }
      if (content.length > CREATE_DOCUMENT_MAX_CHARS) {
        return {
          output: `content é longo demais (${content.length} caracteres, limite ${CREATE_DOCUMENT_MAX_CHARS}); divida-o em vários documentos`,
          isError: true,
          summary: 'conteúdo longo demais',
        }
      }
      try {
        const { path } = await deps.create({ title: title || 'Sem título', format, content })
        created.push(path)
        const name = path.split(/[\\/]/).pop() ?? path
        return {
          output: `salvo ${name}\n  caminho: ${path}\n  cite: ${kbCiteHref(path)}`,
          summary: name,
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error)
        return {
          output: `não foi possível salvar o documento: ${message}`,
          isError: true,
          summary: 'falha ao salvar',
        }
      }
    },
    verifyResponse: (finalText, executed: readonly ExecutedToolCall[]) => {
      const savers = [CREATE_DOCUMENT_TOOL, ...(deps.alsoSaves ?? [])]
      const saved = executed.some((c) => savers.includes(c.name) && c.ok)
      if (saved || !CLAIMS_SAVED.test(finalText)) return null
      return (
        'A resposta diz que um arquivo foi salvo, mas create_document não teve sucesso nesta execução. ' +
        'Diga claramente que nada foi salvo; chame create_document apenas se o usuário pediu um arquivo.'
      )
    },
  }
}
