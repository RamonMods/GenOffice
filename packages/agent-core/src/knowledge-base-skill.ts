import type { AgentToolDef } from './types'
import type { AgentSkill } from './skill'

/**
 * The reader's starred files (the star corpus) as a searchable, readable body.
 *
 * The skill is storage-agnostic on purpose: deps inject the three main-process
 * operations (list / search / read), so it carries no electron or index
 * dependency and is testable with mocks. The shell's knowledge-base gateway
 * (apps/shell/src/main/knowledge-base.ts) is the intended backing; the docs
 * panel is the first consumer.
 *
 * Citation contract: when an answer draws on corpus content it must cite the
 * source as [file name](filenav:///abs/path) — the same primitive the editors'
 * in-document citations use (see @genoffice/ui file-nav). verifyResponse
 * rejects answers citing paths outside the corpus, the way fabricated
 * block numbers are rejected elsewhere.
 */

/** characters of corpus listing injected before it degrades to names only */
const KB_CONTEXT_MAX_CHARS = 12_000

export interface KbFileInfo {
  /** absolute path — also the citation target */
  path: string
  /** owning folder name ("contracts" in /Users/me/contracts/lease.docx) */
  folder: string
  name: string
  ext: string
  sizeBytes: number
}

export interface KbSearchHit {
  path: string
  name: string
  ext: string
  folder: string
  /** plain text around the first match (marks stripped) */
  snippet: string
}

export interface KbReadPage {
  text: string
  /** total parsed characters — page on with offset when more remains */
  totalChars: number
  offset: number
}

export interface KnowledgeBaseDeps {
  /** the corpus as of this turn (a synchronous snapshot; the panel refreshes it on the starred-list IPC) */
  listFiles(): KbFileInfo[]
  search(query: string, limit: number): Promise<KbSearchHit[]>
  /** one page of parsed text; path must be a corpus file */
  read(path: string, offset: number): Promise<KbReadPage>
  /** paths outside the corpus an answer may still cite (files a sibling skill wrote this session) */
  alsoCitable?(): readonly string[]
}

/** keep in sync with FILE_NAV_SCHEME in @genoffice/ui (this package must not
 *  depend on @genoffice/ui: the ui barrel reaches react) */
const FILE_NAV_SCHEME = 'filenav://'

export function formatKbSize(bytes: number): string {
  if (bytes < 1024) return `${bytes}B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)}KB`
  return `${(bytes / 1024 / 1024).toFixed(1)}MB`
}

const KB_SYSTEM_PROMPT = `## Base de conhecimento
Os arquivos favoritados do usuário formam uma base de conhecimento; o contexto de cada turno
os lista (índice | pasta | arquivo | tipo | tamanho, com o caminho absoluto).
- Use-a apenas quando a solicitação se referir a esses arquivos ou ao conteúdo deles; redigir,
  reescrever ou formatar o documento aberto apenas a partir da solicitação nunca precisa dela.
- Quando a solicitação envolver a base de conhecimento, faça uma busca primeiro (search_knowledge_base)
  e leia os arquivos promissores (read_knowledge_file, paginado por offset) antes de
  responder. Nunca responda perguntas sobre o acervo usando apenas nomes de arquivos.
- Uma busca de texto completo corresponde apenas a palavras literais: se não retornar nada, tente
  sinônimos ou o outro idioma antes de
  concluir que o acervo não cobre o assunto.
- Arquivos longos são paginados: o resultado informa totalChars e o próximo offset; continue
  lendo com offset enquanto for relevante.
- Cite o que você usou: logo após uma afirmação extraída de um arquivo, escreva
  [nome do arquivo](link de citação) — copie o link de citação exatamente como a lista ou um
  resultado de busca o fornece (ele já vem codificado). Cite SOMENTE links que aparecem ali;
  nunca invente ou edite um.
- Diga claramente quando a base de conhecimento não contiver a resposta.`

/**
 * Same encoding as @genoffice/ui file-nav (agent-core must not depend on the
 * renderer package): everything the inline-link grammar rejects — spaces, %,
 * CJK, parens, Windows backslashes — is percent-encoded, slashes stay.
 */
export function kbCiteHref(path: string): string {
  const encoded = encodeURIComponent(path)
    .replace(/%2F/g, '/')
    .replace(/\(/g, '%28')
    .replace(/\)/g, '%29')
  return FILE_NAV_SCHEME + encoded
}

const KB_EMPTY =
  'A base de conhecimento está vazia (o usuário não favoritou nenhum arquivo). Não faça novas buscas; trabalhe a partir da solicitação e do documento aberto.'

export function createKnowledgeBaseSkill(deps: KnowledgeBaseDeps): AgentSkill {
  const readTool: AgentToolDef = {
    name: 'read_knowledge_file',
    description:
      'Lê uma página do texto processado de um arquivo da base de conhecimento (docx/pdf/xlsx/pptx/text são processados localmente). ' +
      'Leia com offset=0 primeiro; continue com o offset retornado enquanto totalChars indicar que resta mais.',
    inputSchema: {
      type: 'object',
      properties: {
        index: {
          type: 'integer',
          description: 'índice do arquivo na lista da base de conhecimento (base 0)',
        },
        offset: { type: 'integer', description: 'posição do caractere a partir da qual ler; padrão 0' },
      },
      required: ['index'],
    },
  }
  const searchTool: AgentToolDef = {
    name: 'search_knowledge_base',
    description:
      'Busca de texto completo na base de conhecimento (arquivos favoritados do usuário), para solicitações que se referem a esses arquivos; não serve para redigir ou formatar apenas a partir da solicitação. ' +
      'Retorna os arquivos correspondentes com um trecho ao redor da primeira ocorrência. Correspondência por palavra literal: ' +
      'tente sinônimos ou o outro idioma quando não retornar nada.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'palavras de busca' },
        limit: { type: 'integer', description: 'máximo de resultados, padrão 8, no máximo 20' },
      },
      required: ['query'],
    },
  }
  const empty = () => deps.listFiles().length === 0
  return {
    id: 'knowledge-base',
    // nothing starred: no prompt section and no tools, so a model that searches
    // on reflex (BYOK/local endpoints did, repeatedly) has nothing to call
    get systemPrompt() {
      return empty() ? '' : KB_SYSTEM_PROMPT
    },
    get tools() {
      return empty() ? [] : [readTool, searchTool]
    },
    buildContext: () => {
      const files = deps.listFiles()
      if (files.length === 0) return ''
      const plain = files.map(
        (f, i) => `${i} | ${f.folder} | ${f.name} | .${f.ext} | ${formatKbSize(f.sizeBytes)}`,
      )
      const paths = files.map((f, i) => `${i} = ${f.path}\n    cite: ${kbCiteHref(f.path)}`)
      const withPaths = `${plain.join('\n')}\n\nCaminhos absolutos e links de citação:\n${paths.join('\n')}`
      // a large corpus degrades to names only: the model searches first anyway,
      // and search results carry the real paths to cite
      const body =
        withPaths.length <= KB_CONTEXT_MAX_CHARS
          ? withPaths
          : `${plain.join('\n')}\n\n(O acervo é grande; os links de citação vêm com os resultados da busca.)`
      return `Base de conhecimento — arquivos favoritados do usuário (índice | pasta | arquivo | tipo | tamanho):\n${body}`
    },
    executeTool: async (call) => {
      if (deps.listFiles().length === 0) {
        return { output: KB_EMPTY, isError: true, summary: 'base de conhecimento vazia' }
      }
      if (call.name === 'search_knowledge_base') {
        const query = typeof call.input.query === 'string' ? call.input.query : ''
        if (!query.trim()) {
          return { output: 'query não pode estar vazia', isError: true, summary: 'search ""' }
        }
        const limitRaw = Number(call.input.limit)
        const limit = Number.isFinite(limitRaw)
          ? Math.min(Math.max(1, Math.floor(limitRaw)), 20)
          : 8
        const hits = await deps.search(query, limit)
        if (hits.length === 0)
          return { output: 'nenhuma correspondência. Tente sinônimos ou o outro idioma.', summary: query }
        const lines = hits.map(
          (h) =>
            `${h.folder}/${h.name}\n  caminho: ${h.path}\n  cite: ${kbCiteHref(h.path)}\n  trecho: ${h.snippet}`,
        )
        return { output: lines.join('\n'), summary: `${query} → ${hits.length}` }
      }
      if (call.name === 'read_knowledge_file') {
        const files = deps.listFiles()
        const index = Number(call.input.index)
        const file = Number.isInteger(index) ? files[index] : undefined
        if (!file) {
          return {
            output: `nenhum arquivo da base de conhecimento no índice ${call.input.index}`,
            isError: true,
            summary: `read #${call.input.index}`,
          }
        }
        const offsetRaw = Number(call.input.offset)
        const page = await deps.read(file.path, Number.isFinite(offsetRaw) ? offsetRaw : 0)
        const more = page.offset + page.text.length < page.totalChars
        return {
          output:
            `${file.folder}/${file.name} — caracteres ${page.offset}..${page.offset + page.text.length} de ${page.totalChars}` +
            (more ? ' (continue com offset)' : ' (fim)') +
            `\n\n${page.text}`,
          summary: `${file.name}@${page.offset}`,
        }
      }
      return { output: `ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
    },
    verifyResponse: (response) => {
      const cited = [...response.matchAll(/\]\((filenav:\/\/[^)\s]+)\)/g)].map((m) => m[1]!)
      if (cited.length === 0) return null
      // hrefs arrive percent-encoded (the renderer's own generator encodes
      // spaces/CJK); compare decoded forms so a hand-written raw href passes too
      const known = new Set([
        ...deps.listFiles().map((f) => f.path),
        ...(deps.alsoCitable?.() ?? []),
      ])
      const unknown = cited.filter((href) => {
        const raw = href.slice(FILE_NAV_SCHEME.length)
        let decoded = raw
        try {
          decoded = decodeURIComponent(raw)
        } catch {
          // keep the raw text
        }
        return !known.has(decoded)
      })
      if (unknown.length === 0) return null
      return (
        `Estas citações não são caminhos da base de conhecimento: ${unknown.join(', ')}. ` +
        'Cite apenas os links de citação fornecidos literalmente na lista da base de conhecimento ou em um resultado de busca, ou remova a citação.'
      )
    },
  }
}
