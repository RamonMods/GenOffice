import type { AgentSkill } from '@genoffice/agent-core'
import { t } from '../i18n/locale'

/**
 * Web-search AgentSkill (same source as docs/slides web_search):
 * the main process queries Serper/DuckDuckGo; the renderer only receives
 * titles/links/snippets.
 */

const SEARCH_SYSTEM_PROMPT = `## Web search
- When you need up-to-date information, data, or facts beyond the workbook, use web_search; never fabricate numbers from memory.
- When writing search results into the workbook, you must attribute the data source (load_guide: data-attribution first).`

export function createSearchSkill(): AgentSkill {
  return {
    id: 'search',
    systemPrompt: SEARCH_SYSTEM_PROMPT,
    tools: [
      {
        name: 'web_search',
        description:
          'Pesquise na web por informações textuais (referências/dados/fatos). Use quando precisar de informações atualizadas ou tiver dúvidas sobre um fato. Retorna títulos/links/trechos.',
        inputSchema: {
          type: 'object',
          properties: {
            query: { type: 'string', description: 'Palavras-chave da pesquisa' },
            maxResults: { type: 'integer', description: 'Número máximo de resultados, padrão 6' },
          },
          required: ['query'],
        },
      },
    ],
    executeTool: async (call) => {
      if (call.name !== 'web_search') {
        return { output: `Ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
      }
      const query = String(call.input.query ?? '').trim()
      if (!query) {
        return { output: 'query não deve estar vazia', isError: true, summary: t('aiToolWebSearch') }
      }
      const r = await window.desktopApi.webSearch(query, Number(call.input.maxResults) || 6)
      // a backend failure must not read as "no results" — the model would fabricate conclusions
      if (r.method === 'error') {
        return {
          output: `falha na busca na web (erro do serviço, não um resultado vazio — você pode tentar novamente): ${r.error ?? 'erro desconhecido'}`,
          isError: true,
          summary: t('aiToolWebSearch'),
        }
      }
      const lines: string[] = []
      if (r.answer) lines.push(`Resposta direta: ${r.answer}\n`)
      r.results.forEach((it, i) =>
        lines.push(`${i + 1}. ${it.title}\n   ${it.url}\n   ${it.snippet}`),
      )
      return {
        output: lines.join('\n') || '(nenhum resultado)',
        mutated: false,
        summary: t('aiToolWebSearchDone', { query, count: r.results.length }),
      }
    },
  }
}
