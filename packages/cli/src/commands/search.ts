import { imageSearchTool, webSearchTool } from '@genoffice/ai-search'
import { flagBool, flagString } from '../args'
import { aiSettingsPath, prepareCloud } from '../cloud'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

export const searchCommand: CommandDef = {
  name: 'search',
  summary:
    'Pesquisa na web ou por imagens pelo provedor configurado no GenOffice (RamonMods, Serper, Serply, SearchApi, Tavily, Parallel).',
  usage: 'search <query> [--images] [--max <n>]',
  options: [
    { name: 'images', description: 'pesquisa imagens em vez de páginas da web' },
    { name: 'max', value: 'n', description: 'quantidade de resultados (padrão 6 para web, 8 para imagens; no máximo 20)' },
  ],
  async run(args, ctx) {
    const query = args.positionals.join(' ').trim()
    if (!query)
      throw new CliError(EXIT.usage, 'falta o argumento <query>', undefined, { reason: 'missing_argument' })
    const max = resultCount(flagString(args, 'max'))
    await prepareCloud(ctx.env)
    const settings = aiSettingsPath(ctx.env)
    if (flagBool(args, 'images')) {
      const r = await imageSearchTool(settings, query, max ?? 8)
      if (r.method === 'error') throw new CliError(EXIT.app, r.error ?? 'a pesquisa de imagens falhou')
      return {
        summary: `${r.images.length} imagens para "${query}" via ${r.method}`,
        detail: { method: r.method, images: r.images },
      }
    }
    const r = await webSearchTool(settings, query, max ?? 6)
    if (r.method === 'error') throw new CliError(EXIT.app, r.error ?? 'a pesquisa falhou')
    return {
      summary: `${r.results.length} resultados para "${query}" via ${r.method}`,
      detail: {
        method: r.method,
        ...(r.answer ? { answer: r.answer } : {}),
        results: r.results,
      },
    }
  },
}

/** `--max` clamped to 1..20; undefined when absent so each mode keeps its own default. */
export function resultCount(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined
  const n = Number(raw)
  if (!Number.isFinite(n))
    throw new CliError(EXIT.usage, `--max deve ser um número, recebido "${raw}"`, undefined, {
      reason: 'invalid_argument',
    })
  return Math.min(20, Math.max(1, Math.floor(n)))
}
