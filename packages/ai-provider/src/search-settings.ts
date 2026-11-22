import type {
  AiSearchProviderId,
  AiSearchProviderMeta,
  AiSearchSettings,
  AiSettings,
} from './types'

export const AI_SEARCH_PROVIDERS: AiSearchProviderMeta[] = [
  {
    id: 'ramonmods',
    label: 'RamonMods',
    keyPlaceholder: 'Não é necessário - faça login na RamonMods',
    imageSearch: true,
  },
  { id: 'serper', label: 'Serper', keyPlaceholder: 'Chave de API do Serper', imageSearch: true },
  { id: 'serply', label: 'Serply', keyPlaceholder: 'Chave de API do Serply', imageSearch: true },
  { id: 'searchapi', label: 'SearchApi', keyPlaceholder: 'Chave de API do SearchApi', imageSearch: true },
  { id: 'tavily', label: 'Tavily', keyPlaceholder: 'tvly-...', imageSearch: false },
  { id: 'parallel', label: 'Parallel', keyPlaceholder: 'Chave de API da Parallel', imageSearch: false },
  // exa/firecrawl: AI-search APIs without an image endpoint (like Tavily/Parallel)
  { id: 'exa', label: 'Exa', keyPlaceholder: 'Chave de API da Exa', imageSearch: false },
  { id: 'firecrawl', label: 'Firecrawl', keyPlaceholder: 'fc-...', imageSearch: false },
  // keenable: keyless like Parallel; a key only raises the rate limits
  { id: 'keenable', label: 'Keenable', keyPlaceholder: 'Chave de API da Keenable', imageSearch: false },
]

export function defaultAiSearchSettings(): AiSearchSettings {
  return {
    provider: 'ramonmods',
    providers: {
      serper: { apiKey: '' },
      serply: { apiKey: '' },
      searchapi: { apiKey: '' },
      tavily: { apiKey: '' },
      parallel: { apiKey: '' },
      exa: { apiKey: '' },
      firecrawl: { apiKey: '' },
      keenable: { apiKey: '' },
    },
  }
}

export function resolveAiSearchSettings(
  stored: Partial<AiSearchSettings> | undefined,
): AiSearchSettings {
  const defaults = defaultAiSearchSettings()
  if (!stored) return defaults
  const providers = { ...defaults.providers }
  for (const id of [
    'serper',
    'serply',
    'searchapi',
    'tavily',
    'parallel',
    'exa',
    'firecrawl',
    'keenable',
  ] as const) {
    const key = stored.providers?.[id]?.apiKey
    if (typeof key === 'string') providers[id] = { apiKey: key.trim() }
  }
  return { provider: stored.provider ?? defaults.provider, providers }
}

/**
 * Parallel and Keenable can run keylessly; other custom providers require a key or fall back
 * to RamonMods.
 */
export function activeSearchProvider(settings: Pick<AiSettings, 'search'>): AiSearchProviderId {
  const search = settings.search
  if (!search || search.provider === 'ramonmods') return 'ramonmods'
  if (!AI_SEARCH_PROVIDERS.some((m) => m.id === search.provider)) return 'ramonmods'
  if (search.provider === 'parallel' || search.provider === 'keenable') return search.provider
  // Trim-aware: a whitespace-only key from in-memory settings falls back
  // instead of sending `Bearer    ` to the search backend.
  return search.providers?.[search.provider]?.apiKey?.trim() ? search.provider : 'ramonmods'
}
