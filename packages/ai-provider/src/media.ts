import type {
  AiMediaProviderConfig,
  AiMediaProviderId,
  AiMediaProviderMeta,
  AiMediaSettings,
  AiSettings,
} from './types'

export const OPENAI_IMAGES_BASE_URL = 'https://api.openai.com/v1'
export const GEMINI_MEDIA_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta'
export const ARK_BASE_URL = 'https://ark.cn-beijing.volces.com/api/v3'
export const ZHIPU_BASE_URL = 'https://open.bigmodel.cn/api/paas/v4'
export const XAI_BASE_URL = 'https://api.x.ai/v1'
/** DashScope root: images ride /api/v1/services/aigc/..., understanding rides /compatible-mode/v1 */
export const DASHSCOPE_BASE_URL = 'https://dashscope.aliyuncs.com'
export const MINIMAX_BASE_URL = 'https://api.minimax.io/v1'
export const STABILITY_BASE_URL = 'https://api.stability.ai'
export const BFL_BASE_URL = 'https://api.bfl.ai'
/** DeepSeek's OpenAI-compatible root; the Vision + Files API live at https://api.deepseek.com */
export const DEEPSEEK_MEDIA_BASE_URL = 'https://api.deepseek.com/v1'

// Model ids verified against vendor docs 2026-09; keep chat-capable analysis
// models in step with the chat catalog in providers.ts.
export const AI_MEDIA_PROVIDERS: AiMediaProviderMeta[] = [
  {
    id: 'ramonmods',
    label: 'RamonMods',
    description: 'Geração de imagens, análise de mídia e busca com seu login da RamonMods',
    keyPlaceholder: 'Não é necessário - faça login na RamonMods',
    defaultBaseUrl: '',
    imageProtocol: 'openai-images',
    imageModels: [],
    defaultImageModel: '',
    analysisProtocol: 'openai-chat',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: true,
  },
  {
    id: 'openai',
    label: 'OpenAI',
    description: 'GPT Image para geração e edição; modelos de chat GPT para análise de imagens',
    keyPlaceholder: 'sk-...',
    defaultBaseUrl: OPENAI_IMAGES_BASE_URL,
    imageProtocol: 'openai-images',
    // GPT Image 2.5 (sunburst = quality, flare = fast) per the OpenAI models page
    // 2026-09-24, same token rates as GPT Image 2, both on generations and edits.
    // GPT-6 reads images over Chat Completions; the tool-call caveat that keeps
    // it off the chat provider does not apply to analysis.
    imageModels: [
      'gpt-image-2.5-sunburst',
      'gpt-image-2.5-flare',
      'gpt-image-2',
      'gpt-image-1.5',
      'gpt-image-1',
      'gpt-image-1-mini',
    ],
    defaultImageModel: 'gpt-image-2',
    analysisProtocol: 'openai-chat',
    analysisModels: [
      'gpt-6.1-terra',
      'gpt-6-sol',
      'gpt-6-luna',
      'gpt-5.6-sol',
      'gpt-5.6-terra',
      'gpt-5.6-luna',
      'gpt-5.5',
      'gpt-5.4-mini',
    ],
    defaultAnalysisModel: 'gpt-5.6-luna',
    videoAnalysis: false,
  },
  {
    id: 'gemini',
    label: 'Gemini',
    description:
      'Geração nativa de imagens pelo Gemini (família Nano Banana); o Gemini lê imagens, vídeos e áudio',
    keyPlaceholder: 'AIza...',
    defaultBaseUrl: GEMINI_MEDIA_BASE_URL,
    imageProtocol: 'gemini',
    // per the Google models page 2026-10-28: Nano Banana 2.1 leads the family
    // and Imagen 4 is shut down, so no Imagen id is listed
    imageModels: [
      'gemini-nano-banana-2.1',
      'gemini-3.1-flash-image',
      'gemini-3-pro-image',
      'gemini-3.1-flash-lite-image',
      'gemini-2.5-flash-image',
    ],
    defaultImageModel: 'gemini-nano-banana-2.1',
    analysisProtocol: 'gemini',
    analysisModels: [
      'gemini-4-argon',
      'gemini-3.8-flash',
      'gemini-3.7-flash',
      'gemini-3.1-pro-preview',
      'gemini-3.6-flash',
    ],
    defaultAnalysisModel: 'gemini-3.8-flash',
    videoAnalysis: true,
  },
  {
    id: 'doubao',
    label: 'Doubao (Volcengine Ark)',
    description: 'Geração e edição de imagens com Seedream; o Doubao Seed lê imagens e vídeos',
    keyPlaceholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
    defaultBaseUrl: ARK_BASE_URL,
    imageProtocol: 'openai-images',
    // per the Ark model list 2026-10-12: Seedream 5.0 split into a pro cut
    // (layer split, interactive editing, 14-language text) and a flash cut
    // (same feature set, cheaper and faster), so 5.0 is no longer one model.
    // The 4.5 and 4.0 IDs are both flagged for retirement there and come out.
    // 5-0-260128 stays — only its lite alias is on the way out.
    imageModels: [
      'doubao-seedream-5-0-pro-260628',
      'doubao-seedream-5-0-flash-260915',
      'doubao-seedream-5-0-260128',
    ],
    defaultImageModel: 'doubao-seedream-5-0-pro-260628',
    analysisProtocol: 'openai-chat',
    analysisModels: [
      'doubao-seed-2-1-pro-260915',
      'doubao-seed-2-1-pro-260628',
      'doubao-seed-2-1-turbo-260628',
    ],
    defaultAnalysisModel: 'doubao-seed-2-1-pro-260915',
    videoAnalysis: true,
  },
  {
    id: 'glm',
    label: 'GLM (Zhipu)',
    description: 'Geração de imagens com CogView; o GLM-V lê imagens e vídeos',
    keyPlaceholder: 'xxxxxxxx.xxxxxxxx',
    defaultBaseUrl: ZHIPU_BASE_URL,
    imageProtocol: 'openai-images',
    // per the Zhipu API reference 2026-11-02: the image flagship is GLM-Image
    // (CogView stays at 4 — there is no CogView 4.5 or 5)
    imageModels: ['glm-image', 'cogview-4-250304', 'cogview-3-flash'],
    defaultImageModel: 'glm-image',
    analysisProtocol: 'openai-chat',
    analysisModels: ['glm-4.6v', 'glm-4.6v-flash'],
    defaultAnalysisModel: 'glm-4.6v',
    videoAnalysis: true,
  },
  {
    id: 'xai',
    label: 'Grok (xAI)',
    description: 'Geração de imagens com Grok Imagine; o Grok lê imagens',
    keyPlaceholder: 'xai-...',
    defaultBaseUrl: XAI_BASE_URL,
    imageProtocol: 'openai-images',
    imageModels: ['grok-imagine-image-2.0', 'grok-2-image-1212'],
    defaultImageModel: 'grok-imagine-image-2.0',
    analysisProtocol: 'openai-chat',
    analysisModels: ['grok-4.6', 'grok-4.5'],
    defaultAnalysisModel: 'grok-4.6',
    videoAnalysis: false,
  },
  {
    id: 'qwen',
    label: 'Qwen (Alibaba Cloud)',
    description: 'Geração com Qwen-Image; o Qwen-VL lê imagens e vídeos (DashScope)',
    keyPlaceholder: 'sk-...',
    defaultBaseUrl: DASHSCOPE_BASE_URL,
    imageProtocol: 'dashscope',
    // per the Model Studio text-to-image docs 2026-11-02: qwen-image 3.0 and
    // 2.0 are the published families (3.0 pro flagship, 2.0 pro balanced);
    // 2.1 ids are not on the EN/ZH pages, so they are not listed. The older
    // qwen-image-plus/qwen-image ids stay listed, not seen retired
    imageModels: [
      'qwen-image-3.0-pro',
      'qwen-image-3.0',
      'qwen-image-2.0-pro',
      'qwen-image-2.0',
      'qwen-image-plus',
      'qwen-image',
    ],
    defaultImageModel: 'qwen-image-3.0-pro',
    analysisProtocol: 'openai-chat',
    analysisModels: ['qwen3-vl-plus', 'qwen3-vl-flash', 'qwen3.8-max', 'qwen3.7-plus'],
    defaultAnalysisModel: 'qwen3-vl-plus',
    videoAnalysis: true,
  },
  {
    id: 'minimax',
    label: 'MiniMax',
    description:
      'Geração de imagens; o MiniMax-M3 lê imagens e vídeos (use api.minimaxi.com/v1 para CN)',
    keyPlaceholder: 'eyJ...',
    defaultBaseUrl: MINIMAX_BASE_URL,
    imageProtocol: 'minimax',
    // per the platform docs 2026-11-14: MiniMax-H3-Image and its Highspeed
    // twin speak the OpenAI-Images-shaped /v1/images/* endpoints on both
    // platforms — confirmed on the international docs (api.minimax.io) and
    // on the CN platform (platform.minimaxi.com, api base api.minimax.cn);
    // image-01 keeps api.minimaxi.com as its CN host
    imageModels: ['image-01', 'MiniMax-H3-Image', 'MiniMax-H3-Image-Highspeed'],
    defaultImageModel: 'image-01',
    analysisProtocol: 'openai-chat',
    analysisModels: ['MiniMax-M3'],
    defaultAnalysisModel: 'MiniMax-M3',
    videoAnalysis: true,
  },
  {
    id: 'stability',
    label: 'Stability AI',
    description: 'Stable Image Ultra/Core e Stable Diffusion 3.5 pela API REST da Stability',
    keyPlaceholder: 'sk-...',
    defaultBaseUrl: STABILITY_BASE_URL,
    imageProtocol: 'stability',
    // per docs.stability.ai 2026-11-02: generation lives at
    // /v2beta/stable-image/generate/{ultra|core|sd3}; sd3.5 ids pick the model
    // inside the sd3 endpoint
    imageModels: ['stable-image-ultra', 'stable-image-core', 'sd3.5-large', 'sd3.5-medium'],
    defaultImageModel: 'stable-image-ultra',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: false,
  },
  {
    id: 'bfl',
    label: 'FLUX (Black Forest Labs)',
    description: 'Geração de imagens FLUX pela API da Black Forest Labs',
    keyPlaceholder: 'Chave de API',
    defaultBaseUrl: BFL_BASE_URL,
    imageProtocol: 'bfl',
    // per docs.bfl.ai 2026-11-02: the endpoint is /v1/<id>; FLUX.2 Pro is the
    // production-recommended default, FLUX 3 Image is the newest generation
    imageModels: ['flux-2-pro', 'flux-3-image', 'flux-pro-1.1'],
    defaultImageModel: 'flux-2-pro',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: false,
  },
  {
    id: 'microsoft',
    label: 'Microsoft (MAI-Image)',
    description:
      'Geração e edição de imagens MAI-Image pelo Microsoft Foundry; a URL base é o endpoint do seu recurso (https://<resource>.services.ai.azure.com)',
    keyPlaceholder: 'Chave de API',
    needsBaseUrl: true,
    defaultBaseUrl: '',
    imageProtocol: 'mai',
    // per learn.microsoft.com "deploy and use MAI image models" 2026-11-03: all
    // five are Preview Global Standard deployments; the model field carries the
    // deployment name, so a differently named deployment goes into the field
    imageModels: [
      'MAI-Image-2.6',
      'MAI-Image-2.6-Flash',
      'MAI-Image-2.5-Pro',
      'MAI-Image-2.5-Flash',
      'MAI-Image-2.5',
    ],
    defaultImageModel: 'MAI-Image-2.6',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: false,
  },
  {
    id: 'deepseek',
    label: 'DeepSeek',
    // native multimodal (V4.1 Flash = wire id deepseek-flash): reads JPEG/PNG/GIF/WebP,
    // no image generation and no video/audio, so it only appears for image analysis
    description: 'O DeepSeek V4.1 Flash lê imagens (multimodal nativo); sem geração de imagens',
    keyPlaceholder: 'sk-...',
    defaultBaseUrl: DEEPSEEK_MEDIA_BASE_URL,
    imageModels: [],
    defaultImageModel: '',
    analysisProtocol: 'openai-chat',
    analysisModels: ['deepseek-flash'],
    defaultAnalysisModel: 'deepseek-flash',
    videoAnalysis: false,
  },
  {
    id: 'custom',
    label: 'Personalizado',
    description: 'Qualquer endpoint compatível com OpenAI: /images/generations e /chat/completions',
    keyPlaceholder: 'Chave de API',
    needsBaseUrl: true,
    defaultBaseUrl: '',
    imageProtocol: 'openai-images',
    imageModels: [],
    defaultImageModel: '',
    analysisProtocol: 'openai-chat',
    analysisModels: [],
    defaultAnalysisModel: '',
    videoAnalysis: false,
  },
]

export type MediaCapability = 'image' | 'analysis' | 'video'

export function getMediaProviderMeta(id: AiMediaProviderId): AiMediaProviderMeta | undefined {
  return AI_MEDIA_PROVIDERS.find((m) => m.id === id)
}

export function providerHasCapability(
  meta: AiMediaProviderMeta,
  capability: MediaCapability,
): boolean {
  if (capability === 'image') return !!meta.imageProtocol
  if (capability === 'video') return !!meta.analysisProtocol && meta.videoAnalysis
  return !!meta.analysisProtocol
}

export function defaultAiMediaSettings(): AiMediaSettings {
  const providers = {} as AiMediaSettings['providers']
  for (const meta of AI_MEDIA_PROVIDERS) {
    providers[meta.id] = {
      apiKey: '',
      imageModel: meta.defaultImageModel,
      analysisModel: meta.defaultAnalysisModel,
      baseUrl: meta.needsBaseUrl ? '' : undefined,
    }
  }
  return {
    imageProvider: 'ramonmods',
    analysisProvider: 'ramonmods',
    videoAnalysisProvider: 'ramonmods',
    providers,
  }
}

/**
 * Merge a stored media block over defaults, trimming pasted whitespace. The
 * pre-catalog `provider` field (one choice for both tools) seeds both
 * per-capability choices. Unknown provider ids are kept and simply never activate.
 */
export function resolveAiMediaSettings(
  stored: Partial<AiMediaSettings> | undefined,
): AiMediaSettings {
  const defaults = defaultAiMediaSettings()
  if (!stored) return defaults
  const providers = { ...defaults.providers }
  for (const [id, config] of Object.entries(stored.providers ?? {})) {
    if (!config || typeof config !== 'object') continue
    // Hand-edited settings files can carry non-string values: trim only
    // strings (like the search-settings guard) instead of crashing.
    const str = (v: unknown, fallback: string): string =>
      typeof v === 'string' ? v.trim() : fallback
    const base = providers[id as AiMediaProviderId]
    providers[id as AiMediaProviderId] = {
      apiKey: str(config.apiKey, base?.apiKey ?? ''),
      imageModel: str(config.imageModel, base?.imageModel ?? ''),
      analysisModel: str(config.analysisModel, base?.analysisModel ?? ''),
      ...(config.baseUrl !== undefined
        ? { baseUrl: str(config.baseUrl, base?.baseUrl ?? '') }
        : base?.baseUrl !== undefined
          ? { baseUrl: base.baseUrl }
          : {}),
    }
  }
  const legacy = stored.provider
  const analysisProvider = stored.analysisProvider ?? legacy ?? defaults.analysisProvider
  return {
    imageProvider: stored.imageProvider ?? legacy ?? defaults.imageProvider,
    analysisProvider,
    // a pre-split file used one vendor for all media analysis
    videoAnalysisProvider: stored.videoAnalysisProvider ?? analysisProvider,
    providers,
  }
}

/** Key (or base URL for custom) present — the minimum for a BYOK media provider to be honored */
export function mediaConfigUsable(
  meta: AiMediaProviderMeta,
  config: AiMediaProviderConfig | undefined,
): boolean {
  if (!config) return false
  // Trim-aware like activeProvider: whitespace-only survivors of in-memory
  // settings are not usable configs.
  if (meta.needsBaseUrl) return !!config.baseUrl?.trim()
  return !!config.apiKey?.trim()
}

/**
 * The stored provider for one capability, honored only when it exists, has
 * that capability and is usable; anything else falls back to ramonmods so a
 * half-filled setup degrades to the signed-in default.
 */
export function activeMediaProvider(
  settings: Pick<AiSettings, 'media'>,
  capability: MediaCapability,
): AiMediaProviderId {
  const media = settings.media
  if (!media) return 'ramonmods'
  const id =
    capability === 'image'
      ? media.imageProvider
      : capability === 'video'
        ? media.videoAnalysisProvider
        : media.analysisProvider
  if (!id || id === 'ramonmods') return 'ramonmods'
  const meta = getMediaProviderMeta(id)
  if (!meta || !providerHasCapability(meta, capability)) return 'ramonmods'
  if (!mediaConfigUsable(meta, media.providers?.[id])) return 'ramonmods'
  return id
}

/** the active BYOK config for one capability, or null when it runs through RamonMods */
export function activeMediaConfig(
  settings: Pick<AiSettings, 'media'>,
  capability: MediaCapability,
): { provider: Exclude<AiMediaProviderId, 'ramonmods'>; config: AiMediaProviderConfig } | null {
  const provider = activeMediaProvider(settings, capability)
  if (provider === 'ramonmods') return null
  return { provider, config: settings.media!.providers[provider] }
}

function byokModel(
  settings: Pick<AiSettings, 'media'>,
  capability: MediaCapability,
): string | null {
  const active = activeMediaConfig(settings, capability)
  if (!active) return null
  const meta = getMediaProviderMeta(active.provider)!
  return capability === 'image'
    ? active.config.imageModel || meta.defaultImageModel
    : active.config.analysisModel || meta.defaultAnalysisModel
}

function capabilityAvailable(
  settings: Pick<AiSettings, 'media' | 'ramonmodsToolsEnabled'> | null | undefined,
  ramonmodsLoggedIn: boolean,
  capability: MediaCapability,
): boolean {
  if (!settings) return ramonmodsLoggedIn
  const model = byokModel(settings, capability)
  if (model !== null) return model !== ''
  return ramonmodsLoggedIn && settings.ramonmodsToolsEnabled === true
}

/** live predicate for the generate_image tool: BYOK image model configured, or ramonmods login + cloud tools on */
export function imageGenerationAvailable(
  settings: Pick<AiSettings, 'media' | 'ramonmodsToolsEnabled'> | null | undefined,
  ramonmodsLoggedIn: boolean,
): boolean {
  return capabilityAvailable(settings, ramonmodsLoggedIn, 'image')
}

/** live predicate for the analyze_media tool: image analysis or video analysis reachable */
export function mediaAnalysisAvailable(
  settings: Pick<AiSettings, 'media' | 'ramonmodsToolsEnabled'> | null | undefined,
  ramonmodsLoggedIn: boolean,
): boolean {
  return (
    capabilityAvailable(settings, ramonmodsLoggedIn, 'analysis') ||
    capabilityAvailable(settings, ramonmodsLoggedIn, 'video')
  )
}

export function videoAnalysisAvailable(
  settings: Pick<AiSettings, 'media' | 'ramonmodsToolsEnabled'> | null | undefined,
  ramonmodsLoggedIn: boolean,
): boolean {
  return capabilityAvailable(settings, ramonmodsLoggedIn, 'video')
}
