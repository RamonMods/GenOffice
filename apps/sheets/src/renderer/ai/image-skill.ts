import type { AgentSkill } from '@genoffice/agent-core'
import { t } from '../i18n/locale'

/**
 * Image acquisition AgentSkill: image_search (shared main-process channel, same
 * source as docs/slides) and generate_image (sheets-owned channel).
 * Both return a URL; placement happens through the normal propose_operations
 * add_image path, which downloads the URL in the main process on apply.
 */

const PLACEMENT_PROMPT = `- Para colocar uma imagem em uma planilha, passe a URL para propose_operations {op:"add_image", sheetId, path:"<https url>", anchorCell} — detalhes dos campos no guia de gráficos. A imagem é ancorada nessa célula e gravada no arquivo ao salvar (apenas xlsx importado).
- Insira apenas as imagens que o usuário pediu; a correção dos dados sempre prevalece sobre a decoração.`

const IMAGES_SYSTEM_PROMPT = `## Imagens
- image_search encontra imagens reais na web (retorna entradas diretas de imageUrl); generate_image cria uma ilustração com IA quando não existe uma imagem real adequada ou o usuário pede expressamente arte gerada.
${PLACEMENT_PROMPT}`

const IMAGES_SYSTEM_PROMPT_NO_GEN = `## Imagens
- image_search encontra imagens reais na web (retorna entradas diretas de imageUrl).
${PLACEMENT_PROMPT}`

/**
 * `imageGen` is a live predicate (RamonMods login + cloud-tools toggle, or a
 * BYOK media key); the loop re-reads tools and systemPrompt before every
 * request, so generate_image appears and disappears without rebuilding the loop.
 */
export function createImageSkill(imageGen: () => boolean = () => true): AgentSkill {
  const allTools = [
    {
      name: 'image_search',
      description:
        'Pesquisa imagens na web. Retorna uma lista numerada de entradas diretas de imageUrl com tamanhos em pixels; ' +
        'escolha uma e insira com propose_operations add_image (path = a URL).',
      inputSchema: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Palavras-chave da pesquisa de imagens (em inglês funciona melhor)' },
          maxResults: { type: 'integer', description: 'Número máximo de resultados, padrão 8' },
        },
        required: ['query'],
      },
    },
    {
      name: 'generate_image',
      description:
        'Gera uma imagem com IA a partir de um prompt de texto. Retorna uma URL para inserir ' +
        'com propose_operations add_image. Use para ilustrações/arte decorativa; prefira image_search para assuntos do mundo real.',
      inputSchema: {
        type: 'object',
        properties: {
          prompt: {
            type: 'string',
            description: 'O que desenhar — assunto, estilo, composição (em inglês funciona melhor)',
          },
          aspectRatio: {
            type: 'string',
            description: 'Proporção como "1:1", "16:9", "4:3"; padrão 1:1',
          },
        },
        required: ['prompt'],
      },
    },
  ]
  return {
    id: 'images',
    get systemPrompt() {
      return imageGen() ? IMAGES_SYSTEM_PROMPT : IMAGES_SYSTEM_PROMPT_NO_GEN
    },
    get tools() {
      return imageGen() ? allTools : allTools.filter((t) => t.name !== 'generate_image')
    },
    executeTool: async (call) => {
      if (call.name === 'image_search') {
        const query = String(call.input.query ?? '').trim()
        if (!query) {
          return {
            output: 'query não pode estar vazia',
            isError: true,
            summary: t('aiToolImageSearch'),
          }
        }
        const result = await window.desktopApi.imageSearch(
          query,
          Number(call.input.maxResults) || 8,
        )
        // A backend failure must not read as an empty gallery — the model
        // would fabricate image choices
        if (result.method === 'error') {
          return {
            output: `falha na pesquisa de imagens (erro do serviço, não um resultado vazio — você pode tentar novamente): ${result.error ?? 'erro desconhecido'}`,
            isError: true,
            summary: t('aiToolImageSearch'),
          }
        }
        const lines = result.images.map(
          (image, index) =>
            `${index + 1}. ${image.title || '(sem título)'} [${image.width ?? '?'}x${image.height ?? '?'}]\n   ${image.imageUrl}`,
        )
        return {
          output: lines.join('\n') || '(sem imagens)',
          mutated: false,
          summary: t('aiToolImageSearchDone', { query, count: result.images.length }),
        }
      }
      if (call.name === 'generate_image') {
        const prompt = String(call.input.prompt ?? '').trim()
        if (!prompt) {
          return { output: 'prompt não pode estar vazio', isError: true, summary: t('aiToolGenImage') }
        }
        const aspectRatio = String(call.input.aspectRatio ?? '').trim()
        const result = await window.desktopApi.generateImage({
          prompt,
          ...(aspectRatio ? { aspectRatio } : {}),
        })
        if (!result.url) {
          return {
            output: `falha na geração da imagem: ${result.error ?? 'erro desconhecido'}`,
            isError: true,
            summary: t('aiToolGenImage'),
          }
        }
        return {
          output: `Imagem gerada: ${result.url}\nInsira com propose_operations {op:"add_image", path:"${result.url}", ...}.`,
          mutated: false,
          summary: t('aiToolGenImageDone'),
        }
      }
      return { output: `Ferramenta desconhecida: ${call.name}`, isError: true, summary: call.name }
    },
  }
}
