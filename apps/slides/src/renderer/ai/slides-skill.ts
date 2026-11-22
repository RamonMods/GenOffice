import type { AgentSkill, ToolDisplay } from '@genoffice/agent-core'
import type {
  GroupRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderSlide,
  ShapeRenderNode,
} from '@genoffice/pptx-render'
import type { AgentToolCall, AgentToolDef } from '../../shared/ipc'
import { OP_GROUPS, opGuide, opGuideCatalog, opSignatureIndex } from '@genoffice/pptx-ops/op-docs'
import { auditSlideLayout, formatAudit } from '@genoffice/pipelines/slides/layout-audit'
import { runLayoutScript, type LayoutScriptElement } from './layout-script'
import {
  extractLayoutSkeleton,
  formatSkeletonForPrompt,
  skeletonRole,
  type LayoutSkeleton,
} from './layout-skeleton'
import { t } from '../i18n/locale'
import systemPrompt from './prompts/system.md?raw'
import { placeholderInstruction } from './redact'
import { textForModel, redactLabelsOf } from './redact-view'
import { redactGuardFor } from './redact-guard'

/**
 * Slides capability as an AgentSkill: deck outline context + three tools (read structure /
 * read one slide / edit element text). Changes go through the existing slides:edit-text IPC;
 * the main process applies them and returns the new RenderSlide, which applySlide writes
 * back into React state — the same pipeline as manual editing.
 */

// ── Generation progress events (for the onProgress callback; renderer memory only, never persisted or journaled) ──

/** Upper bound on generate_deck approx_pages; bounds planner round trips and the progress note. */
export const MAX_APPROX_PAGES = 200
const MAX_PROGRESS_NOTE_PAGES = 40

/** Per-page progress status */
export type PageProgressStatus = 'pending' | 'running' | 'done' | 'error'

/** Per-page progress entry */
export interface PageProgressItem {
  title: string
  status: PageProgressStatus
  /** Failure reason (when status='error'); cleared after a successful retry */
  error?: string
}

/** Progress event union type (all stages share one callback; the UI dispatches by stage) */
export type DeckProgressEvent =
  | { stage: 'style'; label: string; status: 'running' | 'done' | 'error'; summary: string }
  | {
      stage: 'plan'
      label: string
      done: number
      total: number
      status: 'running' | 'done' | 'error'
      summary: string
    }
  | {
      stage: 'images'
      label: string
      done: number
      total: number
      status: 'running' | 'done' | 'error'
      summary: string
    }
  | {
      stage: 'pages'
      label: string
      done: number
      total: number
      status: 'running' | 'done' | 'error'
      summary: string
      pages: PageProgressItem[]
    }
  | {
      stage: 'done'
      total: number
      summary: string
      /** absent on success; the card must not read a failed or stopped run as "done" */
      outcome?: 'failed' | 'cancelled'
    }

/** Panel/skill access point to the currently open deck (refs provided by App, stay fresh across renders). */
export interface DeckAccess {
  getSlides(): RenderSlide[]
  getCurrent(): number
  getSelectedIds(): string[]
  applySlide(slideIndex: number, updated: RenderSlide): void
  /** Replace the whole deck (after adding/removing slides) and jump to the goTo slide */
  applyDeck(slides: RenderSlide[], goTo?: number): void
  /**
   * Generation progress callback (optional): called by generate_deck stages; the UI updates
   * the progress card and top progress bar in real time. Passed only through renderer
   * memory, never persisted or journaled.
   */
  onProgress?(event: DeckProgressEvent): void
  /** Land generated pages: each pageMarkers entry redeems a one-slide pptx, merged into / replacing the current deck. Returns total page count or an error.
   *  mode="insert_at" inserts a single page at position insertAt (later pages shift) — used to re-insert failed pages at their original position.
   *  On pipeline failure it automatically falls back to element-level mode; fallbackReason explains why (ok is still true).
   *  deckName = presentation name derived from user input, used as the file name when the new draft is saved (instead of "Untitled-timestamp"). */
  landGeneratedPages?(
    pageMarkers: string[],
    mode?: 'replace' | 'append' | 'insert_at',
    deckName?: string,
    insertAt?: number,
  ): Promise<{
    ok: boolean
    pages?: number
    appendedFrom?: number
    insertedIndex?: number
    error?: string
    fallbackReason?: string
    imageFailures?: { page: number; url: string }[]
  }>
  /** Redo one slide in place: land the marker's page as a replacement for slide slideIndex (other slides untouched; undoable with ⌘Z). */
  regenerateSlide?(
    slideIndex: number,
    marker: string,
  ): Promise<{ ok: boolean; error?: string; imageFailures?: { page: number; url: string }[] }>
  /** Survey: shows a card with options and waits for the user's choices, returning an answer summary. */
  askClarification?(questions: ClarifyQuestion[]): Promise<{ answers: string; cancelled?: boolean }>
  /**
   * In-tool image search (embedded in the tool):
   * given English keywords, returns an array of real image URLs (at most N).
   * On search failure returns an empty array (fail-open; doesn't block the main generation path).
   */
  searchImages?(query: string, maxResults: number): Promise<string[]>
  /** live predicate (ramonmods login && cloud-tools toggle, or a BYOK media key); false hides generate_image */
  imageGenAvailable?(): boolean
  /** same for analyze_media */
  mediaAnalysisAvailable?(): boolean
  /**
   * Single-page generation used by generate_deck / regenerate_slide: one LLM request writes a
   * structured slide spec and the main process builds it directly into a one-slide pptx.
   * Returns a marker string that goes into a landGeneratedPages pageMarkers slot.
   */
  generatePageLocal?(args: {
    pageIndex: number
    totalPages: number
    coreHook: string
    style: string
    title: string
    brief: string
    layout: string
    images: string[]
    context?: string
    topic?: string
    canvasW: number
    canvasH: number
    /** Formatted template-chrome block (layout-skeleton.ts) for this page's role; absent without a template */
    skeleton?: string
    signal?: AbortSignal
  }): Promise<{ ok: boolean; marker?: string; error?: string; imageFailures?: string[] }>
  /**
   * In-tool Style Skill generation:
   * a dedicated LLM call focused on producing a complete structured visual style guide
   * (color rules/fonts/layout variants per page type/overall style). Promotes style from an
   * "outline side-product" to a "dedicated deliverable" — less AI-looking, consistent across pages.
   */
  generateStyleSkill?(args: {
    topic: string
    questionnaire?: string
    styleHint?: string
    signal?: AbortSignal
  }): Promise<{ ok: boolean; styleSkill?: string; error?: string }>
  /**
   * In-tool planning: given topic + page count, the LLM produces a structured outline.
   * Fixes "missing pages at the input side" at the root — the main agent doesn't hand-write dozens of pages of pages JSON (avoids the argument being truncated by max_tokens).
   * style is already produced by generateStyleSkill; this function only outputs core_hook + per-page outlines (styleSkill serves as a reference for consistency).
   * Batched recursion is scheduled by the skill (continueFrom keeps the narrative coherent across batches).
   */
  planDeckOutline?(args: {
    topic: string
    count: number
    startPage: number
    context?: string
    styleSkill?: string
    continueFrom?: { coreHook: string }
    signal?: AbortSignal
  }): Promise<{
    ok: boolean
    // Same loose shape as OutlineJson (outline-json.ts): the LLM output is
    // only validated field-by-field at the point of use.
    outline?: { core_hook?: unknown; pages?: unknown }
    error?: string
  }>
  /**
   * Persist the current draft's Style Skill as a sidecar file (same directory and name as the draft, .styleskill.json).
   * fail-open: failure doesn't block the main path.
   */
  saveSidecar?(data: { topic: string; styleSkill: string; createdAt: string }): Promise<void>
  /**
   * Save styleSkill into userData/style-templates/<name>.json for later reuse.
   * `layout` is the deck's extracted chrome skeleton (title/brand-image slots,
   * accents, backgrounds) — generation uses it to keep those elements consistent
   * across pages (see layout-skeleton.ts).
   */
  saveStyleTemplate?(
    name: string,
    data: { topic: string; styleSkill: string; createdAt: string; layout?: LayoutSkeleton },
  ): Promise<{ ok: boolean; error?: string }>
  /**
   * List saved Style templates (name + topic + createdAt).
   */
  listStyleTemplates?(): Promise<Array<{ name: string; topic: string; createdAt: string }>>
  /**
   * Load the content of a given Style template (styleSkill text plus, for
   * templates saved by a build with skeleton support, the layout skeleton).
   */
  loadStyleTemplate?(name: string): Promise<{
    ok: boolean
    styleSkill?: string
    topic?: string
    layout?: LayoutSkeleton
    error?: string
  }>
  fitWidthPx: number
  /** Base retry backoff in ms for single-page generation failures (default 2000; tests pass 0 to disable backoff) */
  retryBackoffMs?: number
  /**
   * Names of text attachments in the current conversation that were never read with
   * read_attachment. When non-empty, generate_deck refuses to run until they are read
   * (decks must be built from attachment content, not generic filler).
   */
  unreadTextAttachments?(): string[]
  /**
   * Resolve a user image attachment by file name (an `attachment://` reference in
   * insert_web_image / replace_image) to its raw bytes, so the original file is
   * embedded as-is — the model must never recreate an attached image (r182 family).
   */
  resolveAttachmentImage?(
    name: string,
  ): Promise<{ ok: true; base64: string; ext: string } | { ok: false; error: string }>
}

/** `attachment://<file name>` → decoded file name, or null when not an attachment reference. */
export function attachmentRefName(url: string): string | null {
  if (!url.toLowerCase().startsWith('attachment://')) return null
  const raw = url.slice('attachment://'.length).trim()
  try {
    return decodeURIComponent(raw)
  } catch {
    return raw
  }
}

/** Single survey question structure (with options). */
export interface ClarifyQuestion {
  id: string
  label: string
  description?: string
  /** Option text array (≤5 per question); the frontend automatically appends "Other (fill in)" */
  options: string[]
  /** Multi-select (single-select by default) */
  multi?: boolean
}

const TOOLS: AgentToolDef[] = [
  {
    name: 'read_slide',
    description:
      'Lê todos os elementos de uma página com o texto completo (sem truncamento) e as cores atuais (preenchimento/texto/contorno, em hexadecimal). Chame antes de reescrever uma página.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Número da página (baseado em 0)' },
      },
      required: ['slideIndex'],
    },
  },
  {
    name: 'execute_slide_script',
    description:
      "[Ferramenta preferencial para editar os elementos existentes de um slide] Executa seu script de edição em JavaScript em uma página; um único script cobre: posição/tamanho/alinhamento/distribuição/ajustes relativos/texto/estilo/preenchimento/contorno." +
      ' Em tempo de execução, o script recebe automaticamente a geometria e o texto reais de todos os elementos da página (els) — **não é preciso chamar read_slide antes**; leitura e escrita combinadas, calcule a partir de els dentro do script.' +
      ' O script inteiro é uma transação atômica: geometria em um lote, o restante na ordem do script, uma única etapa de desfazer; se qualquer operação falhar, tudo é revertido e a página permanece inalterada — corrija o script e reenvie-o inteiro. Uma auditoria de layout (sobreposição/fora dos limites/estouro de texto) é retornada ao final.' +
      ' Muito mais confiável do que chamadas individuais set_element_* — o cálculo de coordenadas acontece no local da execução, não de memória. Se a auditoria apontar problemas, chame esta ferramenta novamente logo para corrigir.\n' +
      'Ambiente de script (DSL síncrono limitado, semelhante a JavaScript; sem APIs externas nem globais do ambiente):\n' +
      '- els: array, cada item {id,type,text,x,y,w,h,rotation,fontSizePt?,fill?,textColor?,strokeColor?,inGroup?,groupId?,locked?} (pixels, origem no canto superior esquerdo; fill/textColor/strokeColor são as cores atuais em #RRGGBB, somente leitura — escreva via setFill/setStyle/setStroke; inGroup+groupId=filho de grupo editável diretamente (todas as primitivas funcionam, coordenadas absolutas como mostrado); inGroup sem groupId=aninhado em um subgrupo, somente leitura — aplique apply_ops ungroupElement no grupo externo primeiro; locked=decoração de layout, somente leitura)\n' +
      '- canvas: {w,h} tamanho da tela (px)\n' +
      "- setBox(id, {x?,y?,w?,h?,rotation?}): define a caixa-alvo de um elemento, informe apenas os campos a alterar\n" +
      '- moveBy(id, dx, dy): movimento relativo (esquerda = dx negativo, cima = dy negativo)\n' +
      '- resizeBy(id, dw, dh): redimensionamento relativo\n' +
      '- setText(id, textOrParagraphs): substitui o texto por completo; passe uma string (dividida em parágrafos por \\n) ou um array de parágrafos (mesmo formato dos parágrafos de apply_ops setText). Trechos/strings simples herdam o tamanho, a cor e a fonte originais — não adicione fontSize/color/fontFamily a menos que o usuário tenha pedido para reestilizar\n' +
      '- setStyle(id, {fontSize?,color?,bold?,italic?,underline?,align?,fontFamily?}): altera o estilo sem mudar o texto, informe apenas os campos a alterar\n' +
      "- setFill(id, colorOrNone): preenchimento sólido '#RRGGBB' ou 'none'\n" +
      '- setStroke(id, {color?,widthPt?} | null): contorno; passe null para remover\n' +
      '- log(...): saída de depuração (devolvida a você); o valor retornado também é devolvido a você (coloque um resumo nele)\n' +
      '- Cálculos suportados: const/let, aritmética, if/for/for...of/while, funções/arrows, literais de objeto/array JSON, Math, regex.test e métodos seguros de array/string. Sem classes, async, módulos, construtores, protótipos ou código dinâmico.\n' +
      'Exemplo 1 — três cartões com a mesma largura e o mesmo espaçamento:\n' +
      'const cards = els.filter(e => /card/.test(e.id));\n' +
      'const gap = 32, w = (canvas.w - 2*80 - (cards.length-1)*gap) / cards.length;\n' +
      'cards.forEach((c, i) => setBox(c.id, { x: 80 + i*(w+gap), y: 200, w, h: 320 }));\n' +
      "Exemplo 2 — mover o título um pouco para a esquerda: moveBy('title', -30, 0);\n" +
      "Exemplo 3 — deixar o título azul e negrito: setStyle('t1', { color: '#1a73e8', bold: true });",
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Número da página (baseado em 0)' },
        code: {
          type: 'string',
          description:
            'Corpo do script JS (código síncrono; pode usar els/canvas além de setBox/moveBy/resizeBy/setText/setStyle/setFill/setStroke/log; pode retornar um resumo)',
        },
        explanation: {
          type: 'string',
          description:
            'Uma frase descrevendo o que este script faz (≤60 caracteres, exibida ao usuário)',
        },
      },
      required: ['slideIndex', 'code'],
    },
  },
  {
    name: 'web_search',
    description:
      'Pesquisa na web por informações textuais (material/dados/fatos). Use quando precisar de informações atualizadas ou tiver dúvidas sobre fatos. Retorna título/link/trecho.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Palavras-chave de busca' },
        maxResults: { type: 'integer', description: 'Número máximo de resultados, padrão 6' },
      },
      required: ['query'],
    },
  },
  {
    name: 'image_search',
    description:
      'Pesquisa imagens (para as imagens dos slides). Retorna uma lista de imageUrl; após escolher, insira com insert_web_image.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Palavras-chave de busca de imagens (em inglês funciona melhor)' },
        maxResults: { type: 'integer', description: 'Número máximo de resultados, padrão 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'generate_image',
    description:
      'Geração/edição de imagens por IA. De texto para imagem, ou passe referenceImageUrls para editar uma imagem; retorna a URL de uma imagem. Imagens NOVAS: insira com insert_web_image. Edição de uma imagem JÁ EXISTENTE no slide (remoção de fundo/aumento de resolução/etc.): substitua-a no lugar com replace_image — não insira uma duplicata. Use para ilustrações/ícones/fundos personalizados e imagens consistentes com o estilo; para fotos/capturas de tela reais, continue usando image_search. NUNCA a use para recriar uma imagem anexada pelo usuário (logo, foto) — incorpore o arquivo original com insert_web_image / replace_image e url=attachment://<file name>. Ícones/logos/recortes que precisam ficar sobre o conteúdo do slide exigem transparentBackground:true — pedir um fundo transparente no prompt NÃO funciona (os modelos pintam um tabuleiro cinza falso nos pixels).',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: {
          type: 'string',
          description:
            'Descrição da imagem; em inglês funciona melhor (mantenha literalmente qualquer texto a ser renderizado na imagem)',
        },
        model: {
          type: 'string',
          description:
            "Opcional; usa o modelo configurado por padrão. Apenas RamonMods — especifique para fins especiais: fal-bria-rmbg=remoção de fundo, fal-ai/recraft-clarity-upscale=aumento de resolução, flux-pro/outpaint=outpaint, fal-ai/image-editing/text-removal=remover marca d'água de texto",
        },
        referenceImageUrls: {
          type: 'array',
          items: { type: 'string' },
          description: 'URLs de imagens de referência / imagens a editar (obrigatório para tarefas de edição)',
        },
        aspectRatio: {
          type: 'string',
          description: 'Proporção: 1:1|4:3|16:9|9:16|3:4|2:3|3:2|auto',
        },
        transparentBackground: {
          type: 'boolean',
          description:
            'Defina como true quando o resultado precisar ter um fundo realmente transparente (ícones, logos, recortes colocados sobre o conteúdo do slide). O aplicativo remove o fundo automaticamente após a geração; nunca dependa do prompt para a transparência.',
        },
      },
      required: ['prompt'],
    },
  },
  {
    name: 'analyze_media',
    description:
      'Analisa conteúdo de mídia: entende imagens/áudio/vídeo (vídeo e áudio exigem RamonMods ou Gemini como provedor de mídia). Passe URLs de mídia (ou caminhos de arquivos locais) e os requisitos de análise; retorna o texto da análise. Vídeo permite extrair pontos-chave, estrutura e intervalos de tempo — ideal para transformar o material do usuário em conteúdo útil para o deck.',
    inputSchema: {
      type: 'object',
      properties: {
        mediaUrls: {
          type: 'array',
          items: { type: 'string' },
          description: 'Lista de URLs de mídia ou caminhos de arquivos locais',
        },
        requirements: {
          type: 'string',
          description:
            'Requisitos de análise (em inglês): o que extrair e como o resultado será usado (ex.: extrair pontos-chave para os slides)',
        },
      },
      required: ['mediaUrls', 'requirements'],
    },
  },
  {
    name: 'insert_web_image',
    description:
      'Baixa uma URL de imagem obtida de image_search ou generate_image e a insere em uma página (coordenadas em pixels). w×h é um quadro de layout, não um alvo de esticamento: a imagem mantém sua proporção, preenche o quadro e o excedente é recortado no centro (object-fit: cover) — escolha o quadro livremente para o layout. ' +
      'Para colocar uma imagem ANEXADA PELO USUÁRIO (logo, foto, captura de tela), passe url=attachment://<file name> (o nome exato da lista de anexos) — o aplicativo incorpora os bytes do arquivo original como estão. Nunca recrie uma imagem anexada com generate_image e nunca peça base64.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        url: {
          type: 'string',
          description:
            'Link direto da imagem (imageUrl de image_search), ou attachment://<file name> para incorporar uma imagem anexada pelo usuário como está',
        },
        x: { type: 'number' },
        y: { type: 'number' },
        w: { type: 'number' },
        h: { type: 'number' },
      },
      required: ['slideIndex', 'url', 'x', 'y', 'w', 'h'],
    },
  },
  {
    name: 'replace_image',
    description:
      'Substitui a imagem de origem de uma figura por uma URL (de image_search ou generate_image) no lugar — posição, tamanho, ordem-z, borda e efeitos são preservados. Esta é a ferramenta para fluxos de "alterar/editar esta imagem com IA": por exemplo, execute generate_image com referenceImageUrls para remoção de fundo/aumento de resolução/edição e depois replace_image com a URL retornada. Uma imagem nova com proporção diferente nunca é esticada: ela preenche o quadro e é recortada no centro (object-fit: cover). keepCrop mantém a janela de recorte existente e só é correto quando a imagem nova tem a mesma geometria em pixels da antiga (ex.: saída de remoção de fundo).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'ID do elemento de imagem' },
        url: { type: 'string', description: 'Link direto da imagem' },
        keepCrop: { type: 'boolean', description: 'Mantém a janela de recorte existente (padrão false)' },
      },
      required: ['slideIndex', 'sourceId', 'url'],
    },
  },
  {
    name: 'ask_clarification',
    description:
      "[Chame antes de criar um deck inteiramente novo] Mostra um cartão de questionário com opções, permitindo que o usuário faça as escolhas principais deste deck (público/cenário/tom/foco etc.); as escolhas do usuário determinam diretamente o Core Hook e o estilo do deck. As perguntas devem ser específicas para o tema, cada uma representando um verdadeiro trade-off (as opções representam direções diferentes). Faça de 2 a 4 perguntas, com ≤5 opções cada. Após chamar, aguarde o usuário terminar de escolher no cartão e gere o deck quando tiver as respostas. Não repita as perguntas no texto da sua resposta.",
    inputSchema: {
      type: 'object',
      properties: {
        questions: {
          type: 'array',
          description: 'Lista de perguntas (de 2 a 4 perguntas)',
          items: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'ID único da pergunta (curto, em inglês/pinyin)' },
              label: { type: 'string', description: 'Texto da pergunta' },
              description: { type: 'string', description: 'Observação opcional de uma linha' },
              options: {
                type: 'array',
                items: { type: 'string' },
                description:
                  'Opções (≤5); o frontend acrescenta automaticamente "Decida por mim" e "Outro"',
              },
              multi: { type: 'boolean', description: 'Multisseleção (seleção única por padrão)' },
            },
            required: ['id', 'label', 'options'],
          },
        },
      },
      required: ['questions'],
    },
  },
  {
    name: 'plan_deck',
    description:
      "[Ao criar um deck inteiramente novo, chame depois de pesquisar o material/as imagens e antes de generate_deck] Gera um plano estruturado: o Core Hook + o esquema de estilo unificado + título/brief de conteúdo/layout/palavras-chave de imagem de cada página. Pense no deck inteiro primeiro, para evitar começar bem e terminar fraco. O plano é exibido ao usuário.",
    inputSchema: {
      type: 'object',
      properties: {
        core_hook: {
          type: 'string',
          description:
            "A âncora narrativa do deck (uma frase, com tensão, de preferência contendo um número ou um contraste contraintuitivo)",
        },
        style: {
          type: 'string',
          description:
            'Sistema de design unificado: cores primárias/secundárias, tom da fonte, margens de conteúdo, estilo de cartões/cantos (ex.: "azul escuro como primária + detalhes dourados, aparência de painel de dados"); todas as páginas o seguem',
        },
        pages: {
          type: 'array',
          description: 'Plano por página',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', description: 'Título da página' },
              type: { type: 'string', description: 'cover|content|data|closing' },
              brief: {
                type: 'string',
                description:
                  'Descrição do conteúdo da página (use dados/fatos reais; diga o que vai em cada região)',
              },
              layout: {
                type: 'string',
                description:
                  'Layout (ex.: three_column_cards/hero_big_number/two_column/timeline/left_text_right_image); páginas de conteúdo não podem repetir',
              },
              image_queries: {
                type: 'array',
                items: { type: 'string' },
                description:
                  "Palavras-chave de busca de imagens em inglês para os espaços de imagem desta página (uma por espaço; [] para nenhuma imagem)",
              },
            },
            required: ['title', 'brief', 'layout'],
          },
        },
      },
      required: ['core_hook', 'style', 'pages'],
    },
  },
  {
    name: 'regenerate_slide',
    description:
      '[Refazer/redesenhar uma página existente] Regenera a página a partir do seu brief e a substitui no lugar (outras páginas intactas, com opção de desfazer).' +
      ' Use quando o usuário disser "refaça esta página / redesenhe-a / tente outro layout / deixe-a mais bonita"; não desmonte a página elemento por elemento com as ferramentas nativas.' +
      " Fluxo: primeiro read_slide para obter o conteúdo atual da página, depois verifique as páginas vizinhas no esquema do deck para compreender o estilo do deck;" +
      ' escreva um brief detalhado — o que manter (copie o texto/dados reais para o brief literalmente), o que alterar e o layout desejado; o estilo do deck é aplicado automaticamente.' +
      ' Se a página precisar de imagens, use image_search primeiro e passe URLs reais em image_urls.' +
      ' Se a geração falhar, geralmente é um erro temporário: NÃO fique repetindo a tentativa — faça as alterações concretas no lugar com execute_slide_script (ou diga ao usuário para tentar novamente em alguns minutos).',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer', description: 'Página a refazer (baseada em 0)' },
        brief: {
          type: 'string',
          description:
            'Brief de conteúdo e layout da nova página: o que vai em cada região (copie para o brief o texto/dados reais a manter) e o layout a usar (ex.: three_column_cards/hero_big_number/two_column/timeline).',
        },
        title: { type: 'string', description: 'Título da página' },
        layout: { type: 'string', description: 'Nome da intenção de layout (opcional)' },
        image_urls: {
          type: 'array',
          items: { type: 'string' },
          description: 'URLs reais de imagens http(s) para esta página (use image_search primeiro; [] para nenhuma)',
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Obrigatório quando o brief contém números específicos (%, dinheiro, magnitudes): de onde vieram — 'user'/'document'/'search' (execute web_search primeiro)/'sample' (divulgue ao usuário)",
        },
      },
      required: ['slideIndex', 'brief'],
    },
  },
  {
    name: 'generate_deck',
    description:
      '[Primeira escolha para criar um deck inteiramente novo — pipeline autoguiada: busca automática de imagens, geração página a página com exibição em tempo real, sem páginas faltando]' +
      ' Uso recomendado (especialmente com muitas páginas): passe apenas topic + approx_pages (+ opcionalmente style/context); o sistema planeja o esquema internamente (em lotes automáticos acima de 12 páginas), **busca imagens automaticamente** (sem image_search prévio — o sistema busca internamente a partir das palavras-chave image_queries planejadas e preenche as URLs reais antes de desenhar as páginas), desenha as páginas uma a uma e as coloca na tela conforme terminam.' +
      ' Você não precisa escrever dezenas de páginas manualmente, e nem "somente a página 1 foi gerada" nem "os argumentos foram truncados" podem acontecer — a quantidade de páginas é garantida pelo loop do sistema.' +
      ' (Se você já conhece cada página, pode passar core_hook+style+pages diretamente; pages[].image_queries aceita palavras-chave de busca de imagens em inglês, pesquisadas internamente; se você já conhece URLs http(s) reais, passe-as diretamente — o sistema respeita as URLs existentes e não faz nova busca.)' +
      ' Para adicionar algumas páginas a um deck existente, passe pages (briefs apenas das páginas novas) + insert_mode:"append".',
    inputSchema: {
      type: 'object',
      properties: {
        topic: {
          type: 'string',
          description:
            "[Recomendado] A descrição do tema/dos requisitos do deck (com topic você não escreve as páginas manualmente; o sistema planeja internamente)",
        },
        approx_pages: {
          type: 'integer',
          minimum: 1,
          maximum: MAX_APPROX_PAGES,
          description: `Número esperado de páginas (usado junto com topic; no máximo ${MAX_APPROX_PAGES})`,
        },
        context: {
          type: 'string',
          description:
            'Opcional: material/dados/respostas de questionário reais obtidos com web_search, para que o planejamento interno use conteúdo real',
        },
        core_hook: {
          type: 'string',
          description:
            'Opcional: uma âncora narrativa que você já decidiu (recomendado junto com pages)',
        },
        style: {
          type: 'string',
          description:
            'Sistema de design unificado: cores primárias/secundárias, fontes, margens de conteúdo, cantos dos cartões (obrigatório com pages; opcional como dica de estilo com topic)',
        },
        pages: {
          type: 'array',
          description:
            'Opcional: passe diretamente quando você já conhece cada página (serão geradas tantas páginas quantas planejadas). Com muitas páginas, prefira topic e o planejamento interno, para evitar argumentos longos demais e truncados',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string', description: 'Título da página' },
              type: { type: 'string', description: 'cover|content|data|closing' },
              brief: {
                type: 'string',
                description: 'Descrição do conteúdo da página (use dados/fatos reais)',
              },
              layout: { type: 'string', description: 'Layout (páginas de conteúdo não podem repetir)' },
              image_queries: {
                type: 'array',
                items: { type: 'string' },
                description:
                  "Palavras-chave de busca de imagens em inglês para os espaços de imagem desta página (o sistema busca internamente e preenche as URLs reais); se você já conhece URLs http(s) reais, passe-as diretamente (respeitadas, sem nova busca); [] para nenhuma imagem",
              },
            },
            required: ['title', 'brief', 'layout'],
          },
        },
        insert_mode: {
          type: 'string',
          enum: ['replace', 'append'],
          description:
            'replace (padrão, deck novo inteiro) = substitui tudo; append = acrescenta ao final',
        },
        style_template: {
          type: 'string',
          description:
            "Opcional: nome de um modelo de estilo salvo (de list_style_templates); quando informado, a Etapa 0 é ignorada e o styleSkill do modelo é usado diretamente, sem regeneração de estilo. Um modelo salvo a partir de um deck concluído também traz seu esqueleto de layout, que fixa a geometria de título/imagem de marca em todas as páginas para um visual consistente",
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Obrigatório quando topic/context/briefs contêm números específicos (%, dinheiro, magnitudes): de onde vieram — 'user'/'document'/'search' (execute web_search primeiro)/'sample' (divulgue ao usuário)",
        },
      },
    },
  },
  {
    name: 'save_style_template',
    description:
      '[Salva o estilo do deck atual como um modelo reutilizável] Salva o Style Skill (guia de estilo visual) da apresentação atual com o nome informado, além do esqueleto de layout do deck (caixa de título recorrente, espaço da imagem de marca, formas de destaque e fundos). Na próxima vez que gerar um deck, passe o argumento style_template para reutilizar ambos diretamente e ignorar a geração de estilo. Chame quando o usuário disser "salve este estilo" / "salve como modelo".',
    inputSchema: {
      type: 'object',
      properties: {
        name: {
          type: 'string',
          description: 'Nome do modelo (curto, ex.: "minimal-blue" ou "tech-dark")',
        },
      },
      required: ['name'],
    },
  },
  {
    name: 'list_style_templates',
    description:
      'Lista todos os modelos de estilo salvos (nome + topic + createdAt). Quando o usuário disser "use o estilo da última vez" ou "use algum modelo", chame isto primeiro para ver o que existe e depois passe o nome do modelo desejado no argumento style_template de generate_deck.',
    inputSchema: { type: 'object', properties: {}, required: [] },
  },
  {
    name: 'edit_chart',
    description:
      'Modifica um gráfico (inclusive gráficos de arquivos importados; a primeira edição converte-o automaticamente em editável): altera tipo/dados/cores/elementos do gráfico. Opções de kind: bar/barStacked/line/area/pie/doughnut. colorScheme: default/colorful/colorful2/mono-accent1..6 (derivados do tema); as chaves antigas blue/warm/cool/mono ainda funcionam.',
    inputSchema: {
      type: 'object',
      properties: {
        slideIndex: { type: 'integer' },
        sourceId: { type: 'string', description: 'ID do elemento de gráfico (type=chart)' },
        kind: {
          type: 'string',
          enum: ['bar', 'barStacked', 'line', 'area', 'pie', 'doughnut'],
          description: 'Altera o tipo de gráfico (opcional)',
        },
        categories: {
          type: 'array',
          items: { type: 'string' },
          description: 'Rótulos do eixo X/categorias (opcional)',
        },
        series: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              name: { type: 'string' },
              values: { type: 'array', items: { type: 'number' } },
            },
            required: ['name', 'values'],
          },
          description: 'Séries de dados (opcional)',
        },
        dataSource: {
          type: 'string',
          enum: ['user', 'document', 'search', 'sample'],
          description:
            "Obrigatório ao passar series: procedência dos valores ('user'/'document'/'search'/'sample'; 'search' exige um web_search prévio; 'sample' precisa ser divulgado ao usuário)",
        },
        colorScheme: {
          type: 'string',
          description: 'Esquema de cores (opcional): default/colorful/colorful2/mono-accent1..6',
        },
        title: { type: 'string', description: 'Título do gráfico (opcional)' },
        legendPos: {
          type: 'string',
          enum: ['b', 't', 'r', 'l', 'none'],
          description: 'Posição da legenda (opcional)',
        },
        dataLabels: { type: 'boolean', description: 'Alterna rótulos de dados (opcional)' },
        gridlines: { type: 'boolean', description: 'Alterna linhas de grade do eixo de valores (opcional)' },
        switchRowCol: {
          type: 'boolean',
          description: 'Inverte linhas/colunas: categorias ↔ séries (opcional)',
        },
      },
      required: ['slideIndex', 'sourceId'],
    },
  },
  {
    name: 'apply_ops',
    description:
      '[Superfície canônica de edição] Aplica uma lista de operações canônicas de edição como UMA transação — atômica por padrão: qualquer falha reverte tudo, nada é aplicado pela metade. Defina dry_run:true para validar o plano sem tocar no deck (teste lotes arriscados). Esta é A ferramenta para toda edição sem uma ferramenta dedicada (preenchimento, contorno, transformação, exclusão, ordem-z, agrupamento, recorte, opacidade, efeitos, links, células de tabela, estrutura e estilo, fundo da página, notas do apresentador, exclusão/movimentação/duplicação de página, transições, seções, tema) e para lotes de várias páginas ou muitos elementos; uma única operação é um lote perfeitamente válido. Para cálculos de layout de uma página, prefira execute_slide_script.\n' +
      'Endereçamento: toda operação usa target:{slide, el?} — slide = índice baseado em 0 ou o identificador durável "s_<n>"; el = o ID de um elemento do esquema/read_slide (ids e_* são duráveis). Filhos de grupo: coloque o ID do filho em target.el e adicione group:"<group id>".\n' +
      'As unidades são EMU no espaço do documento. read_slide informa px e seu fator exato "1 px = N EMU" — converta com esse N (9525 apenas em um deck 16:9 padrão; outros tamanhos de página diferem). Tamanhos de fonte são em pt. setText em um elemento existente: trechos com apenas `text` herdam a formatação original do trecho (tamanho, cor, fonte) — deixe fontSize/color/fontFamily de fora a menos que o usuário tenha pedido para alterar o estilo.\n' +
      'Referência completa das operações (o mesmo executor usado por toda superfície de edição), uma assinatura por operação; ? marca campos opcionais:\n' +
      opSignatureIndex() +
      '\naddChart também aceita dataSource:"user"|"document"|"search"|"sample" (procedência dos números, verificada antes da execução do lote; "search" exige um web_search nesta conversa).' +
      '\nPara tabelas de campos, exemplos JSON executáveis e erros comuns, chame load_guide com os nome(s) de grupo acima antes de um lote que você ainda não fez. ' +
      "O erro de uma operação que falha também devolve sua assinatura exata de uma linha, e dry_run testa o lote inteiro sem tocar no deck. Um nome de operação desconhecido devolve o vocabulário completo.",
    inputSchema: {
      type: 'object',
      properties: {
        ops: {
          type: 'array',
          items: { type: 'object' },
          description: 'A lista de operações, aplicadas em ordem como uma transação (no máximo 50)',
        },
        dry_run: { type: 'boolean', description: 'Valida apenas o plano; o deck permanece intacto' },
        isolation: {
          type: 'string',
          enum: ['atomic', 'per_op'],
          description: 'atomic (padrão): tudo ou nada. per_op: operações independentes, falhas são ignoradas.',
        },
      },
      required: ['ops'],
    },
  },
  {
    name: 'load_guide',
    description:
      'Carrega a documentação completa de um ou mais grupos de operações para o contexto: tabelas de campos com tipos e unidades, exemplos JSON executáveis, erros comuns. Chame antes de um lote apply_ops que use operações que você ainda não usou nesta conversa; vários grupos podem ser carregados de uma vez.\n' +
      'Grupos disponíveis:\n' +
      opGuideCatalog(),
    inputSchema: {
      type: 'object',
      properties: {
        groups: {
          type: 'array',
          items: { type: 'string', enum: [...OP_GROUPS] },
          description: 'Nomes dos grupos a carregar, ex.: ["element","slide"]',
        },
      },
      required: ['groups'],
    },
  },
]

/** Collect readable text of nodes (including nested group children); returns a list of [sourceId, type, text] */
/** Find one node by id in the node tree (including groups). */
function findNodeById(nodes: RenderNode[], id: string): RenderNode | undefined {
  for (const n of nodes) {
    if (n.sourceId === id || n.durableId === id) return n
    if (n.type === 'group') {
      const hit = findNodeById(n.children, id)
      if (hit) return hit
    }
  }
  return undefined
}

/**
 * Editable context of an element: a top-level node, or a direct child of a top-level group
 * (with groupId + the group's absolute origin for abs↔group-local px conversion — child render
 * boxes are group-local, matching the in-group edit IPCs). Deeper nesting returns {nested:true}:
 * the main process patches one level only, so those stay read-only until ungrouped.
 */
type EditTarget =
  { node: RenderNode; groupId?: string; groupOrigin?: { x: number; y: number } } | { nested: true }
function resolveEditTarget(slide: RenderSlide, sourceId: string): EditTarget | null {
  const matches = (n: RenderNode) => n.sourceId === sourceId || n.durableId === sourceId
  for (const n of slide.nodes) {
    if (matches(n)) return { node: n }
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      const child = g.children.find(matches)
      if (child)
        return {
          node: child,
          groupId: n.sourceId,
          groupOrigin: { x: Math.round(n.box.x), y: Math.round(n.box.y) },
        }
      if (findNodeById(g.children, sourceId)) return { nested: true }
    }
  }
  return null
}

/** Editable element ids of a slide (locked layout decorations excluded), with short text hints. */
function availableIdList(slide: RenderSlide, withText = false): string {
  return collectNodeInfos(slide.nodes)
    .filter((n) => !n.locked)
    .map((n) => (withText && n.text?.trim() ? `${n.id} ("${preview(n.text, 12)}")` : n.id))
    .join(', ')
}

/**
 * Shared not-found / nested-in-subgroup error text for element-targeting tools. Guided: the
 * not-found branch lists the ids that DO exist — models self-correct from the list, but
 * blind-retry a bare "not found".
 */
function targetError(
  target: EditTarget | null,
  sourceId: string,
  pageNo: number,
  slide?: RenderSlide,
): string | null {
  if (!target) {
    const avail = slide ? ` Elements on this page: [${availableIdList(slide)}].` : ''
    return `Element ${sourceId} not found on page ${pageNo}.${avail} (e_* ids are durable across edits/saves; call read_slide when in doubt)`
  }
  if ('nested' in target)
    return `Element ${sourceId} is nested inside a sub-group; ungroup the outer group first (apply_ops ungroupElement), or edit the sub-group as a whole`
  return null
}

/** Element info shared by outline/read_slide/edit scripts (includes absolute geometry; locked = layout decoration, read-only). */
type NodeInfo = LayoutScriptElement

/**
 * The element's text as the model reads it.
 *
 * Delegates to `textForModel`, which substitutes a withheld span for its marker.
 * `forDisplay` defaults to false, so this whole file gets the redacted view
 * without having to know withholding exists; the reader's own find-and-replace
 * asks for the raw words explicitly.
 */
function nodeText(n: RenderNode, forDisplay = false): string {
  return textForModel(n, forDisplay)
}

/** Max font size of the text (pt, converted back from px); returns undefined when there is no text. */
function nodeMaxFontPt(n: RenderNode): number | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  let maxPx = 0
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) if (r.fontSizePx > maxPx) maxPx = r.fontSizePx
  }
  return maxPx > 0 ? Math.round((maxPx * 72) / 96) : undefined
}

/** Normalize a render color to #RRGGBB (strips alpha); undefined when not a hex color. */
function hex6(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#([0-9a-fA-F]{6})/.exec(color.trim())
  return m ? `#${m[1].toUpperCase()}` : undefined
}

/** Dominant text color = the run color covering the most characters (bullets excluded). */
function dominantTextColor(n: RenderNode): string | undefined {
  if (n.type !== 'shape' && n.type !== 'text') return undefined
  const weight = new Map<string, number>()
  for (const line of (n as ShapeRenderNode).text?.lines ?? []) {
    for (const r of line.runs) {
      if (r.isBullet) continue
      const c = hex6(r.color)
      if (c) weight.set(c, (weight.get(c) ?? 0) + r.text.length)
    }
  }
  let best: string | undefined
  let max = 0
  for (const [c, w] of weight) {
    if (w > max) {
      best = c
      max = w
    }
  }
  return best
}

/** Readable colors of a node (solid fill / dominant text color / stroke); pictures only expose stroke. */
function nodeColors(n: RenderNode): Pick<NodeInfo, 'fill' | 'textColor' | 'strokeColor'> {
  const out: Pick<NodeInfo, 'fill' | 'textColor' | 'strokeColor'> = {}
  if (n.type === 'shape' || n.type === 'text') {
    const s = n as ShapeRenderNode
    if (s.fill.kind === 'solid') {
      const c = hex6(s.fill.color)
      if (c) out.fill = c
    }
    const stroke = hex6(s.stroke?.color)
    if (stroke) out.strokeColor = stroke
    const text = dominantTextColor(n)
    if (text) out.textColor = text
  } else if (n.type === 'picture') {
    const stroke = hex6((n as PictureRenderNode).stroke?.color)
    if (stroke) out.strokeColor = stroke
  }
  return out
}

/**
 * Collect node info (including nested group children). A child's box is in group-local
 * coordinates (ext/chExt scaling already baked into geometry at build time); here we add the
 * group offset to convert to absolute coordinates and set the inGroup flag. Direct children of a
 * top-level group also carry groupId (editable via the in-group pipeline); deeper nesting stays
 * read-only (the main process patches one level only).
 */
function collectNodeInfos(
  nodes: RenderNode[],
  ox = 0,
  oy = 0,
  parent?: { id: string; topLevel: boolean },
): NodeInfo[] {
  const out: NodeInfo[] = []
  for (const n of nodes) {
    const b = n.box
    const abs = {
      x: Math.round(ox + b.x),
      y: Math.round(oy + b.y),
      w: Math.round(b.w),
      h: Math.round(b.h),
    }
    const base: NodeInfo = {
      // Durable id when the element's bytes carry one — survives regenerate/
      // ungroup/save, so the AI can keep addressing across turns
      id: n.durableId ?? n.sourceId,
      type: n.type,
      text: nodeText(n),
      ...abs,
      rotation: b.rotationDeg,
      ...(parent ? { inGroup: true } : {}),
      ...(parent?.topLevel ? { groupId: parent.id } : {}),
      ...(n.decoration ? { locked: true } : {}),
      ...nodeColors(n),
    }
    const fontPt = nodeMaxFontPt(n)
    if (fontPt !== undefined) base.fontSizePt = fontPt
    out.push(base)
    if (n.type === 'group') {
      const g = n as GroupRenderNode
      out.push(
        ...collectNodeInfos(g.children, abs.x, abs.y, {
          id: n.durableId ?? n.sourceId,
          topLevel: !parent,
        }),
      )
    }
  }
  return out
}

function preview(text: string, max = 50): string {
  const flat = text.replace(/\n/g, ' / ')
  return flat.length > max ? `${flat.slice(0, max)}…` : flat
}

function buildDeckOutline(slides: RenderSlide[], current: number, selectedIds: string[]): string {
  const canvas = slides[0] ? `Canvas ${slides[0].widthPx}×${slides[0].heightPx}px.` : ''
  const lines: string[] = [
    `The presentation has ${slides.length} pages; page ${current + 1} is currently shown. ${canvas}`,
    `(Page order is the current actual order and may differ from generation time or earlier conversation; the user's "page N" refers to this outline)`,
  ]
  if (selectedIds.length > 0) {
    const currentSlide = slides[current]
    const selectedRefs = selectedIds.map((id) => {
      const node = currentSlide ? findNodeById(currentSlide.nodes, id) : undefined
      return node?.durableId ?? node?.sourceId ?? id
    })
    lines.push(`User selected elements: ${selectedRefs.join(', ')}`)
  }
  slides.forEach((slide, i) => {
    lines.push(`Page ${i + 1} (slideIndex=${i}):`)
    const infos = collectNodeInfos(slide.nodes)
    const fillCount = new Map<string, number>()
    for (const n of infos) {
      if (n.fill) fillCount.set(n.fill, (fillCount.get(n.fill) ?? 0) + 1)
    }
    const mainFills = [...fillCount.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([c, count]) => (count > 1 ? `${c}×${count}` : c))
    if (mainFills.length > 0) lines.push(`  main fills: ${mainFills.join(' ')}`)
    for (const n of infos) {
      lines.push(`  - ${n.id} | ${n.type}${n.text ? ` | "${preview(n.text)}"` : ''}`)
    }
  })
  lines.push('(Use read_slide to see element positions/sizes/colors)')
  return lines.join('\n')
}

/**
 * Element inventory of one slide as read_slide reports it (ids + geometry + colors + text).
 * Shared by the read_slide tool and the post-generation layout QC pass (slide-qc.ts), so the
 * QC model maps screenshot pixels back to the same ids/coordinates the edit tools accept.
 */
export function formatSlideDump(slide: RenderSlide): string {
  const infos = collectNodeInfos(slide.nodes)
  const parts = infos.map((n) => {
    const flags = [
      n.groupId
        ? `in group ${n.groupId} (directly editable)`
        : n.inGroup
          ? 'nested in a sub-group (read-only; apply_ops ungroupElement on the outer group to edit)'
          : '',
      n.locked ? 'layout decoration (read-only)' : '',
    ]
      .filter(Boolean)
      .join(' ')
    const rot = n.rotation ? ` rotation ${Math.round(n.rotation)}°` : ''
    const font = n.fontSizePt ? ` font ${n.fontSizePt}pt` : ''
    const colors = [
      n.fill ? `fill${n.fill}` : '',
      n.textColor ? `text${n.textColor}` : '',
      n.strokeColor ? `stroke${n.strokeColor}` : '',
    ]
      .filter(Boolean)
      .join(' ')
    const head = `${n.id} | ${n.type}${flags ? ` | ${flags}` : ''} | pos(${n.x},${n.y}) size ${n.w}×${n.h}${rot}${font}${colors ? ` | ${colors}` : ''}`
    return n.text ? `${head}\n${n.text}` : `${head} | (no text)`
  })
  const colorlessTypes = [
    ...new Set(
      infos
        .filter((n) => !n.fill && !n.textColor && !n.strokeColor)
        .filter((n) => n.type === 'picture' || n.type === 'chart')
        .map((n) => n.type),
    ),
  ]
  const colorNote = colorlessTypes.length
    ? `\n(${colorlessTypes.join('/')} colors not available)`
    : ''
  // Report the real px→EMU factor: render px carry the viewport scale, so ×9525 only
  // holds for decks whose baseline width is exactly the fit width (standard 16:9 at 1280).
  const pxToEmu = +(9525 / slide.scale).toFixed(2)
  return `Canvas ${slide.widthPx}×${slide.heightPx}px (1 px = ${pxToEmu} EMU)\n${parts.join('\n---\n') || '(no elements on this page)'}${colorNote}`
}

/** tools that need a media provider: RamonMods login + cloud tools, or a BYOK media key in Settings */
function hiddenMediaTools(access: DeckAccess): Set<string> {
  const hidden = new Set<string>()
  if (access.imageGenAvailable?.() === false) hidden.add('generate_image')
  if (access.mediaAnalysisAvailable?.() === false) hidden.add('analyze_media')
  return hidden
}

function mediaToolsOffNote(hidden: Set<string>): string {
  if (hidden.size === 0) return ''
  const plural = hidden.size > 1
  return `\n\nNote: ${[...hidden].join(' and ')} ${plural ? 'are' : 'is'} currently unavailable (no image/media provider: signed out of RamonMods or cloud tools off, and no media API key in Settings). Do not call or promise ${plural ? 'them' : 'it'}; for imagery use image_search + insert_web_image instead.`
}

export function createSlidesSkill(access: DeckAccess): AgentSkill {
  // The HTML pipeline was already used in this conversation → later calls without an explicit mode default to append.
  /**
   * Every label withheld anywhere in the deck, in first-seen order.
   *
   * Read from the render tree the tools already use, so the instruction and the
   * text the model is shown cannot disagree about what is withheld.
   */
  function deckRedactLabels(access: DeckAccess): string[] {
    const seen = new Set<string>()
    for (const slide of access.getSlides()) {
      for (const label of redactLabelsOf(slide)) seen.add(label)
    }
    return [...seen]
  }

  // Safety net for when the AI ignores the "pass all pages at once" constraint: separate calls no longer overwrite each other (P0-1).
  const state: SkillState = { htmlGenerated: false }
  return {
    id: 'slides',
    // live like tools: the off-note overrides the prose that still mentions the hidden tools
    get systemPrompt() {
      // spans the reader withheld are named here, from the live deck, so the
      // model can see the marker exists and write around it. A deck with none
      // pays nothing and is not told to look for markers that are not there.
      const labels = deckRedactLabels(access)
      const redact = labels.length > 0 ? `\n\n${placeholderInstruction(labels)}` : ''
      return systemPrompt + mediaToolsOffNote(hiddenMediaTools(access)) + redact
    },
    // live view: the predicates are re-read before every model request
    get tools() {
      const hidden = hiddenMediaTools(access)
      return hidden.size ? TOOLS.filter((t) => !hidden.has(t.name)) : TOOLS
    },
    buildContext: () => {
      const outline = `<deck outline>\n${buildDeckOutline(access.getSlides(), access.getCurrent(), access.getSelectedIds())}\n</deck outline>`
      const progress = buildProgressNote(state)
      return progress ? `${outline}\n${progress}` : outline
    },
    executeTool: (call, signal) => executeTool(access, call, state, signal),
  }
}

interface SkillState {
  htmlGenerated: boolean
  /** A web_search ran in this conversation — unlocks dataSource:'search' in the figure gate */
  webSearched?: boolean
  /** Number of pages most recently planned by plan_deck, used by the progress checklist to remind the AI to finish */
  plannedPages?: number
  /** Per-page titles planned by plan_deck (order = page order), used to name unfinished pages in the checklist */
  plannedTitles?: string[]
  /** Whether each planned page has been generated (aligned with plannedTitles) — names unfinished pages accurately even when a middle page fails */
  pageDone?: boolean[]
  /** Style Skill produced by the most recent generate_deck (used by save_style_template) */
  lastStyleSkill?: string
  /** Topic of the most recent generate_deck (used by save_style_template) */
  lastTopic?: string
}

/**
 * [Hard constraint against "hand-building from scratch"] Sometimes the AI skips the HTML
 * pipeline and assembles a whole deck element by element with addElement/addSmartArt ops —
 * such hand-built pages look crude and the layout falls apart (root cause of screenshot issues).
 * "From-scratch" detection: this session hasn't used the HTML pipeline (!htmlGenerated) AND the
 * deck has almost no real content (≤ 2 non-decoration elements with text, i.e. blank/initial
 * template). If so, reject and steer toward generate_deck. Adding a single element to an
 * existing rich deck / fine-tuning after the HTML pipeline are unaffected.
 */
function blockScratchBuild(
  kind: 'element' | 'smartart',
  slides: RenderSlide[],
  state?: SkillState,
): { output: string; isError: true; mutated: false; summary: string } | null {
  if (state?.htmlGenerated) return null // Went through the HTML pipeline; subsequent native edits are legitimate
  let contentEls = 0
  for (const slide of slides) {
    for (const n of collectNodeInfos(slide.nodes)) {
      if (!n.locked && n.text && n.text.trim() !== '') contentEls += 1
    }
  }
  if (contentEls > 2) return null // Deck already has real content; this is a refinement scenario, allow it
  const label = kind === 'smartart' ? t('aiLabelInsertSmartart') : t('aiFailNewElement')
  return {
    output:
      "For blank/from-scratch scenarios don't hand-assemble pages element by element with addElement/addSmartArt/addTable/addChart ops (crude layout). " +
      'Use the generation pipeline instead: new whole deck → generate_deck; new pages for an existing deck → generate_deck(pages, insert_mode:"append"). ' +
      'The pipeline designs each page and builds it into editable elements. Use insert ops only when the deck already has polished content and one element needs refining.',
    isError: true,
    mutated: false,
    summary: t('aiSumFromScratchGuard', { label }),
  }
}

/**
 * Build the generation progress checklist text — injected to the AI each turn via buildContext
 * so it "sees" which pages are still missing, a mechanical reminder to finish (rather than a
 * one-shot prompt constraint). Returns an empty string when there is no plan.
 */
function buildProgressNote(state?: SkillState): string {
  if (!state || !state.plannedPages) return ''
  const planned = state.plannedPages
  const flags = state.pageDone ?? new Array(planned).fill(false)
  const done = flags.filter(Boolean).length
  const titles = state.plannedTitles ?? []
  if (done >= planned) {
    return `<generation-progress>\n✅ Complete: ${planned} pages planned, all generated (${done} pages).\n</generation-progress>`
  }
  // Name unfinished pages one by one from pageDone (page numbers stay accurate when a middle page fails)
  const remaining: string[] = []
  let omitted = 0
  for (let i = 0; i < planned; i++) {
    if (flags[i]) continue
    if (remaining.length < MAX_PROGRESS_NOTE_PAGES) {
      remaining.push(`page ${i + 1}${titles[i] ? ` "${titles[i]}"` : ''}`)
    } else omitted++
  }
  if (omitted > 0) remaining.push(`and ${omitted} more`)
  return (
    `<generation-progress>\n` +
    `⚠️ Incomplete: ${planned} pages planned, ${done} generated, ${planned - done} still missing.\n` +
    `Unfinished: ${remaining.join(', ')}.\n` +
    `Immediately fill in the unfinished pages above with generate_deck(pages: briefs for the missing pages, insert_mode:"append"); do not stop and do not substitute native tools.\n` +
    `</generation-progress>`
  )
}

const fail = (summary: string, output: string) => ({
  output,
  isError: true,
  mutated: false,
  summary,
})

// ── Figure-provenance gate ────────────────────────────────────
// Prompt rules ("search before writing data") did not stop invented numbers being
// delivered as fact, so provenance is enforced at the tool layer: chart data and
// data-dense briefs must declare a dataSource, 'search' is only accepted after a
// real web_search in this conversation, and 'sample' figures must be disclosed.

/** Specific figures: percentages, money, magnitude units, decimals — not bare years/counts */
const SPECIFIC_FIGURE_RE =
  /\d[\d,.]*\s*(?:%|％|亿|萬|万|兆|billion|million|\bbn\b|\bmn\b)|[¥￥$€£]\s*\d|\d+\.\d+/g

function countSpecificFigures(text: string): number {
  return text.match(SPECIFIC_FIGURE_RE)?.length ?? 0
}

/**
 * Returns an error message when the declared dataSource does not justify the figures
 * this call carries, null when the call may proceed.
 */
function dataSourceGateError(call: AgentToolCall, state: SkillState | undefined): string | null {
  return dataSourceGateErrorForInput(call.input, state)
}

/** Same gate over a plain argument object (an apply_ops op carries dataSource as an extra field). */
function dataSourceGateErrorForInput(
  input: Record<string, unknown>,
  state: SkillState | undefined,
): string | null {
  const src = String(input.dataSource ?? '')
  if (src === 'user' || src === 'document' || src === 'sample') return null
  if (src === 'search') {
    if (state?.webSearched) return null
    return (
      "dataSource is 'search' but no web_search has run in this conversation. " +
      'Run web_search first and build the figures from the results (cite them to the user), ' +
      "or declare 'user'/'document' if the figures actually came from the user or this deck."
    )
  }
  return (
    'This call carries specific figures, so dataSource is required: ' +
    "'user' (figures supplied by the user or attachments), 'document' (read from this deck), " +
    "'search' (from web_search results — run it first), or 'sample' (illustrative placeholders; " +
    'you must tell the user they are NOT real data). Never present invented numbers as facts.'
  )
}

/** Appended to a successful tool output when the model declared the figures illustrative. */
const SAMPLE_DATA_NOTE =
  '\nNOTE: dataSource is "sample" — you MUST tell the user these figures are illustrative placeholders, not real data, and offer to research real numbers with web_search.'

/** Append the missing-image report to the tool output: the model learns which pages lack images and how to fix them, instead of silently treating it as success. */
function imageFailNote(fails?: { page: number; url: string }[]): string {
  if (!fails?.length) return ''
  const detail = fails.map((f) => `page ${f.page} (${f.url})`).join(', ')
  return `\n⚠️ Missing images: ${detail} failed to download/convert; those image slots are blank on the page. Re-run image_search with more generic English keywords, pick a working image, patch it onto the page with insert_web_image (slideIndex = page number - 1), then reply to the user.`
}

/**
 * Activity-chip label per op for single-op apply_ops runs, so the panel still
 * reads "Edit text" / "Move element" instead of a bare count now that the
 * dedicated tools are gone. Multi-op batches fall back to the count summary.
 */
const OP_LABEL_KEYS: Record<string, string> = {
  setText: 'aiOpEditText',
  setFont: 'aiOpEditFormat',
  setParagraphFormat: 'aiOpEditFormat',
  setTransform: 'aiOpMoveElement',
  setFill: 'aiOpSetFill',
  setStroke: 'aiOpSetStroke',
  deleteElement: 'aiOpDeleteElement',
  ungroupElement: 'aiOpUngroup',
  setPictureSrcRect: 'aiOpCropImage',
  setPictureOpacity: 'aiOpPictureOpacity',
  setNotes: 'aiOpSpeakerNotes',
  deleteSlide: 'aiOpDeleteSlide',
  setBackground: 'aiOpBackground',
  setTableCell: 'aiOpEditTable',
  tableStructure: 'aiOpTableStructure',
  tableMerge: 'aiOpTableStructure',
  setTableStyle: 'aiOpTableStyle',
  duplicateSlide: 'aiOpNewSlide',
  addElement: 'aiOpNewShape',
  addChart: 'aiOpInsertChart',
  addSmartArt: 'aiOpInsertSmartart',
  addTable: 'aiOpInsertTable',
}

/** Localized summary for an apply_ops run: the op's label when exactly one op ran, else the count. */
function applyOpsSummary(ops: unknown[], count: number): string {
  if (count === 1 && ops.length === 1) {
    const op = ops[0] as { op?: unknown; kind?: unknown } | null
    const name = typeof op?.op === 'string' ? op.op : ''
    const key =
      name === 'addElement' && op?.kind === 'textbox' ? 'aiOpNewTextbox' : OP_LABEL_KEYS[name]
    if (key) return t(key as Parameters<typeof t>[0])
  }
  return t('aiSumApplyOps', { count })
}

/** Insert ops subject to the anti-scratch-build guard when the deck is still blank. */
const SCRATCH_GUARDED_OPS = new Set(['addElement', 'addSmartArt', 'addTable', 'addChart'])

interface OpPreflight {
  /** ops to send over the IPC (AI-layer fields stripped) */
  ops: unknown[]
  /** an addChart declared its figures illustrative — the output must say so */
  sampleData: boolean
}

/**
 * apply_ops runs the same policy gates the dedicated tools run before an op
 * reaches the executor: the anti-scratch-build guard for inserts on a blank
 * deck, and figure provenance for chart data. `dataSource` is an AI-layer
 * field the registry does not know; it is checked here and stripped before the
 * ops cross the IPC. Malformed entries pass through untouched so the executor
 * returns its own guided error for them.
 */
function preflightOps(
  opsIn: unknown[],
  slides: RenderSlide[],
  state: SkillState | undefined,
): OpPreflight | ReturnType<typeof fail> {
  const ops: unknown[] = []
  let sampleData = false
  for (const [i, raw] of opsIn.entries()) {
    if (
      typeof raw !== 'object' ||
      raw === null ||
      typeof (raw as { op?: unknown }).op !== 'string'
    ) {
      ops.push(raw)
      continue
    }
    const op = { ...(raw as Record<string, unknown>) }
    const name = op.op as string
    // A span withheld from the model is a mark over real words, so an edit that
    // rewrites or drops those words costs the reader data they chose to keep.
    // Checked here, before anything is dispatched — apply_ops is atomic, so a
    // refusal leaves the deck exactly as it was.
    const guard = redactGuardFor(op, slides)
    if (guard) {
      return fail(t('aiFailApplyOps'), `ops[${i}] ${name}: ${guard.reason}`)
    }
    if (SCRATCH_GUARDED_OPS.has(name)) {
      const blocked = blockScratchBuild(
        name === 'addSmartArt' ? 'smartart' : 'element',
        slides,
        state,
      )
      if (blocked) return { ...blocked, output: `ops[${i}] ${name}: ${blocked.output}` }
    }
    if (name === 'addChart') {
      const gateErr = dataSourceGateErrorForInput(op, state)
      if (gateErr) return fail(t('aiFailApplyOps'), `ops[${i}] ${name}: ${gateErr}`)
      if (op.dataSource === 'sample') sampleData = true
      delete op.dataSource
    }
    ops.push(op)
  }
  return { ops, sampleData }
}

/**
 * Layout audit of the pages a transaction touched (numeric slide prefixes of
 * the journal echo), capped so a deck-wide batch does not flood the reply.
 * Empty when every audited page passes.
 */
function auditTouchedPages(
  records: Array<{ target?: string }> | undefined,
  slides: RenderSlide[],
): string {
  const pages = new Set<number>()
  for (const rec of records ?? []) {
    const prefix = rec.target?.split('/')[0] ?? ''
    if (/^\d+$/.test(prefix)) {
      const idx = Number(prefix)
      if (slides[idx]) pages.add(idx)
    }
    if (pages.size >= 4) break
  }
  const failing = [...pages]
    .sort((a, b) => a - b)
    .map((idx) => ({ page: idx + 1, issues: auditSlideLayout(slides[idx]!) }))
    .filter((a) => a.issues.length > 0)
  if (failing.length === 0) return ''
  return (
    `\n<layout-audit>⚠️ Found issue(s) on ${failing.length} page(s):\n` +
    failing.map((a) => `page ${a.page}:\n${a.issues.map((s) => `- ${s}`).join('\n')}`).join('\n') +
    '\n→ Fix issues your edit caused before replying: execute_slide_script on the affected page (it reads live geometry) or another apply_ops with setTransform; at most 2 fix rounds. Issues that already existed and that you did not touch are for your judgment only — do not report them to the user, and never quote element ids in the reply.\n</layout-audit>'
  )
}

async function executeTool(
  access: DeckAccess,
  call: AgentToolCall,
  state?: SkillState,
  signal?: AbortSignal,
) {
  const slides = access.getSlides()
  switch (call.name) {
    case 'read_slide': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide)
        return fail(t('aiFailReadSlide'), `slideIndex out of range (0-${slides.length - 1})`)
      return {
        output: formatSlideDump(slide),
        mutated: false,
        summary: t('aiSumReadSlide', { n: idx + 1 }),
      }
    }

    case 'execute_layout_script':
    case 'execute_slide_script': {
      const idx = Number(call.input.slideIndex)
      const slide = slides[idx]
      if (!slide) return fail(t('aiFailScript'), `slideIndex out of range (0-${slides.length - 1})`)
      const code = String(call.input.code ?? '').trim()
      if (!code) return fail(t('aiFailScript'), 'code must not be empty')
      const infos = collectNodeInfos(slide.nodes)
      const r = runLayoutScript(code, infos, { w: slide.widthPx, h: slide.heightPx })
      const logsStr = r.logs.length ? `\nlog output:\n${r.logs.join('\n')}` : ''
      if (r.error) {
        return fail(
          t('aiFailScript'),
          `Script execution error: ${r.error}${logsStr}\n(You can fix the script and retry; see els for the element list and geometry)`,
        )
      }
      const returnedStr = r.returned !== undefined ? `\nScript returned: ${r.returned}` : ''
      if (r.ops.length === 0 && r.edits.length === 0) {
        return {
          output: `Script finished but called no edit primitives (setBox/moveBy/setText/setStyle/setFill/setStroke); the page was not modified.${returnedStr}${logsStr}`,
          mutated: false,
          summary: t('aiSumScriptNoop', { n: idx + 1 }),
        }
      }
      // ── Dispatch: the whole script is ONE transaction — the collected primitives go
      //   to the main process in a single IPC and apply atomically through the op
      //   executor (any failure rolls the deck back there and returns a guided error).
      const applied = await window.slidesApi.applyEditScript({
        slideIndex: idx,
        fitWidthPx: access.fitWidthPx,
        boxes: r.ops,
        edits: r.edits,
      })
      if (!applied || 'error' in applied) {
        return fail(
          t('aiFailScript'),
          `${applied && 'error' in applied ? applied.error : t('aiErrUnknown')}\n` +
            `The page is unchanged (the script is atomic). Elements on page ${idx + 1}: ` +
            `[${availableIdList(slide, true)}]. Fix the script and resend it whole.${returnedStr}${logsStr}`,
        )
      }
      access.applySlide(idx, applied.slide)
      const counts = { text: 0, style: 0, fill: 0, stroke: 0 }
      for (const e of r.edits) counts[e.kind] += 1
      const parts: string[] = []
      if (r.ops.length > 0) parts.push(`layout ${r.ops.length} element(s)`)
      if (counts.text > 0) parts.push(`text ${counts.text} item(s)`)
      if (counts.style > 0) parts.push(`style ${counts.style} item(s)`)
      if (counts.fill > 0) parts.push(`fill ${counts.fill} item(s)`)
      if (counts.stroke > 0) parts.push(`stroke ${counts.stroke} item(s)`)
      const issues = auditSlideLayout(applied.slide)
      const notesStr = applied.notes?.length ? `\nNote:\n${applied.notes.join('\n')}` : ''
      return {
        output: `Applied the edit script to page ${idx + 1}: ${parts.join(', ')}.${returnedStr}${logsStr}${notesStr}${formatAudit(issues)}`,
        mutated: true,
        summary: t('aiSumScript', { n: idx + 1, count: r.ops.length + r.edits.length }),
      }
    }

    case 'web_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailWebSearch'), 'query must not be empty')
      const r = await window.slidesApi.webSearch(query, Number(call.input.maxResults) || 6)
      // a backend failure must not read as "no results" — the model would fabricate conclusions
      if (r.method === 'error') {
        return fail(
          t('aiFailWebSearch'),
          `web search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
        )
      }
      if (state) state.webSearched = true
      // output for the LLM: title+URL+summary (each summary truncated to 120 chars to stay lean)
      const SNIPPET_MAX = 120
      const lines: string[] = []
      if (r.answer) lines.push(`Direct answer: ${r.answer.slice(0, 300)}\n`)
      r.results.forEach((it, i) => {
        const snip =
          it.snippet.length > SNIPPET_MAX ? it.snippet.slice(0, SNIPPET_MAX) + '…' : it.snippet
        lines.push(`${i + 1}. ${it.title}\n   ${it.url}\n   ${snip}`)
      })
      // display side channel: link list (full title+URL for the UI, not in LLM context)
      const display: ToolDisplay = {
        kind: 'links',
        items: r.results.map((it) => ({ url: it.url, title: it.title })),
      }
      return {
        output: lines.join('\n') || '(no results)',
        mutated: false,
        summary: t('aiSumWebSearch', { query, count: r.results.length }),
        display,
      }
    }

    case 'image_search': {
      const query = String(call.input.query ?? '').trim()
      if (!query) return fail(t('aiFailImageSearch'), 'query must not be empty')
      const r = await window.slidesApi.imageSearch(query, Number(call.input.maxResults) || 8)
      // a backend failure must not read as an empty gallery — the model would fabricate image choices
      if (r.method === 'error') {
        return fail(
          t('aiFailImageSearch'),
          `image search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
        )
      }
      // output for the LLM: keep the existing format (the LLM needs to read URLs into image_queries; format unchanged)
      const lines = r.images.map(
        (im, i) =>
          `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
      )
      // display side channel: image list (for UI thumbnails, not in LLM context)
      const display: ToolDisplay = {
        kind: 'images',
        items: r.images.map((im) => ({ url: im.imageUrl, title: im.title || undefined })),
      }
      return {
        output: lines.join('\n') || '(no images)',
        mutated: false,
        summary: t('aiSumImageSearch', { query, count: r.images.length }),
        display,
      }
    }

    case 'generate_image': {
      const prompt = String(call.input.prompt ?? '').trim()
      if (!prompt) return fail(t('aiFailGenImage'), 'prompt must not be empty')
      const refs = Array.isArray(call.input.referenceImageUrls)
        ? (call.input.referenceImageUrls as unknown[]).map(String).filter(Boolean)
        : undefined
      const r = await window.slidesApi.generateImage({
        prompt,
        model: call.input.model ? String(call.input.model) : undefined,
        referenceImageUrls: refs,
        aspectRatio: call.input.aspectRatio ? String(call.input.aspectRatio) : undefined,
        transparentBackground: call.input.transparentBackground === true,
      })
      if (!r.url) return fail(t('aiFailGenImage'), r.error ?? 'Generation failed')
      const display: ToolDisplay = {
        kind: 'images',
        items: [{ url: r.url, title: prompt.slice(0, 60) }],
      }
      return {
        output:
          `Image generated, URL: ${r.url}\n` +
          'New imagery: insert it with insert_web_image. If this edits an existing slide picture (e.g. background removal), swap it in place with replace_image instead.',
        mutated: false,
        summary: t('aiSumGenImage', {
          prompt: `${prompt.slice(0, 20)}${prompt.length > 20 ? '…' : ''}`,
        }),
        display,
      }
    }

    case 'analyze_media': {
      const mediaUrls = Array.isArray(call.input.mediaUrls)
        ? (call.input.mediaUrls as unknown[]).map(String).filter(Boolean)
        : []
      const requirements = String(call.input.requirements ?? '').trim()
      if (!mediaUrls.length) return fail(t('aiFailMedia'), 'mediaUrls must not be empty')
      if (!requirements) return fail(t('aiFailMedia'), 'requirements must not be empty')
      const r = await window.slidesApi.analyzeMedia({ mediaUrls, requirements })
      if (!r.text) return fail(t('aiFailMedia'), r.error ?? 'Analysis failed')
      // Analysis text can be very long; truncate to protect context (first 6000 chars are enough to generate deck content)
      const MAX_LEN = 6000
      const text = r.text.length > MAX_LEN ? r.text.slice(0, MAX_LEN) + '\n…(truncated)' : r.text
      return {
        output: text,
        mutated: false,
        summary: t('aiSumParseMedia', { count: mediaUrls.length }),
      }
    }

    case 'insert_web_image': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailInsertImage'), `slideIndex out of range (0-${slides.length - 1})`)
      const url = String(call.input.url ?? '')
      const attachmentName = attachmentRefName(url)
      let payload: { url: string } | { base64: string; ext: string }
      if (attachmentName != null) {
        if (!access.resolveAttachmentImage)
          return fail(t('aiFailInsertImage'), 'Attachment embedding is unavailable here')
        const resolved = await access.resolveAttachmentImage(attachmentName)
        if (!resolved.ok) return fail(t('aiFailInsertImage'), resolved.error)
        payload = { base64: resolved.base64, ext: resolved.ext }
      } else {
        // file:// = a BYOK-generated image in the local store (the main process only resolves its own files)
        if (!/^(https?|file):\/\//.test(url)) return fail(t('aiFailInsertImage'), 'Invalid url')
        payload = { url }
      }
      const r = await window.slidesApi.insertImageUrl({
        slideIndex: idx,
        ...payload,
        xPx: Number(call.input.x),
        yPx: Number(call.input.y),
        wPx: Number(call.input.w),
        hPx: Number(call.input.h),
        fitWidthPx: access.fitWidthPx,
      })
      if (!r)
        return fail(
          t('aiFailInsertImage'),
          'Download or insertion failed (the image may be inaccessible)',
        )
      access.applySlide(idx, r.slide)
      return {
        output: `Inserted the image on page ${idx + 1}, element id=${r.sourceId}.`,
        mutated: true,
        summary: t('aiSumInsertImage', { n: idx + 1 }),
      }
    }

    case 'replace_image': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      const failKey = 'aiFailReplaceImage' as const
      const slide = slides[idx]
      if (!slide) return fail(t(failKey), `slideIndex out of range (0-${slides.length - 1})`)
      const target = resolveEditTarget(slide, sourceId)
      const terr = targetError(target, sourceId, idx + 1, slides[idx])
      if (terr || !target || 'nested' in target) return fail(t(failKey), terr!)
      if (target.node.type !== 'picture')
        return fail(t(failKey), `Element ${sourceId} is not a picture (type: ${target.node.type})`)
      if (target.groupId)
        return fail(
          t(failKey),
          `Element ${sourceId} is inside a group; this tool only supports top-level pictures — ungroup it first (apply_ops ungroupElement)`,
        )

      const url = String(call.input.url ?? '')
      const attachmentName = attachmentRefName(url)
      let payload: { url: string } | { base64: string; ext: string }
      if (attachmentName != null) {
        if (!access.resolveAttachmentImage)
          return fail(t(failKey), 'Attachment embedding is unavailable here')
        const resolved = await access.resolveAttachmentImage(attachmentName)
        if (!resolved.ok) return fail(t(failKey), resolved.error)
        payload = { base64: resolved.base64, ext: resolved.ext }
      } else {
        if (!/^(https?|file):\/\//.test(url)) return fail(t(failKey), 'Invalid url')
        payload = { url }
      }
      const updated = await window.slidesApi.replacePictureUrl({
        slideIndex: idx,
        sourceId,
        ...payload,
        ...(call.input.keepCrop ? { keepSrcRect: true } : {}),
      })
      if (!updated)
        return fail(
          t(failKey),
          'Replacement failed (the image may be inaccessible, or the element is not a replaceable picture)',
        )
      access.applySlide(idx, updated)
      return {
        output: `Replaced the image of picture ${sourceId} on page ${idx + 1} in place (frame/z-order/effects kept).`,
        mutated: true,
        summary: t('aiSumReplaceImage', { n: idx + 1 }),
      }
    }

    case 'ask_clarification': {
      if (!access.askClarification)
        return fail(
          t('aiFailClarify'),
          'The current environment does not support questionnaire cards',
        )
      const raw = Array.isArray(call.input.questions) ? call.input.questions : []
      const questions: ClarifyQuestion[] = raw
        .map((q: Record<string, unknown>, i: number) => ({
          id: String(q.id ?? `q${i + 1}`),
          label: String(q.label ?? ''),
          description: q.description ? String(q.description) : undefined,
          options: Array.isArray(q.options)
            ? q.options.map((o: unknown) => String(o)).slice(0, 5)
            : [],
          multi: !!q.multi,
        }))
        .filter((q) => q.label && q.options.length > 0)
      if (questions.length === 0)
        return fail(
          t('aiFailClarify'),
          'questions must be non-empty and every question needs options',
        )
      const r = await access.askClarification(questions)
      if (r.cancelled) {
        return {
          output:
            'The user skipped the questionnaire. Decide the Core Hook and style yourself based on professional judgment and generate directly.',
          mutated: false,
          summary: t('aiSumClarifySkipped'),
        }
      }
      return {
        output: `User questionnaire answers:\n${r.answers}\nDecide the Core Hook and style accordingly, then generate with generate_deck.`,
        mutated: false,
        summary: t('aiSumClarifyDone'),
      }
    }

    case 'plan_deck': {
      const coreHook = String(call.input.core_hook ?? '').trim()
      const style = String(call.input.style ?? '').trim()
      const pages = Array.isArray(call.input.pages) ? call.input.pages : []
      if (!coreHook || !style || pages.length === 0) {
        return fail(t('aiFailPlan'), 'plan_deck requires core_hook + style + non-empty pages')
      }
      // Planning summary echoed back to the user
      const lines = pages.map((p: Record<string, unknown>, i: number) => {
        const q =
          Array.isArray(p.image_queries) && p.image_queries.length
            ? ` [images: ${p.image_queries.length}]`
            : ''
        return `Page ${i + 1} [${String(p.layout ?? '')}] ${String(p.title ?? '')} — ${String(p.brief ?? '').slice(0, 40)}${q}`
      })
      if (state) {
        state.plannedPages = pages.length
        state.plannedTitles = pages.map(
          (p: Record<string, unknown>) => String(p.title ?? '').trim() || 'Untitled',
        )
        state.pageDone = new Array(pages.length).fill(false) // New planning round, reset per-page progress
      }
      const summary = t('aiSumPlan', { count: pages.length, hook: coreHook })
      return {
        output: `Plan confirmed:\nCore Hook: ${coreHook}\nStyle: ${style}\n${lines.join('\n')}\nNow follow this plan and call generate_deck once, passing core_hook + style + all ${pages.length} pages (each page's brief strictly following the plan above). Each turn a <generation-progress> note tells you how many pages remain; do not stop before they are complete.`,
        mutated: false,
        summary,
      }
    }

    case 'regenerate_slide': {
      const idx = Number(call.input.slideIndex)
      if (!slides[idx])
        return fail(t('aiFailRegen'), `slideIndex out of range (0-${slides.length - 1})`)
      if (!access.regenerateSlide || !access.generatePageLocal)
        return fail(
          t('aiFailRegen'),
          'The current environment does not support the page-redo pipeline',
        )
      const brief = String(call.input.brief ?? '').trim()
      if (!brief) return fail(t('aiFailRegen'), 'brief must not be empty')
      // Figure-provenance gate
      if (countSpecificFigures(`${String(call.input.title ?? '')}\n${brief}`) >= 2) {
        const gateErr = dataSourceGateError(call, state)
        if (gateErr) return fail(t('aiFailRegen'), gateErr)
      }
      const regenImages = Array.isArray(call.input.image_urls)
        ? (call.input.image_urls as unknown[]).map(String).filter((u) => /^https?:\/\//.test(u))
        : []
      // One retry then give up (same semantics as generate_deck pages).
      const backoff = access.retryBackoffMs ?? 2000
      let marker: string | null = null
      let genImageFails: string[] = []
      let lastErr = ''
      const regenArgs = {
        pageIndex: idx + 1,
        totalPages: slides.length,
        coreHook: '',
        style: state?.lastStyleSkill ?? '',
        title: String(call.input.title ?? ''),
        brief,
        layout: String(call.input.layout ?? ''),
        images: regenImages,
        canvasW: 1280,
        canvasH: 720,
      }
      for (let attempt = 0; attempt < 2 && !marker; attempt++) {
        if (attempt > 0 && backoff > 0) await new Promise((r) => setTimeout(r, backoff))
        const res = await access.generatePageLocal(regenArgs)
        if (res.ok && res.marker) {
          marker = res.marker
          if (Array.isArray(res.imageFailures)) genImageFails = res.imageFailures
        } else lastErr = res.error ?? t('aiErrUnknown')
      }
      if (!marker)
        return fail(
          t('aiFailRegen'),
          `Page generation failed (2 attempts): ${lastErr}. This is usually a temporary service error — do not keep calling regenerate_slide in a loop. Instead, make the requested changes in place with execute_slide_script / set_element_* (group children are editable too), or tell the user to retry in a few minutes. The page was not modified.`,
        )
      const r = await access.regenerateSlide(idx, marker)
      if (!r.ok)
        return fail(
          t('aiFailRegen'),
          `${r.error || 'Redo failed'} (the page was not modified; retry once, or edit it in place with execute_slide_script)`,
        )
      if (state) state.htmlGenerated = true
      const allImageFails = [
        ...(r.imageFailures ?? []),
        ...genImageFails.map((url) => ({ page: idx + 1, url })),
      ]
      return {
        output:
          `Redid page ${idx + 1} in place from the brief (other pages untouched; the user can undo). Fine-tune afterwards with execute_slide_script / set_element_* tools.` +
          imageFailNote(allImageFails.length ? allImageFails : undefined),
        mutated: true,
        summary: t('aiSumRegen', { n: idx + 1 }),
      }
    }

    case 'generate_deck': {
      // ── Self-driven pipeline:
      //   1) Plan: use pages if passed; with topic, the tool plans the outline via LLM (batched recursion over threshold) — fixes missing pages at the input side.
      //   2) Generate: batched concurrent page generation (one retry per page), **each batch lands immediately → frontend shows pages one by one**.
      //      The LLM (app AI transport) writes a slide spec that is built directly into a pptx.
      if (!access.generatePageLocal)
        return fail(
          t('aiFailGenDeck'),
          'No page generation pipeline is available in this environment',
        )
      if (!access.landGeneratedPages)
        return fail(
          t('aiFailGenDeck'),
          'The current environment does not support the page landing pipeline',
        )

      // Hard gate: with unread text attachments present, refuse to generate.
      // The prompt already demands reading them first, but prompt rules alone are not
      // enforcement — this check is, and it is fully computable from the tool-call history.
      {
        const unread = access.unreadTextAttachments?.() ?? []
        if (unread.length > 0) {
          return fail(
            t('aiFailGenDeck'),
            `Text attachment(s) not read yet: ${unread.join(', ')}. Read each one with read_attachment first (paginate long files), then call generate_deck again and pass the key facts you read (names, figures, deals, next steps) in the context argument — deck content must come from the attachments, not generic filler.`,
          )
        }
      }

      const canvasW = 1280
      const canvasH = 720
      const insertMode: 'replace' | 'append' =
        call.input.insert_mode === 'append' ? 'append' : 'replace'
      const PLAN_BATCH = 12 // Per-batch planning cap (kept slightly conservative against truncation)
      const GEN_BATCH = 2 // Per-page generation concurrency (opus large output + proxy concurrent streams time out easily; lowered to 2, stability first)
      const BACKOFF_MS = access.retryBackoffMs ?? 2000 // Retry backoff base (rate limits/overload are mostly transient; immediate retries would hit them again)

      let coreHook = String(call.input.core_hook ?? '').trim()
      const style = String(call.input.style ?? '').trim()
      const pages: Array<Record<string, unknown>> = Array.isArray(call.input.pages)
        ? (call.input.pages as Array<Record<string, unknown>>)
        : []
      const topic = String(call.input.topic ?? '').trim()
      const context = String(call.input.context ?? '').trim() || undefined
      // Every per-page request re-sends the context; cap it so N pages don't multiply a huge attachment
      const PAGE_CONTEXT_MAX = 8000
      const pageContext =
        context && context.length > PAGE_CONTEXT_MAX ? context.slice(0, PAGE_CONTEXT_MAX) : context
      const styleTemplateName = String(call.input.style_template ?? '').trim() || undefined

      // Figure-provenance gate: a data-dense request must say where its numbers came from
      {
        const briefText = [
          topic,
          context ?? '',
          ...pages.map((p) => `${String(p.title ?? '')} ${String(p.brief ?? '')}`),
        ].join('\n')
        if (countSpecificFigures(briefText) >= 2) {
          const gateErr = dataSourceGateError(call, state)
          if (gateErr) return fail(t('aiFailGenDeck'), gateErr)
        }
      }

      // ── Step 0: generate the Style Skill independently — one focused LLM call thinking only about the design system, less AI-looking.
      // Prefer the user-passed style as the style preference; without pages, generate a full Style Skill from topic.
      // When full pages+style are passed, respect the user's style (don't regenerate).
      // When style_template is passed, load the template directly and skip Step 0 (no LLM style generation).
      let styleSkill = ''
      let templateSkeleton: LayoutSkeleton | undefined
      if (styleTemplateName && access.loadStyleTemplate) {
        // Preferred: load from a saved template (fail-open: on load failure continue normal generation)
        try {
          const tr = await access.loadStyleTemplate(styleTemplateName)
          if (tr.ok && tr.styleSkill) styleSkill = tr.styleSkill
          if (tr.ok && tr.layout) templateSkeleton = tr.layout
        } catch {
          /* fail-open */
        }
      }
      // User clicked stop → checked inside each stage loop, abort as soon as possible (pages already landed are kept)
      const cancelled = () => signal?.aborted === true
      const cancelResult = (landed: number, totalPages: number) => ({
        output: `The user stopped generation. ${landed} page(s) already landed${totalPages ? ` (${totalPages} planned)` : ''} and remain on the canvas.`,
        mutated: landed > 0,
        summary: landed > 0 ? t('aiSumStoppedKept', { n: landed }) : t('aiErrStopped'),
      })
      if (cancelled()) return cancelResult(0, 0)

      const needStyleGen =
        !styleSkill && access.generateStyleSkill && (topic || pages.length === 0 || !style)
      if (needStyleGen) {
        access.onProgress?.({
          stage: 'style',
          label: t('aiStageStyle'),
          status: 'running',
          summary: t('aiStageStyleRunning'),
        })
        const styleTopic =
          topic || coreHook || (pages[0] ? String(pages[0].title ?? '') : 'Presentation')
        const sr = await access.generateStyleSkill!({
          topic: styleTopic,
          ...(style ? { styleHint: style } : {}),
          ...(context ? { questionnaire: context } : {}),
          ...(signal ? { signal } : {}),
        })
        if (sr.ok && sr.styleSkill) styleSkill = sr.styleSkill
        access.onProgress?.({
          stage: 'style',
          label: t('aiStageStyle'),
          status: 'done',
          summary: t('aiStageStyleDone'),
        })
      }
      if (!styleSkill) styleSkill = style // Fallback: use the user-passed style, or empty

      // ── Step 1: plan the outline — without pages, plan in-tool from topic (batched recursion over PLAN_BATCH; layouts chosen per the Style Skill).
      const approxForProgress = Math.min(
        MAX_APPROX_PAGES,
        Math.max(1, parseInt(String(call.input.approx_pages ?? '0'), 10) || pages.length || 1),
      )
      if (pages.length === 0) {
        const approx = approxForProgress
        if (!topic || !approx) {
          return fail(
            t('aiFailGenDeck'),
            'generate_deck requires [topic + approx_pages] (system plans internally), or pass [core_hook + style + pages] directly.',
          )
        }
        if (!access.planDeckOutline)
          return fail(
            t('aiFailGenDeck'),
            'The current environment does not support internal planning; pass pages directly.',
          )
        access.onProgress?.({
          stage: 'plan',
          label: t('aiStagePlan'),
          done: 0,
          total: approx,
          status: 'running',
          summary: t('aiStagePlanRunning'),
        })
        let planned = 0
        while (planned < approx) {
          if (cancelled()) return cancelResult(0, approx)
          const count = Math.min(PLAN_BATCH, approx - planned)
          const r = await access.planDeckOutline({
            topic,
            count,
            startPage: planned + 1,
            ...(context ? { context } : {}),
            ...(styleSkill ? { styleSkill } : {}),
            ...(planned > 0 && coreHook ? { continueFrom: { coreHook } } : {}),
            ...(signal ? { signal } : {}),
          })
          if (
            !r.ok ||
            !r.outline ||
            !Array.isArray(r.outline.pages) ||
            r.outline.pages.length === 0
          ) {
            if (pages.length === 0) {
              // On failure the progress must be finalized, otherwise the UI spins forever at "Planning outline…"
              access.onProgress?.({
                stage: 'plan',
                label: t('aiStagePlan'),
                done: 0,
                total: approx,
                status: 'error',
                summary: t('aiStagePlanFailed'),
              })
              return fail(t('aiFailGenDeck'), `Planning failed: ${r.error || 'outline is empty'}`)
            }
            break // A later planning batch failed; at least generate what was already planned
          }
          if (planned === 0) {
            coreHook = String(r.outline.core_hook ?? coreHook).trim()
          }
          pages.push(...r.outline.pages)
          planned += r.outline.pages.length
          access.onProgress?.({
            stage: 'plan',
            label: t('aiStagePlan'),
            done: planned,
            total: approx,
            status: 'running',
            summary: t('aiStagePlanProgress', { done: planned, total: approx }),
          })
          if (r.outline.pages.length < count) break // The model returned fewer than requested; stop
        }
        access.onProgress?.({
          stage: 'plan',
          label: t('aiStagePlan'),
          done: pages.length,
          total: pages.length,
          status: 'done',
          summary: t('aiStagePlanDone', { n: pages.length }),
        })
      }

      // ── Step 1.5: in-tool image search —
      // walk every page's image_queries and replace "English keywords (non-URL)" with real image URLs.
      // Entries that are already http(s) URLs are respected upstream, not re-searched; pages whose search failed keep an empty array (fail-open).
      // The same keyword is searched once per deck (fetching several candidates at once); allocation skips already-used URLs to avoid duplicate images across pages.
      if (access.searchImages) {
        const isUrl = (s: string) => /^https?:\/\//i.test(s)
        const normKw = (s: string) => s.toLowerCase().replace(/\s+/g, ' ')
        // Collect deduplicated search keywords across the deck
        const uniqueKeywords: string[] = []
        const seenKw = new Set<string>()
        for (const p of pages) {
          if (!Array.isArray(p.image_queries)) continue
          for (const q of p.image_queries as unknown[]) {
            const s = String(q).trim()
            if (!s || isUrl(s)) continue
            const k = normKw(s)
            if (!seenKw.has(k)) {
              seenKw.add(k)
              uniqueKeywords.push(s)
            }
          }
        }
        // Each keyword is searched only once, fetching several candidates so cross-page dedup can pick unused images
        const candidatesByKw = new Map<string, string[]>()
        const totalSearches = uniqueKeywords.length
        if (totalSearches > 0) {
          access.onProgress?.({
            stage: 'images',
            label: t('aiStageImages'),
            done: 0,
            total: totalSearches,
            status: 'running',
            summary: t('aiStageImagesRunning', { done: 0, total: totalSearches }),
          })
          let imagesDone = 0
          await Promise.all(
            uniqueKeywords.map(async (kw) => {
              if (cancelled()) return
              try {
                const urls = await access.searchImages!(kw, 5)
                candidatesByKw.set(normKw(kw), urls)
              } catch {
                /* fail-open */
              }
              imagesDone++
              access.onProgress?.({
                stage: 'images',
                label: t('aiStageImages'),
                done: imagesDone,
                total: totalSearches,
                status: 'running',
                summary: t('aiStageImagesRunning', { done: imagesDone, total: totalSearches }),
              })
            }),
          )
        }
        // Allocate page by page in order: explicit URLs are kept as-is (and counted as used); keywords pick from candidates, preferring images the deck hasn't used yet
        const usedUrls = new Set<string>()
        for (const p of pages) {
          if (!Array.isArray(p.image_queries)) continue
          for (const q of p.image_queries as unknown[]) {
            const s = String(q).trim()
            if (isUrl(s)) usedUrls.add(s)
          }
        }
        for (const p of pages) {
          if (!Array.isArray(p.image_queries) || p.image_queries.length === 0) continue
          const resolvedUrls: string[] = []
          for (const q of p.image_queries as unknown[]) {
            const s = String(q).trim()
            if (!s) continue
            if (isUrl(s)) {
              resolvedUrls.push(s)
              continue
            }
            const candidates = candidatesByKw.get(normKw(s)) ?? []
            // If all candidates are used, fall back to reusing the first one (an image beats no image)
            const pick = candidates.find((u) => !usedUrls.has(u)) ?? candidates[0]
            if (pick) {
              usedUrls.add(pick)
              resolvedUrls.push(pick)
            }
          }
          p.image_queries = resolvedUrls
        }
        if (totalSearches > 0) {
          access.onProgress?.({
            stage: 'images',
            label: t('aiStageImages'),
            done: totalSearches,
            total: totalSearches,
            status: 'done',
            summary: t('aiStageImagesDone', { n: totalSearches }),
          })
        }
      }

      if (!coreHook) coreHook = topic || 'Presentation'
      if (!styleSkill)
        styleSkill =
          'Unified clean modern style: main background #FFFFFF, main text #1A1A2E, primary accent #2563EB, secondary accent #F59E0B, cards #F8FAFC, borders #E2E8F0; sans-serif fonts; titles 40-56px, body 16-20px'
      const total = pages.length
      if (total === 0) return fail(t('aiFailGenDeck'), 'No pages to generate (the plan is empty).')

      // Store the plan in the progress state (shared with the todolist mechanism; buildContext injects it every turn)
      // Also record styleSkill+topic for the save_style_template tool
      if (state) {
        state.plannedPages = total
        state.plannedTitles = pages.map((p) => String(p.title ?? '').trim() || 'Untitled')
        state.pageDone = new Array(total).fill(false)
        state.lastStyleSkill = styleSkill
        state.lastTopic = topic || coreHook || ''
      }

      // Presentation name: prefer the cover (page 1) title, then the user-entered topic / coreHook.
      // New drafts are saved under this name instead of "Untitled-timestamp".
      const deckName = String(pages[0]?.title ?? '').trim() || topic || coreHook

      // ── Step 2: generate page by page + land as we go (frontend shows pages one by one).
      // Each page (LLM spec → pptx-engine build) is a one-slide pptx temp file; genOne returns
      // its marker and landing reads the bytes.
      // Land strictly in page order: nextToLand pointer; a page lands only when its marker is ready, keeping page order intact.
      const markerByIndex: (string | null)[] = new Array(total).fill(null)
      // Per-page completion flags (aligned with pages; same reference as state.pageDone, used by buildContext progress injection)
      const doneFlags: boolean[] = state?.pageDone ?? new Array(total).fill(false)
      const genFailed: number[] = [] // Page indexes (0-based) whose page generation failed
      const landFailed: number[] = [] // Page indexes (0-based) whose page generated but landing failed
      const degraded: number[] = [] // Page indexes (0-based) that "landed" via the plain-text fallback — must be reported, otherwise dead pages appear silently
      const deckImageFails: { page: number; url: string }[] = [] // Image download/conversion failures (page numbers are deck-global 1-based)
      const pageErrors: (string | undefined)[] = new Array(total).fill(undefined) // Last failure reason per page
      let landedPages = 0
      let firstDone = false
      let baseOffset = 0 // Number of existing pages before generated page 0 in the deck (>0 in append mode); used to re-insert retries at their original position
      let nextToLand = 0 // Index of the next page to land (0-based)

      // Initialize per-page progress state (all pages pending)
      const pageProgressItems: PageProgressItem[] = pages.map((p) => ({
        title: String(p.title ?? '').trim() || t('aiPageN', { n: pages.indexOf(p) + 1 }),
        status: 'pending',
      }))

      // Send initial pages progress
      access.onProgress?.({
        stage: 'pages',
        label: t('aiStagePages'),
        done: 0,
        total,
        status: 'running',
        summary: t('aiStagePageRunning', { n: 1, total }),
        pages: [...pageProgressItems],
      })

      type PageGen = NonNullable<DeckAccess['generatePageLocal']>
      // Up to two attempts; returns the marker or the last error.
      const tryGenerate = async (
        pageArgs: Parameters<PageGen>[0],
        pageIndex: number,
      ): Promise<{ marker: string | null; err: string }> => {
        let lastErr = ''
        for (let attempt = 0; attempt < 2; attempt++) {
          if (cancelled()) return { marker: null, err: lastErr }
          if (attempt > 0 && BACKOFF_MS > 0) await new Promise((r) => setTimeout(r, BACKOFF_MS))
          const res = await access.generatePageLocal!(pageArgs)
          if (res.ok && res.marker) {
            if (Array.isArray(res.imageFailures))
              deckImageFails.push(...res.imageFailures.map((url) => ({ page: pageIndex, url })))
            return { marker: res.marker, err: '' }
          }
          lastErr = res.error ?? t('aiErrUnknown')
        }
        return { marker: null, err: lastErr }
      }

      const genOne = async (p: Record<string, unknown>, pageIndex: number) => {
        // Mark as running
        pageProgressItems[pageIndex - 1] = {
          ...pageProgressItems[pageIndex - 1]!,
          status: 'running',
        }
        access.onProgress?.({
          stage: 'pages',
          label: t('aiStagePages'),
          done: landedPages,
          total,
          status: 'running',
          summary: t('aiStagePageRunning', { n: pageIndex, total }),
          pages: [...pageProgressItems],
        })
        const images = Array.isArray(p.image_queries)
          ? (p.image_queries as unknown[])
              .map((x) => String(x))
              .filter((x) => /^https?:\/\//.test(x))
          : []
        // The template's chrome block for this page's role (cover/content/closing):
        // the exact boxes every page of the role must reuse for a consistent look
        const skeleton = templateSkeleton
          ? formatSkeletonForPrompt(templateSkeleton, skeletonRole(pageIndex - 1, total)) ||
            undefined
          : undefined
        const pageArgs = {
          pageIndex,
          totalPages: total,
          coreHook,
          style: styleSkill,
          title: String(p.title ?? ''),
          brief: String(p.brief ?? ''),
          layout: String(p.layout ?? ''),
          images,
          ...(pageContext ? { context: pageContext } : {}),
          ...(topic ? { topic } : {}),
          canvasW,
          canvasH,
          ...(skeleton ? { skeleton } : {}),
          ...(signal ? { signal } : {}),
        }
        const res = await tryGenerate(pageArgs, pageIndex)
        pageErrors[pageIndex - 1] = res.marker ? undefined : res.err
        return res.marker
      }

      // Land in page order: starting from nextToLand, land as many as possible (stop at a gap and wait for it to generate).
      const flushLanded = async () => {
        while (nextToLand < total && markerByIndex[nextToLand] !== null) {
          if (cancelled()) return
          const marker = markerByIndex[nextToLand] as string
          if (marker.length > 0) {
            const m: 'replace' | 'append' = firstDone ? 'append' : insertMode
            const r = await access.landGeneratedPages!([marker], m, deckName)
            if (r.ok) {
              if (r.fallbackReason) {
                degraded.push(nextToLand)
                pageErrors[nextToLand] = r.fallbackReason
              }
              if (r.imageFailures) deckImageFails.push(...r.imageFailures)
              if (!firstDone) baseOffset = r.appendedFrom ?? 0
              landedPages += 1
              firstDone = true
              doneFlags[nextToLand] = true
              pageProgressItems[nextToLand] = { ...pageProgressItems[nextToLand]!, status: 'done' }
              access.onProgress?.({
                stage: 'pages',
                label: t('aiStagePages'),
                done: landedPages,
                total,
                status: 'running',
                summary:
                  landedPages < total
                    ? t('aiStagePageRunning', { n: landedPages + 1, total })
                    : t('aiStageFinishing'),
                pages: [...pageProgressItems],
              })
            } else {
              landFailed.push(nextToLand)
              pageErrors[nextToLand] = t('aiErrLandFailed', { err: r.error ?? t('aiErrUnknown') })
              pageProgressItems[nextToLand] = {
                ...pageProgressItems[nextToLand]!,
                status: 'error',
                error: pageErrors[nextToLand],
              }
            }
          }
          nextToLand += 1
        }
      }

      for (let start = 0; start < total; start += GEN_BATCH) {
        if (cancelled()) break
        const batchIdxs = []
        for (let k = start; k < Math.min(start + GEN_BATCH, total); k++) batchIdxs.push(k)
        const results = await Promise.all(
          batchIdxs.map(async (idx) => ({ idx, marker: await genOne(pages[idx], idx + 1) })),
        )
        if (cancelled()) break
        for (const { idx, marker } of results) {
          if (marker && marker.length > 0) {
            markerByIndex[idx] = marker
          } else {
            markerByIndex[idx] = '' // Empty-string placeholder; doesn't block subsequent landings
            genFailed.push(idx)
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'error',
              error: pageErrors[idx],
            }
          }
        }
        // Batch generated → immediately land everything that can land (frontend shows pages one by one)
        await flushLanded()
      }
      if (!cancelled()) await flushLanded() // Finalize

      // ── One retry round for failed pages, re-inserted at their original page position with
      //   insert_at (target position = existing-page offset + pages completed before this one).
      //   Landing-failed pages re-land (cheap: the one-slide pptx already exists). Pages whose
      //   generation failed get one more attempt here (a JSON spec retry is cheap).
      if (!cancelled()) {
        const retryIdxs = [...new Set([...genFailed, ...landFailed])].sort((a, b) => a - b)
        for (const idx of retryIdxs) {
          if (cancelled()) break
          let marker = markerByIndex[idx]
          if (!marker) marker = await genOne(pages[idx]!, idx + 1)
          if (!marker) {
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'error',
              error: pageErrors[idx],
            }
            continue
          }
          const isFirstLand = !firstDone
          const r = isFirstLand
            ? await access.landGeneratedPages!([marker], insertMode, deckName)
            : await access.landGeneratedPages!(
                [marker],
                'insert_at',
                deckName,
                baseOffset + doneFlags.slice(0, idx).filter(Boolean).length,
              )
          if (r.ok) {
            if (isFirstLand) {
              baseOffset = r.appendedFrom ?? 0
              firstDone = true
            }
            landedPages += 1
            doneFlags[idx] = true
            pageErrors[idx] = r.fallbackReason
            if (r.fallbackReason) degraded.push(idx)
            if (r.imageFailures) deckImageFails.push(...r.imageFailures)
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'done',
              error: undefined,
            }
            access.onProgress?.({
              stage: 'pages',
              label: t('aiStagePages'),
              done: landedPages,
              total,
              status: 'running',
              summary: t('aiStagePageRestored', { n: idx + 1 }),
              pages: [...pageProgressItems],
            })
          } else {
            pageErrors[idx] = t('aiErrReinsertFailed', { err: r.error ?? t('aiErrUnknown') })
            pageProgressItems[idx] = {
              ...pageProgressItems[idx]!,
              status: 'error',
              error: pageErrors[idx],
            }
          }
        }
      }

      // User stopped midway: keep landed pages and finish honestly (no more retries / no failed-page reporting)
      if (cancelled()) {
        if (state) state.htmlGenerated = landedPages > 0 || state.htmlGenerated
        access.onProgress?.({
          stage: 'done',
          total: landedPages,
          summary: t('aiSumStoppedKept', { n: landedPages }),
          outcome: 'cancelled',
        })
        return cancelResult(landedPages, total)
      }

      if (state) state.htmlGenerated = true

      // ── Sidecar persistence (fail-open): write styleSkill to .styleskill.json next to the draft
      if (landedPages > 0 && access.saveSidecar && styleSkill) {
        try {
          await access.saveSidecar({
            topic: topic || coreHook || '',
            styleSkill,
            createdAt: new Date().toISOString(),
          })
        } catch {
          /* fail-open: a sidecar failure doesn't block */
        }
      }

      if (landedPages === 0) {
        access.onProgress?.({
          stage: 'done',
          total: 0,
          summary: t('aiStageAllFailed', { n: total }),
          outcome: 'failed',
        })
        return fail(
          t('aiFailGenDeck'),
          `All ${total} pages failed to generate; retry or check the AI model configuration.`,
        )
      }

      // Send completion progress event
      access.onProgress?.({
        stage: 'done',
        total: landedPages,
        summary: t('aiStageDoneSummary', { n: landedPages }),
      })

      const note = buildProgressNote(state)
      const progressTail = note ? `\n${note}` : ''
      // Faithfully report all unfinished pages (generation and landing failures; all already retried once)
      const stillFailed: number[] = []
      for (let i = 0; i < total; i++) if (!doneFlags[i]) stillFailed.push(i + 1)
      const briefErr = (s?: string) => (s ? (s.length > 80 ? `${s.slice(0, 80)}…` : s) : '')
      const okMsg = `Self-driven generation produced ${landedPages}/${total} pages (pages generated one by one and displayed as they land; failed pages were auto-retried).`
      const failDetail = stillFailed
        .map((n) => `page ${n}${pageErrors[n - 1] ? ` (${briefErr(pageErrors[n - 1])})` : ''}`)
        .join(', ')
      const failMsg = stillFailed.length
        ? ` ⚠️ ${failDetail} still failed after retry (these pages are missing from the deck; later pages shifted forward). To fill them in, call generate_deck again with briefs for just those pages and insert_mode:"append", and tell the user the page is at the end.`
        : ' Fine-tune with the set_element_* / add_* tools.'
      // Pages that went through the plain-text fallback: the page exists in the deck but all design is lost; the model must be told explicitly to redo it in place, not silently treat it as success
      const degradedMsg = degraded.length
        ? ` ⚠️ ${[...degraded]
            .sort((a, b) => a - b)
            .map(
              (i) =>
                `page ${i + 1} (canvas slideIndex=${baseOffset + doneFlags.slice(0, i).filter(Boolean).length})`,
            )
            .join(
              ', ',
            )} degraded to a plain-text fallback page after conversion failure (all layout and styling lost): immediately redo these pages in place with regenerate_slide following the original brief, then reply to the user.`
        : ''
      return {
        output: okMsg + failMsg + degradedMsg + imageFailNote(deckImageFails) + progressTail,
        mutated: true,
        summary: t('aiSumDeckGenerated', { done: landedPages, total }),
      }
    }

    case 'edit_chart': {
      const idx = Number(call.input.slideIndex)
      const sourceId = String(call.input.sourceId ?? '')
      if (!slides[idx])
        return fail(t('aiFailChartEdit'), `slideIndex out of range (0-${slides.length - 1})`)
      const op: import('../../shared/ipc').EditChartOp = { slideIndex: idx, sourceId }
      if (call.input.kind != null)
        op.kind = String(call.input.kind) as import('../../shared/ipc').EditChartOp['kind']
      if (Array.isArray(call.input.categories))
        op.categories = (call.input.categories as unknown[]).map(String)
      if (Array.isArray(call.input.series)) {
        const gateErr = dataSourceGateError(call, state)
        if (gateErr) return fail(t('aiFailChartEdit'), gateErr)
        op.series = (call.input.series as Array<{ name: unknown; values: unknown[] }>).map((s) => ({
          name: String(s.name ?? ''),
          values: (Array.isArray(s.values) ? s.values : []).map(Number),
        }))
      }
      if (call.input.colorScheme != null) op.colorScheme = String(call.input.colorScheme)
      if (call.input.title != null) op.title = String(call.input.title)
      if (call.input.legendPos != null)
        op.legendPos = String(
          call.input.legendPos,
        ) as import('../../shared/ipc').EditChartOp['legendPos']
      if (typeof call.input.dataLabels === 'boolean') op.dataLabels = call.input.dataLabels
      if (typeof call.input.gridlines === 'boolean') op.gridlines = call.input.gridlines
      if (call.input.switchRowCol === true) op.switchRowCol = true
      const updated = await window.slidesApi.editChart(op)
      if (!updated)
        return fail(
          t('aiFailChartEdit'),
          `Operation failed (element ${sourceId} does not exist or is not a chart)`,
        )
      access.applySlide(idx, updated.slide)
      const sampleNote = op.series && call.input.dataSource === 'sample' ? SAMPLE_DATA_NOTE : ''
      return {
        output: `Updated chart ${sourceId} on page ${idx + 1}.${sampleNote}`,
        mutated: true,
        summary: t('aiSumChartEdit', { n: idx + 1 }),
      }
    }

    case 'apply_ops': {
      const opsIn = Array.isArray(call.input.ops) ? (call.input.ops as unknown[]) : null
      if (!opsIn || opsIn.length === 0)
        return fail(t('aiFailApplyOps'), 'ops must be a non-empty array')
      const pre = preflightOps(opsIn, slides, state)
      if (!('ops' in pre)) return pre
      const r = await window.slidesApi.applyTxn?.({
        ops: pre.ops,
        ...(call.input.dry_run === true ? { dryRun: true } : {}),
        ...(call.input.isolation === 'per_op' ? { isolation: 'per_op' as const } : {}),
      })
      if (!r) return fail(t('aiFailApplyOps'), 'The transaction could not be executed')
      const failLines = (r.failures ?? []).map((f) => `- ops[${f.index}]: ${f.error}`).join('\n')
      if (r.dryRun) {
        const planStr = (r.plan ?? []).join('\n')
        return {
          output:
            `Dry run — the deck was NOT modified.\n` +
            (planStr ? `Valid plan:\n${planStr}\n` : '') +
            (failLines ? `Rejected ops:\n${failLines}\n` : 'All ops validated.\n') +
            `Resend with dry_run omitted to apply.`,
          mutated: false,
          summary: t('aiSumApplyOpsDryRun'),
        }
      }
      if (!r.applied) {
        return fail(
          t('aiFailApplyOps'),
          `Nothing was applied${call.input.isolation === 'per_op' ? '' : ' (atomic)'}:\n${failLines || 'unknown failure'}`,
        )
      }
      // Ops can delete slides — clamp the current index into the rebuilt deck
      // (page deletion may leave the current index past the end)
      access.applyDeck(r.slides!, Math.max(0, Math.min(access.getCurrent(), r.slides!.length - 1)))
      const created = (r.records ?? []).flatMap((rec) => rec.created ?? [])
      const doneStr = (r.records ?? [])
        .map((rec) => `${rec.op}${rec.target ? ` @${rec.target}` : ''}`)
        .join(', ')
      const notes = (r.records ?? [])
        .filter((rec) => rec.note)
        .map((rec) => `${rec.target ?? rec.op}: ${rec.note}`)
      return {
        output:
          `Applied ${r.records?.length ?? 0} op(s): ${doneStr}.` +
          (created.length ? ` New element ids: ${created.join(', ')}.` : '') +
          (notes.length ? `\nNote:\n${notes.join('\n')}` : '') +
          (failLines ? `\nSkipped (per_op):\n${failLines}` : '') +
          (pre.sampleData ? SAMPLE_DATA_NOTE : '') +
          auditTouchedPages(r.records, r.slides!),
        mutated: true,
        summary: applyOpsSummary(pre.ops, r.records?.length ?? 0),
      }
    }

    case 'load_guide': {
      const raw = Array.isArray(call.input.groups) ? call.input.groups.map(String) : []
      if (raw.length === 0) return fail(t('aiFailLoadGuide'), 'groups must be a non-empty array')
      const unknown = raw.filter((g) => !opGuide(g))
      if (unknown.length > 0) {
        return fail(
          t('aiFailLoadGuide'),
          `Unknown guide group(s): ${unknown.join(', ')}. Available: ${OP_GROUPS.join(', ')}.`,
        )
      }
      const groups = [...new Set(raw)]
      return {
        output: groups.map((g) => opGuide(g)!).join('\n\n---\n\n'),
        mutated: false,
        summary: t('aiSumLoadGuide', { names: groups.join(', ') }),
      }
    }

    case 'save_style_template': {
      const name = String(call.input.name ?? '').trim()
      if (!name) return fail(t('aiFailSaveTemplate'), 'name must not be empty')
      if (!access.saveStyleTemplate)
        return fail(
          t('aiFailSaveTemplate'),
          'The current environment does not support template saving',
        )
      const styleSkillToSave = state?.lastStyleSkill ?? ''
      const topicToSave = state?.lastTopic ?? ''
      if (!styleSkillToSave)
        return fail(
          t('aiFailSaveTemplate'),
          'The current deck has no Style Skill to save (generate a presentation with generate_deck first)',
        )
      // Chrome skeleton from the current deck (title box, brand-image slot, accents,
      // backgrounds) so pages generated from this template stay geometrically consistent
      const layout = extractLayoutSkeleton(access.getSlides()) ?? undefined
      const r = await access.saveStyleTemplate(name, {
        topic: topicToSave,
        styleSkill: styleSkillToSave,
        createdAt: new Date().toISOString(),
        ...(layout ? { layout } : {}),
      })
      if (!r.ok) return fail(t('aiFailSaveTemplate'), r.error ?? 'Save failed')
      return {
        output: `Saved the style "${name}" as a template${layout ? ' with its layout skeleton (title/brand-image geometry is pinned for future pages)' : ''}; next time pass style_template:"${name}" to reuse it directly.`,
        mutated: false,
        summary: t('aiSumSaveTemplate', { name }),
      }
    }

    case 'list_style_templates': {
      if (!access.listStyleTemplates)
        return fail(
          t('aiFailListTemplates'),
          'The current environment does not support template listing',
        )
      const templates = await access.listStyleTemplates()
      if (templates.length === 0) {
        return {
          output:
            'No saved style templates yet. After generating a deck, call save_style_template(name) to save the current style.',
          mutated: false,
          summary: t('aiSumTemplatesEmpty'),
        }
      }
      const lines = templates.map(
        (t) => `- ${t.name} (topic: ${t.topic || 'unknown'}, saved ${t.createdAt.slice(0, 10)})`,
      )
      return {
        output: `Saved style templates (${templates.length}):\n${lines.join('\n')}\n\nPass style_template:"<template name>" to generate_deck to reuse that style directly.`,
        mutated: false,
        summary: t('aiSumListTemplates', { count: templates.length }),
      }
    }

    default:
      return fail(call.name, `Unknown tool: ${call.name}`)
  }
}
