import { readFileSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import { flagBool, flagString } from '../args'
import { readInput, resolveInput } from '../fs'
import { contactSheet, cropPng } from '../formats/png'
import { describeDeck, openDeck, type ElementSummary } from '../formats/pptx'
import {
  outputDirectory,
  parseScale,
  RENDERABLE,
  renderToPngs,
  type RenderedFile,
} from '../formats/render'
import type { CommandDef } from '../registry'
import { CliError, EXIT, type CommandResult } from '../result'
import { didYouMean } from '../suggest'

const DEFAULT_PAD_PX = 16
const DEFAULT_GRID_COLS = 4
const DEFAULT_TILE_PX = 320

export const renderCommand: CommandDef = {
  name: 'render',
  summary:
    'Um PNG por página de um documento, conforme o renderizador do GenOffice: a imagem que um agente usa para conferir um documento do Word, uma planilha ou uma página recém-criada.',
  usage:
    'render <file> --out <dir> [--page n] [--scale n] [--el e_12[,e_13] [--pad px]] [--grid [--cols n] [--tile px]]',
  options: [
    { name: 'out', value: 'dir', description: 'diretório para os PNGs (<stem>-NN.png; obrigatório)' },
    { name: 'page', value: 'n', description: 'apenas esta página (base 1; padrão: todas as páginas)' },
    {
      name: 'scale',
      value: 'n',
      description: 'pixels por ponto PDF, 0 < n <= 4 (padrão 1 = 72 dpi; 2 = 144 dpi)',
    },
    {
      name: 'el',
      value: 'ids',
      description:
        'pptx: também recorta a página nesses IDs de elemento (separados por vírgula, de `slides read`) como <stem>-NN-<id>.png',
    },
    {
      name: 'pad',
      value: 'px',
      description: `--el: margem ao redor do elemento (padrão ${DEFAULT_PAD_PX})`,
    },
    {
      name: 'grid',
      description:
        'também grava <stem>-grid.png: todas as páginas renderizadas reduzidas em uma única folha de contato',
    },
    {
      name: 'cols',
      value: 'n',
      description: `--grid: blocos por linha (padrão ${DEFAULT_GRID_COLS})`,
    },
    { name: 'tile', value: 'px', description: `--grid: largura do bloco (padrão ${DEFAULT_TILE_PX})` },
  ],
  async run(args, ctx) {
    const path = resolveInput(args.positionals[0], ctx)
    const outDir = outputDirectory(flagString(args, 'out'), ctx)
    let only = pageIndex(flagString(args, 'page'))
    const elIds = flagString(args, 'el')
      ?.split(',')
      .map((s) => s.trim())
      .filter(Boolean)
    let targets: { id: string; slide: number; el: ElementSummary }[] = []
    let deckSize: { cx: number; cy: number } | undefined
    if (elIds?.length) {
      if (extname(path).toLowerCase() !== '.pptx') {
        throw new CliError(EXIT.usage, '--el funciona apenas em arquivos .pptx', undefined, {
          reason: 'unsupported',
        })
      }
      const deck = describeDeck(await openDeck(readInput(path)))
      deckSize = deck.size
      targets = elIds.map((id) => locate(deck.pages, id))
      const slides = new Set(targets.map((t) => t.slide))
      if (only !== undefined && (slides.size > 1 || !slides.has(only))) {
        throw new CliError(
          EXIT.usage,
          `--page ${only + 1} não contém ${elIds.join(', ')}`,
          undefined,
          {
            reason: 'invalid_argument',
            suggestion: `os elemento(s) estão na(s) página(s) ${[...slides].map((s) => s + 1).join(', ')}; remova --page ou ajuste`,
          },
        )
      }
      if (slides.size > 1) {
        throw new CliError(EXIT.usage, '--el: os IDs devem estar na mesma lâmina', undefined, {
          reason: 'invalid_argument',
          suggestion: `estão nas páginas ${[...slides].map((s) => s + 1).join(', ')}; execute uma vez por página`,
        })
      }
      only = [...slides][0]
    }
    const files = await renderToPngs(path, ctx, {
      outDir,
      scale: parseScale(flagString(args, 'scale')),
      only,
      range: { flag: 'page', oneBased: true },
      log: ctx.log,
    })
    const stem = basename(path, extname(path))
    const crops: (RenderedFile & { el: string })[] = []
    if (targets.length) {
      const page = files[0]!
      const png = readFileSync(page.path)
      const pad = whole(flagString(args, 'pad'), DEFAULT_PAD_PX, 'pad', 0)
      const pxPerEmu = page.width / deckSize!.cx
      for (const t of targets) {
        const { png: cropped, rect } = cropPng(
          png,
          {
            x: t.el.box.x * pxPerEmu,
            y: t.el.box.y * pxPerEmu,
            w: t.el.box.cx * pxPerEmu,
            h: t.el.box.cy * pxPerEmu,
          },
          pad,
        )
        const out = join(outDir, `${stem}-${String(page.page + 1).padStart(2, '0')}-${t.id}.png`)
        writeFileSync(out, cropped)
        crops.push({ page: page.page, path: out, width: rect.w, height: rect.h, el: t.id })
      }
    }
    const detail: Record<string, unknown> = {
      files: [...files, ...crops].map((f) => ({ ...f, page: f.page + 1 })),
      formats: RENDERABLE,
      via: path.toLowerCase().endsWith('.pdf') ? 'pdfium' : 'genoffice --headless-export + pdfium',
    }
    if (flagBool(args, 'grid')) {
      const sheet = contactSheet(
        files.map((f) => ({ page: f.page + 1, png: readFileSync(f.path) })),
        positive(flagString(args, 'cols'), DEFAULT_GRID_COLS, 'cols'),
        positive(flagString(args, 'tile'), DEFAULT_TILE_PX, 'tile'),
      )
      const out = join(outDir, `${stem}-grid.png`)
      writeFileSync(out, sheet.png)
      detail.grid = {
        path: out,
        width: sheet.width,
        height: sheet.height,
        cols: sheet.cols,
        rows: sheet.rows,
        tiles: sheet.tiles,
      }
    }
    const extras = [
      crops.length ? `${crops.length} recorte(s)` : '',
      flagBool(args, 'grid') ? 'uma folha de contato' : '',
    ].filter(Boolean)
    const result: CommandResult = {
      summary: `renderizadas ${files.length} página(s) de ${basename(path)} em ${outDir}${extras.length ? ` mais ${extras.join(' e ')}` : ''}`,
      outputPath: outDir,
      detail,
    }
    return result
  },
}

function pageIndex(raw: string | undefined): number | undefined {
  if (raw === undefined) return undefined
  const n = Number(raw)
  if (!Number.isInteger(n) || n < 1) {
    throw new CliError(EXIT.usage, '--page deve ser 1 ou maior', undefined, {
      reason: 'invalid_argument',
    })
  }
  return n - 1
}

function positive(raw: string | undefined, fallback: number, flag: string): number {
  return whole(raw, fallback, flag, 1)
}

function whole(raw: string | undefined, fallback: number, flag: string, min: number): number {
  if (raw === undefined) return fallback
  const n = Number(raw)
  if (!Number.isInteger(n) || n < min) {
    throw new CliError(
      EXIT.usage,
      `--${flag} deve ser um número inteiro maior ou igual a ${min}`,
      undefined,
      {
        reason: 'invalid_argument',
      },
    )
  }
  return n
}

function locate(
  pages: { index: number; elements: ElementSummary[] }[],
  id: string,
): { id: string; slide: number; el: ElementSummary } {
  // Only top-level boxes are in slide coordinates; a group child's box is in its group's space.
  const all: { slide: number; el: ElementSummary }[] = []
  const inGroup = new Map<string, string>()
  const walk = (parent: string, els: ElementSummary[]) => {
    for (const el of els) {
      if (el.id) inGroup.set(el.id, parent)
      if (el.children) walk(el.id ?? parent, el.children)
    }
  }
  for (const p of pages) {
    for (const el of p.elements) {
      all.push({ slide: p.index, el })
      if (el.children) walk(el.id ?? '?', el.children)
    }
  }
  const hit = all.find((e) => e.el.id === id)
  if (!hit) {
    const group = inGroup.get(id)
    if (group) {
      throw new CliError(EXIT.usage, `${id} é um elemento filho do grupo ${group}`, undefined, {
        reason: 'invalid_argument',
        suggestion: `recorte o grupo: --el ${group}`,
      })
    }
    const ids = all.map((e) => e.el.id).filter((x): x is string => !!x)
    const guess = didYouMean(id, ids)
    throw new CliError(
      EXIT.usage,
      `nenhum elemento ${id} na apresentação`,
      { available: ids.slice(0, 50) },
      {
        reason: 'target_not_found',
        suggestion: guess
          ? `você quis dizer ${guess}?`
          : 'execute `genoffice slides read <file>` para ver os IDs',
      },
    )
  }
  return { id, ...hit }
}
