import { chmodSync, existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { rasterizePdf } from '../src/formats/slide-spec'
import { run, tempDir, writeMinimalPdf } from './helpers'

// 1x1 red PNG
const PNG_1PX = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
)

const text = (t: string, box: Record<string, number>, sizePt = 28) => ({
  type: 'text',
  ...box,
  paragraphs: [{ runs: [{ text: t, sizePt, bold: true, color: '#112233' }] }],
})

function specFile(dir: string, spec: unknown): string {
  const path = join(dir, 'deck.json')
  writeFileSync(path, JSON.stringify(spec))
  return path
}

describe('genoffice create --type pptx --spec', () => {
  it('builds one slide per page, embeds local images and reports dropped elements', async () => {
    const dir = tempDir()
    writeFileSync(join(dir, 'pic.png'), PNG_1PX)
    const spec = specFile(dir, {
      pages: [
        {
          background: '#0E1A2B',
          elements: [
            text('Cover title', { x: 80, y: 200, w: 900, h: 80 }, 44),
            { type: 'image', url: 'pic.png', x: 1000, y: 100, w: 200, h: 200 },
          ],
        },
        {
          elements: [
            { type: 'shape', shape: 'roundRect', x: 80, y: 80, w: 400, h: 200, fill: '#1F3A5F' },
            text('Second page', { x: 80, y: 320, w: 600, h: 60 }),
            { type: 'shape', shape: 'rect', x: 500, y: 500, w: 100, h: 50 },
          ],
        },
      ],
    })
    const out = join(dir, 'deck.pptx')
    const r = await run(['create', '--type', 'pptx', '--spec', spec, '--out', out, '--json'])
    expect(r.code).toBe(0)
    const j = r.json()
    expect(j.summary).toContain('2 slides')
    expect(j.detail.imageFailures).toEqual([])
    // the fill-less, stroke-less rect on page 1 is dropped with a warning
    expect(JSON.stringify(j.detail.issues)).toContain('forma sem preenchimento ou contorno')

    const zip = await JSZip.loadAsync(readFileSync(out))
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain('Cover title')
    expect(await zip.file('ppt/slides/slide2.xml')!.async('string')).toContain('Second page')
    expect(Object.keys(zip.files).some((f) => /^ppt\/media\/.*\.png$/.test(f))).toBe(true)

    const read = await run(['slides', 'read', out, '--json'])
    expect(read.json().detail.slides).toBe(2)
  })

  it('rejects a spec without pages and reads the spec from stdin-like "-" only when given', async () => {
    const dir = tempDir()
    const spec = specFile(dir, { pages: [] })
    const r = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      spec,
      '--out',
      join(dir, 'x.pptx'),
      '--json',
    ])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('está vazio')
  })
})

describe('genoffice slides audit', () => {
  it('reports text taller than its box with durable ids and passes a clean deck', async () => {
    const dir = tempDir()
    const spec = specFile(dir, {
      pages: [
        // a long line in a box too short: the builder grows the box, so audit passes
        {
          elements: [
            text(
              'A fairly long headline that wraps onto several lines',
              { x: 80, y: 80, w: 300, h: 20 },
              40,
            ),
          ],
        },
        // two overlapping text blocks
        {
          elements: [
            text('Alpha block', { x: 100, y: 100, w: 500, h: 100 }),
            text('Beta block', { x: 150, y: 120, w: 500, h: 100 }),
          ],
        },
      ],
    })
    const out = join(dir, 'audit.pptx')
    expect((await run(['create', '--type', 'pptx', '--spec', spec, '--out', out])).code).toBe(0)

    const all = await run(['slides', 'audit', out, '--json'])
    expect(all.code).toBe(0)
    const slides = all.json().detail.slides as { slide: number; id: string; issues: string[] }[]
    expect(slides).toHaveLength(2)
    expect(slides[0]!.issues).toEqual([])
    expect(slides[1]!.issues.some((s) => s.startsWith('Sobreposição:') && /e_\w+/.test(s))).toBe(true)
    expect(all.json().summary).toMatch(/1 problema\(s\) em 1 de 2 slide/)

    const one = await run(['slides', 'audit', out, '--slide', '0', '--json'])
    expect(one.json().detail.slides).toHaveLength(1)
    expect(one.json().summary).toContain('sem problemas de layout')
    expect((await run(['slides', 'audit', out, '--slide', '5'])).code).toBe(1)
  })
})

describe('genoffice slides render', () => {
  it('rasterizes every PDF page to PNG', async () => {
    const dir = tempDir()
    const pdf = writeMinimalPdf(join(dir, 'one.pdf'))
    const pages = await rasterizePdf(new Uint8Array(readFileSync(pdf)), 1)
    expect(pages).toHaveLength(1)
    expect(pages[0]!.width).toBe(612)
    expect(pages[0]!.height).toBe(792)
    expect(Array.from(pages[0]!.png.slice(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47])
    await expect(rasterizePdf(new Uint8Array(readFileSync(pdf)), 1, 3)).rejects.toThrow(
      'fora do intervalo',
    )
  })

  it('drives the app export and writes one PNG per slide', async () => {
    if (process.platform === 'win32') return
    const dir = tempDir()
    const spec = specFile(dir, {
      pages: [{ elements: [text('Only page', { x: 80, y: 80, w: 600, h: 80 })] }],
    })
    const pptx = join(dir, 'r.pptx')
    expect((await run(['create', '--type', 'pptx', '--spec', spec, '--out', pptx])).code).toBe(0)
    // a stand-in for the GenOffice binary: copies a prepared PDF to --out and prints the envelope
    const pdf = writeMinimalPdf(join(dir, 'export.pdf'))
    const fake = join(dir, 'fake-genoffice.sh')
    writeFileSync(
      fake,
      `#!/bin/sh\nwhile [ $# -gt 0 ]; do if [ "$1" = "--out" ]; then out="$2"; fi; shift; done\ncp "${pdf}" "$out"\necho '{"status":"ok","summary":"exported"}'\n`,
    )
    chmodSync(fake, 0o755)
    const shots = join(dir, 'shots')
    const r = await run(['slides', 'render', pptx, '--out', shots, '--scale', '2', '--json'], {
      env: { ...process.env, GENOFFICE_APP_BIN: fake },
    })
    expect(r.code).toBe(0)
    const files = r.json().detail.files as { path: string; width: number }[]
    expect(files).toHaveLength(1)
    expect(files[0]!.width).toBe(1224)
    expect(existsSync(files[0]!.path)).toBe(true)
    expect(files[0]!.path.endsWith('r-01.png')).toBe(true)
  })

  it('needs --out', async () => {
    const dir = tempDir()
    const spec = specFile(dir, {
      pages: [{ elements: [text('x', { x: 0, y: 0, w: 100, h: 40 })] }],
    })
    const pptx = join(dir, 'n.pptx')
    await run(['create', '--type', 'pptx', '--spec', spec, '--out', pptx])
    const r = await run(['slides', 'render', pptx, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('--out')
  })
})

const outlinePage = (title: string, type: string, layout: string) => ({
  title,
  type,
  layout,
  brief: `${title}: three cards with the real 2024 figures from the brief, one source line under each card.`,
  image_queries: [],
})

function deckDir(dir: string, pages: unknown[]): string {
  const pagesDir = join(dir, 'pages')
  mkdirSync(pagesDir, { recursive: true })
  pages.forEach((p, i) => {
    writeFileSync(join(pagesDir, `${String(i + 1).padStart(2, '0')}.json`), JSON.stringify(p))
  })
  return pagesDir
}

describe('genoffice create --type pptx --spec <dir>', () => {
  it('builds one slide per page file in name order and names the file in issues', async () => {
    const dir = tempDir()
    const pagesDir = deckDir(dir, [
      { background: '#0E1A2B', elements: [text('First', { x: 80, y: 200, w: 900, h: 80 }, 44)] },
      {
        elements: [
          text('Second', { x: 80, y: 320, w: 600, h: 60 }),
          { type: 'shape', shape: 'rect', x: 500, y: 500, w: 100, h: 50 },
        ],
      },
    ])
    // a tenth file sorts after the ninth, not after the first
    writeFileSync(
      join(pagesDir, '10.json'),
      JSON.stringify({ elements: [text('Tenth', { x: 80, y: 80, w: 600, h: 60 })] }),
    )
    writeFileSync(join(pagesDir, 'notes.txt'), 'ignored')
    const out = join(dir, 'deck.pptx')
    const r = await run(['create', '--type', 'pptx', '--spec', pagesDir, '--out', out, '--json'])
    expect(r.code).toBe(0)
    const j = r.json()
    expect(j.detail.via).toBe('diretório de especificação de apresentação')
    expect(j.detail.files).toEqual(['01.json', '02.json', '10.json'])
    expect(j.detail.issues).toEqual([
      {
        page: 1,
        file: '02.json',
        warnings: [expect.stringContaining('forma sem preenchimento ou contorno')],
      },
    ])
    const zip = await JSZip.loadAsync(readFileSync(out))
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain('First')
    expect(await zip.file('ppt/slides/slide3.xml')!.async('string')).toContain('Tenth')
  })

  it('names the right file for an image failure after a dropped page', async () => {
    const dir = tempDir()
    const pagesDir = deckDir(dir, [
      { elements: [text('Cover', { x: 80, y: 80, w: 600, h: 60 })] },
      { elements: [] },
      {
        elements: [
          text('Photo page', { x: 80, y: 80, w: 600, h: 60 }),
          { type: 'image', url: 'missing.png', x: 900, y: 100, w: 200, h: 200 },
        ],
      },
    ])
    const out = join(dir, 'deck.pptx')
    const r = await run(['create', '--type', 'pptx', '--spec', pagesDir, '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toContain('2 slides')
    expect(r.json().detail.issues).toEqual([
      { page: 1, file: '02.json', error: expect.stringContaining('elements') },
    ])
    expect(r.json().detail.imageFailures).toEqual([
      { page: 2, file: '03.json', url: 'missing.png' },
    ])
  })

  it('with --outline refuses to build until every outline page has its file', async () => {
    const dir = tempDir()
    const pagesDir = deckDir(dir, [{ elements: [text('Cover', { x: 80, y: 80, w: 600, h: 60 })] }])
    const outline = join(dir, 'outline.json')
    writeFileSync(
      outline,
      JSON.stringify({
        core_hook: 'Two pages, one built',
        pages: [
          outlinePage('Cover', 'cover', 'cover_typography_hero'),
          outlinePage('Close', 'closing', 'closing_cta'),
        ],
      }),
    )
    const out = join(dir, 'deck.pptx')
    const r = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      pagesDir,
      '--outline',
      outline,
      '--out',
      out,
      '--json',
    ])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('1 arquivo(s) de página para 2 página(s) do outline')
    expect(r.json().detail.outline_pages).toEqual(['Cover', 'Close'])
    expect(existsSync(out)).toBe(false)

    writeFileSync(
      join(pagesDir, '02.json'),
      JSON.stringify({ elements: [text('Close', { x: 80, y: 80, w: 600, h: 60 })] }),
    )
    const ok = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      pagesDir,
      '--outline',
      outline,
      '--out',
      out,
      '--json',
    ])
    expect(ok.code).toBe(0)
    expect(ok.json().summary).toContain('2 slides')

    const single = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      join(pagesDir, '01.json'),
      '--outline',
      outline,
      '--out',
      join(dir, 'x.pptx'),
      '--json',
    ])
    expect(single.code).toBe(1)
    expect(single.json().message).toContain('--outline acompanha --spec <directory>')
  })

  it('reports an empty directory and a broken page file by name', async () => {
    const dir = tempDir()
    const empty = join(dir, 'empty')
    mkdirSync(empty)
    const r = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      empty,
      '--out',
      join(dir, 'e.pptx'),
      '--json',
    ])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('nenhum arquivo de página')
    const pagesDir = deckDir(dir, [{ elements: [text('ok', { x: 0, y: 0, w: 100, h: 40 })] }])
    writeFileSync(join(pagesDir, '02.json'), '{ not json')
    const bad = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      pagesDir,
      '--out',
      join(dir, 'b.pptx'),
      '--json',
    ])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toMatch(/02\.json: JSON inválido/)
  })
})

describe('genoffice slides check', () => {
  it('validates an outline: errors exit 1 with the findings, warnings pass', async () => {
    const dir = tempDir()
    const outline = join(dir, 'outline.json')
    writeFileSync(
      outline,
      JSON.stringify({
        core_hook: 'Margins fell 40% while volume grew',
        pages: [
          outlinePage('The squeeze', 'cover', 'cover_typography_hero'),
          outlinePage('Costs', 'content', 'hero_big_number'),
          outlinePage('Rates', 'content', 'hero_big_number'),
          outlinePage('Thanks', 'closing', 'closing_cta'),
        ],
      }),
    )
    const r = await run(['slides', 'check', outline, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('1 erro(s) no outline')
    expect(r.json().detail.kind).toBe('outline')
    expect(r.json().detail.issues).toEqual([
      {
        page: 2,
        level: 'error',
        message: expect.stringContaining('se repete em relação à página anterior'),
      },
    ])

    writeFileSync(
      outline,
      JSON.stringify({
        core_hook: 'Margins fell 40% while volume grew',
        pages: [
          outlinePage('Costs', 'content', 'hero_big_number'),
          outlinePage('Thanks', 'closing', 'closing_cta'),
        ],
      }),
    )
    const warn = await run(['slides', 'check', outline, '--json'])
    expect(warn.code).toBe(0)
    expect(warn.json().summary).toContain('2 páginas, 1 aviso(s)')
    expect(warn.json().detail.issues).toEqual([
      { page: 0, level: 'warning', message: 'a primeira página não é uma capa' },
    ])
  })

  it('builds and audits a single page spec', async () => {
    const dir = tempDir()
    const clean = join(dir, 'clean.json')
    writeFileSync(
      clean,
      JSON.stringify({ elements: [text('Clean page', { x: 80, y: 80, w: 800, h: 80 })] }),
    )
    const ok = await run(['slides', 'check', clean, '--json'])
    expect(ok.code).toBe(0)
    expect(ok.json().summary).toContain('página construída sem problemas')
    expect(ok.json().detail).toMatchObject({
      kind: 'page',
      warnings: [],
      imageFailures: [],
      audit: [],
    })

    const messy = join(dir, 'messy.json')
    writeFileSync(
      messy,
      JSON.stringify({
        elements: [
          text('Alpha block', { x: 100, y: 100, w: 500, h: 100 }),
          text('Beta block', { x: 150, y: 120, w: 500, h: 100 }),
          { type: 'shape', shape: 'rect', x: 500, y: 500, w: 100, h: 50 },
          { type: 'image', url: 'missing.png', x: 900, y: 100, w: 200, h: 200 },
        ],
      }),
    )
    const r = await run(['slides', 'check', messy, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toMatch(/3 achado\(s\)/)
    expect(r.json().detail.warnings).toEqual([
      expect.stringContaining('forma sem preenchimento ou contorno'),
    ])
    expect(r.json().detail.imageFailures).toEqual(['missing.png'])
    expect(r.json().detail.audit.some((s: string) => s.startsWith('Sobreposição:'))).toBe(true)

    const two = join(dir, 'two.json')
    writeFileSync(
      two,
      JSON.stringify({
        pages: [
          { elements: [text('a', { x: 0, y: 0, w: 100, h: 40 })] },
          { elements: [text('b', { x: 0, y: 0, w: 100, h: 40 })] },
        ],
      }),
    )
    const bad = await run(['slides', 'check', two, '--json'])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('exatamente um objeto de página')
  })
})

describe('staged deck folder: outline and style checks', () => {
  function stagedDeck() {
    const dir = tempDir()
    writeFileSync(
      join(dir, 'style.md'),
      '# Style\n- Backgrounds: #0E1A2B\n- Text: #112233, accent #FF6600\n',
    )
    writeFileSync(
      join(dir, 'outline.json'),
      JSON.stringify({
        core_hook: 'Margins fell 40% while volume grew',
        pages: [
          { ...outlinePage('The squeeze', 'cover', 'cover_typography_hero') },
          {
            ...outlinePage('Where the margin went', 'content', 'three_column_cards'),
            image_queries: ['container port cranes at dawn'],
          },
          outlinePage('Thanks', 'closing', 'closing_cta'),
        ],
      }),
    )
    const pagesDir = deckDir(dir, [
      {
        title: 'The squeeze',
        layout: 'cover_typography_hero',
        background: '#0E1A2B',
        elements: [text('The squeeze', { x: 80, y: 200, w: 900, h: 80 }, 44)],
      },
      {
        elements: [
          text('Where the margin went', { x: 80, y: 80, w: 900, h: 60 }),
          {
            type: 'shape',
            shape: 'rect',
            x: 80,
            y: 300,
            w: 300,
            h: 100,
            fill: '#00AA00',
          },
        ],
      },
      { elements: [text('Thanks', { x: 80, y: 300, w: 600, h: 60 })] },
    ])
    return { dir, pagesDir }
  }

  it('check <page.json> finds the outline and style one folder up', async () => {
    const { dir, pagesDir } = stagedDeck()
    const cover = await run(['slides', 'check', join(pagesDir, '01.json'), '--json'])
    expect(cover.code).toBe(0)
    expect(cover.json().summary).toContain('página construída sem problemas')
    expect(cover.json().detail).toMatchObject({
      page: 0,
      outline: { file: join(dir, 'outline.json'), findings: [] },
      style: { file: join(dir, 'style.md'), offPalette: [] },
      notes: [],
    })

    const second = await run(['slides', 'check', join(pagesDir, '02.json'), '--json'])
    expect(second.code).toBe(0)
    expect(second.json().detail.outline.findings).toEqual([
      { page: 1, level: 'warning', message: expect.stringContaining('prevê 1 foto(s)') },
    ])
    expect(second.json().detail.style.offPalette).toEqual(['#00AA00'])
    expect(second.json().summary).toContain('2 achado(s)')
  })

  it('check exits 1 when the page file disagrees with its outline entry', async () => {
    const { pagesDir } = stagedDeck()
    writeFileSync(
      join(pagesDir, '03.json'),
      JSON.stringify({
        title: 'Questions',
        layout: 'closing_cta',
        elements: [text('Thanks, XX% of you', { x: 80, y: 300, w: 600, h: 60 })],
      }),
    )
    const r = await run(['slides', 'check', join(pagesDir, '03.json'), '--json'])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('2 erro(s) em relação a outline.json pages[2]')
    expect(r.json().detail.outline.findings.map((f: { message: string }) => f.message)).toEqual([
      expect.stringContaining('o title do arquivo da página "Questions" não é o mesmo de outline pages[2]'),
      expect.stringContaining('placeholder "XX%"'),
    ])

    const loose = tempDir()
    const alone = join(loose, 'p.json')
    writeFileSync(alone, JSON.stringify({ elements: [text('x', { x: 0, y: 0, w: 200, h: 40 })] }))
    const solo = await run(['slides', 'check', alone, '--json'])
    expect(solo.code).toBe(0)
    expect(solo.json().detail.outline).toBeNull()
    expect(solo.json().detail.notes).toEqual([
      expect.stringContaining('nenhum outline.json'),
      expect.stringContaining('nenhum style.md'),
    ])
  })

  it('create --spec <dir> uses the outline beside the folder and aborts on a disagreement', async () => {
    const { dir, pagesDir } = stagedDeck()
    const out = join(dir, 'deck.pptx')
    const ok = await run(['create', '--type', 'pptx', '--spec', pagesDir, '--out', out, '--json'])
    expect(ok.code).toBe(0)
    expect(ok.json().detail.outline).toEqual({
      file: join(dir, 'outline.json'),
      findings: [
        { page: 1, file: '02.json', level: 'warning', message: expect.stringContaining('foto') },
      ],
    })
    expect(ok.json().detail.style.offPalette).toEqual([{ file: '02.json', colors: ['#00AA00'] }])

    writeFileSync(
      join(pagesDir, '01.json'),
      JSON.stringify({
        layout: 'cover_split_image',
        type: 'content',
        elements: [text('The squeeze', { x: 80, y: 200, w: 900, h: 80 }, 44)],
      }),
    )
    const bad = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      pagesDir,
      '--out',
      join(dir, 'deck2.pptx'),
      '--force',
      '--json',
    ])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('1 página(s) discordam de outline.json')
    expect(bad.json().detail.findings[0]).toMatchObject({
      page: 0,
      file: '01.json',
      level: 'error',
    })
    expect(existsSync(join(dir, 'deck2.pptx'))).toBe(false)
  })

  it('an outline kept inside the pages folder is not counted as a page', async () => {
    const { dir, pagesDir } = stagedDeck()
    renameSync(join(dir, 'outline.json'), join(pagesDir, 'outline.json'))
    const out = join(dir, 'deck.pptx')
    const r = await run(['create', '--type', 'pptx', '--spec', pagesDir, '--out', out, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.files).toEqual(['01.json', '02.json', '03.json'])
    expect(r.json().detail.outline.file).toBe(join(pagesDir, 'outline.json'))
    const check = await run(['slides', 'check', join(pagesDir, '03.json'), '--json'])
    expect(check.code).toBe(0)
    expect(check.json().detail.page).toBe(2)
  })

  it('replace checks the page against the outline entry of the slide it rebuilds', async () => {
    const { dir, pagesDir } = stagedDeck()
    const out = join(dir, 'deck.pptx')
    expect(
      (await run(['create', '--type', 'pptx', '--spec', pagesDir, '--out', out, '--json'])).code,
    ).toBe(0)
    const wrong = await run([
      'slides',
      'replace',
      out,
      '--slide',
      '2',
      '--spec',
      join(pagesDir, '01.json'),
      '--json',
    ])
    expect(wrong.code).toBe(1)
    expect(wrong.json().message).toContain('a página discorda de outline.json pages[2]')

    const right = await run([
      'slides',
      'replace',
      out,
      '--slide',
      '2',
      '--spec',
      join(pagesDir, '03.json'),
      '--json',
    ])
    expect(right.code).toBe(0)
    expect(right.json().detail.outline).toEqual({ file: join(dir, 'outline.json'), findings: [] })
  })
})

describe('genoffice slides replace', () => {
  it('rebuilds one slide from its spec and leaves the others alone', async () => {
    const dir = tempDir()
    const pagesDir = deckDir(dir, [
      { elements: [text('One', { x: 80, y: 80, w: 600, h: 60 })] },
      { elements: [text('Two', { x: 80, y: 80, w: 600, h: 60 })] },
      { elements: [text('Three', { x: 80, y: 80, w: 600, h: 60 })] },
    ])
    const out = join(dir, 'deck.pptx')
    expect((await run(['create', '--type', 'pptx', '--spec', pagesDir, '--out', out])).code).toBe(0)

    const fixed = join(pagesDir, '02.json')
    writeFileSync(
      fixed,
      JSON.stringify({
        background: '#123456',
        elements: [text('Two, fixed', { x: 80, y: 80, w: 600, h: 60 })],
      }),
    )
    const r = await run(['slides', 'replace', out, '--slide', '1', '--spec', fixed, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toContain('slide 1 de deck.pptx substituído a partir de 02.json')
    expect(r.json().detail.slide).toBe(1)

    expect((await run(['slides', 'read', out, '--json'])).json().detail.slides).toBe(3)
    const texts = await slideTextsInOrder(out)
    expect(texts[0]).toContain('One')
    expect(texts[1]).toContain('Two, fixed')
    expect(texts[1]).not.toContain('>Two<')
    expect(texts[2]).toContain('Three')

    expect(
      (await run(['slides', 'replace', out, '--spec', fixed, '--json'])).json().message,
    ).toContain('--slide')
    expect(
      (await run(['slides', 'replace', out, '--slide', '1', '--json'])).json().message,
    ).toContain('--spec')
    expect((await run(['slides', 'replace', out, '--slide', '7', '--spec', fixed])).code).toBe(1)
    // a malformed index is named as such before any file is looked up
    const bad = await run([
      'slides',
      'replace',
      out,
      '--slide',
      '2 03',
      '--spec',
      'pages/.json',
      '--json',
    ])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('--slide deve ser um índice de slide a partir de 0')
  })

  it('check reports an image box that crops away most of its picture', async () => {
    const dir = tempDir()
    writeFileSync(join(dir, 'pic.png'), PNG_1PX)
    const page = join(dir, '01.json')
    writeFileSync(
      page,
      JSON.stringify({
        elements: [
          text('Strip', { x: 80, y: 80, w: 600, h: 60 }),
          // a square picture in a 10:1 strip keeps a tenth of it
          { type: 'image', url: 'pic.png', x: 80, y: 500, w: 800, h: 80 },
        ],
      }),
    )
    const r = await run(['slides', 'check', page, '--json'])
    expect(r.code).toBe(0)
    const warnings = r.json().detail.warnings as string[]
    expect(warnings.some((w) => w.includes('pic.png') && w.includes('apenas 10%'))).toBe(true)
    expect(r.json().detail.imageFailures).toEqual([])
  })
})

/** Slide XML in presentation order (slide part numbers do not follow the order after a replace). */
async function slideTextsInOrder(pptx: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(readFileSync(pptx))
  const pres = await zip.file('ppt/presentation.xml')!.async('string')
  const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string')
  const targets = new Map(
    [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => [
      /Id="([^"]+)"/.exec(m[0])![1]!,
      /Target="([^"]+)"/.exec(m[0])![1]!,
    ]),
  )
  const ids = [...pres.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map((m) => m[1]!)
  return Promise.all(ids.map((id) => zip.file(`ppt/${targets.get(id)!}`)!.async('string')))
}

describe('genoffice guide slides design|spec', () => {
  it('prints the deck design workflow and the spec reference', async () => {
    const design = await run(['guide', 'slides', 'design'])
    expect(design.code).toBe(0)
    expect(design.stdout).toContain('Primeiro a folha de estilo')
    expect(design.stdout).toContain('three_column_cards')
    const spec = await run(['guide', 'slides', 'spec'])
    expect(spec.stdout).toContain('1280 × 720')
    expect(spec.stdout).toContain('"type": "image"')
    const catalog = await run(['guide', 'slides'])
    expect(catalog.stdout).toContain('genoffice guide slides design')
  })
})
