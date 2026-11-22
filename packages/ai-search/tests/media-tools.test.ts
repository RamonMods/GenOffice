import { describe, it, expect, vi, beforeEach } from 'vitest'

vi.mock('../src/ramonmods', () => ({
  ramonmodsGenerateImage: vi.fn(),
  ramonmodsAnalyzeMedia: vi.fn(),
  hasGskAuth: vi.fn(() => true),
}))

import { generateImageTool, analyzeMediaTool, RAMONMODS_RMBG_MODEL } from '../src/media-tools'
import { ramonmodsGenerateImage, ramonmodsAnalyzeMedia } from '../src/ramonmods'

const ramonmodsGen = vi.mocked(ramonmodsGenerateImage)
const ramonmodsAnalyze = vi.mocked(ramonmodsAnalyzeMedia)
// nonexistent settings file → defaults: no BYOK media provider, cloud tools on → RamonMods route
const SETTINGS = '/nonexistent/ai-settings.json'

beforeEach(() => {
  ramonmodsGen.mockReset()
  ramonmodsAnalyze.mockReset()
})

describe('generateImageTool transparentBackground (RamonMods route)', () => {
  it('chains a fal-bria-rmbg pass over the generated image and returns the stripped URL', async () => {
    ramonmodsGen
      .mockResolvedValueOnce({ url: 'https://cdn/x/opaque.png', taskId: '1' })
      .mockResolvedValueOnce({ url: 'https://cdn/x/cutout.png', taskId: '2' })
    const r = await generateImageTool(SETTINGS, {
      prompt: 'red podcast icon',
      transparentBackground: true,
    })
    expect(r).toEqual({ url: 'https://cdn/x/cutout.png' })
    expect(ramonmodsGen).toHaveBeenCalledTimes(2)
    expect(ramonmodsGen.mock.calls[1]![0]).toMatchObject({
      model: RAMONMODS_RMBG_MODEL,
      referenceImageUrls: ['https://cdn/x/opaque.png'],
    })
  })

  it('without the flag generation stays a single pass', async () => {
    ramonmodsGen.mockResolvedValueOnce({ url: 'https://cdn/x/opaque.png', taskId: '1' })
    const r = await generateImageTool(SETTINGS, { prompt: 'red podcast icon' })
    expect(r).toEqual({ url: 'https://cdn/x/opaque.png' })
    expect(ramonmodsGen).toHaveBeenCalledTimes(1)
  })

  it('does not chain when the caller already runs the background-removal model', async () => {
    ramonmodsGen.mockResolvedValueOnce({ url: 'https://cdn/x/cutout.png', taskId: '1' })
    const r = await generateImageTool(SETTINGS, {
      prompt: 'remove background',
      model: RAMONMODS_RMBG_MODEL,
      referenceImageUrls: ['https://cdn/x/src.png'],
      transparentBackground: true,
    })
    expect(r).toEqual({ url: 'https://cdn/x/cutout.png' })
    expect(ramonmodsGen).toHaveBeenCalledTimes(1)
  })

  it('falls back to the opaque image when the strip pass fails', async () => {
    ramonmodsGen
      .mockResolvedValueOnce({ url: 'https://cdn/x/opaque.png', taskId: '1' })
      .mockRejectedValueOnce(new Error('rmbg down'))
    const r = await generateImageTool(SETTINGS, {
      prompt: 'red podcast icon',
      transparentBackground: true,
    })
    expect(r).toEqual({ url: 'https://cdn/x/opaque.png' })
  })

  it('a generation failure still surfaces as an error (no half-run chain)', async () => {
    ramonmodsGen.mockRejectedValueOnce(new Error('quota exceeded'))
    const r = await generateImageTool(SETTINGS, {
      prompt: 'red podcast icon',
      transparentBackground: true,
    })
    expect(r).toEqual({ error: 'quota exceeded' })
    expect(ramonmodsGen).toHaveBeenCalledTimes(1)
  })
})

describe('generateImageTool reference budget (RamonMods route)', () => {
  it('rejects too many reference images with the BYOK wording, before calling the CLI', async () => {
    ramonmodsGen.mockResolvedValue({ url: 'https://cdn/x/out.png', taskId: '1' })
    const referenceImageUrls = Array.from({ length: 13 }, (_, i) => `https://cdn/x/${i}.png`)
    const r = await generateImageTool(SETTINGS, { prompt: 'red podcast icon', referenceImageUrls })
    expect(r).toEqual({
      error: 'Muitos itens de mídia em uma única solicitação (13, limite 12); analise em lotes menores',
    })
    expect(ramonmodsGen).not.toHaveBeenCalled()
  })

  it('still passes a reference count within the budget to the CLI', async () => {
    ramonmodsGen.mockResolvedValueOnce({ url: 'https://cdn/x/out.png', taskId: '1' })
    const referenceImageUrls = Array.from({ length: 12 }, (_, i) => `https://cdn/x/${i}.png`)
    const r = await generateImageTool(SETTINGS, { prompt: 'red podcast icon', referenceImageUrls })
    expect(r).toEqual({ url: 'https://cdn/x/out.png' })
    expect(ramonmodsGen).toHaveBeenCalledTimes(1)
  })
})

describe('analyzeMediaTool media budget (RamonMods route)', () => {
  it('rejects too many media URLs with the BYOK wording, before calling the CLI', async () => {
    ramonmodsAnalyze.mockResolvedValueOnce('done')
    const mediaUrls = Array.from({ length: 13 }, (_, i) => `https://cdn/x/${i}.png`)
    const r = await analyzeMediaTool(SETTINGS, { mediaUrls, requirements: 'describe these' })
    expect(r).toEqual({
      error: 'Muitos itens de mídia em uma única solicitação (13, limite 12); analise em lotes menores',
    })
    expect(ramonmodsAnalyze).not.toHaveBeenCalled()
  })
})
