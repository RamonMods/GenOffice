/**
 * Stage 3b: the Word 97-2003 stylesheet (STSH) in the table stream.
 *
 * Layout (MS-DOC 2.9.278 STSH / 2.9.277 STD / POI StyleSheet+StyleDescription):
 *   cbStshi u16, STSHI (cbStshi bytes: cstd, cbSTDBaseInFile, ...),
 *   then cstd STD entries: cbStd u16 + cbStd bytes, each padded to a 2-byte
 *   boundary. STD = StdfBase (10 bytes: info1 sti, info2 stk+istdBase,
 *   info3 cupx+istdNext, bchUpe, grfstd), [StdfPost2000 when base = 0x12],
 *   XstzName (cch u16 + UTF-16LE name + null u16), then cupx UPX entries
 *   (cbUPX u16 + bytes, padded). A paragraph style's UPX[0] is a PAPX grpprl
 *   with a 2-byte istd prefix to skip; UPX[1] is its CHPX grpprl. A character
 *   style's UPX[0] is its CHPX grpprl.
 */
import { LegacyDocError } from './errors'
import type { DefaultFontIndices, FontTable } from './fonts'

export type StyleKind = 'paragraph' | 'character' | 'table' | 'numbering'

export interface Style {
  /** style index (position in the STSH) */
  istd: number
  /** built-in style id (sti); 0x0FFE/4094 for user-defined styles */
  sti: number
  kind: StyleKind
  /** istd of the style this one is based on, when any */
  basedOn?: number
  /** istd the next paragraph uses after this style (paragraph styles) */
  next?: number
  name: string
  /** PAPX sprms only (the 2-byte istd prefix is removed) */
  papxSprms: Uint8Array
  /** CHPX sprms; empty for a paragraph style without its own character properties */
  chpxSprms: Uint8Array
}

export interface StyleSheet {
  /** every STSH slot, including empty cbStd == 0 placeholders, keyed by istd */
  styles: Map<number, Style>
  /** STSHI rgftcStandardChpStsh: the document's default font indices */
  defaultFonts: DefaultFontIndices
  /** the font table, attached by the readers that have the FIB at hand */
  fonts?: FontTable
}

/** The w:styleId the .docx writer emits for a stylesheet slot. A pure function
    of the istd, so the resolver can stamp references into the formatting model
    and the writer can emit the matching definitions without threading a map. */
export const legacyStyleId = (istd: number): string => `w97-${istd}`

const STK_PARAGRAPH = 1
const STK_CHARACTER = 2
const STK_TABLE = 3
const STK_NUMBERING = 4
const ISTD_BASE_NONE = 0x0fff
const MIN_STSHI_BYTES = 18
const WORD_97_STD_BASE = 0x000a
const WORD_2000_STD_BASE = 0x0012
const MAX_STYLES = 0x1000

const EMPTY = new Uint8Array(0)

function badStsh(detail: string): never {
  throw new LegacyDocError('bad_stsh', detail)
}

interface StshReader {
  bytes: Uint8Array
  view: DataView
  end: number
}

function need(reader: StshReader, offset: number, length: number, what: string): void {
  if (offset < 0 || length < 0 || offset + length > reader.end)
    badStsh(`${what} termina fora da folha de estilos de ${reader.end} bytes`)
}

function u16(reader: StshReader, offset: number, what: string): number {
  need(reader, offset, 2, what)
  return reader.view.getUint16(offset, true)
}

function readName(reader: StshReader, offset: number, stdEnd: number): string {
  const cch = u16(reader, offset, 'o comprimento de um nome de estilo')
  const charsStart = offset + 2
  const charsBytes = cch * 2
  if (charsStart + charsBytes + 2 > stdEnd)
    badStsh(`um nome de estilo de ${cch} caracteres ultrapassa sua definição de estilo`)
  let name = ''
  for (let index = 0; index < cch; index++)
    name += String.fromCharCode(reader.view.getUint16(charsStart + index * 2, true))
  return name
}

function styleKind(stk: number, istd: number): StyleKind {
  switch (stk) {
    case STK_PARAGRAPH:
      return 'paragraph'
    case STK_CHARACTER:
      return 'character'
    case STK_TABLE:
      return 'table'
    case STK_NUMBERING:
      return 'numbering'
    default:
      badStsh(`o estilo ${istd} tem um tipo de estilo desconhecido (stk ${stk})`)
  }
}

/**
 * Parse the stylesheet at fcStshf..fcStshf+lcbStshf in the table stream.
 * Every offset/length is bounds-checked; a malformed stylesheet is a
 * LegacyDocError with code 'bad_stsh', never a raw RangeError.
 */
export function parseStsh(tableStream: Uint8Array, fc: number, lcb: number): StyleSheet {
  if (lcb === 0) return { styles: new Map(), defaultFonts: {} }
  if (fc > tableStream.length || lcb > tableStream.length - fc)
    badStsh(
      `a faixa da folha de estilos (${fc}+${lcb}) está fora do fluxo de tabela de ${tableStream.length} bytes`,
    )
  if (lcb < 2 + MIN_STSHI_BYTES)
    badStsh(`a folha de estilos tem apenas ${lcb} bytes, curta demais para seu cabeçalho`)
  const reader: StshReader = {
    bytes: tableStream,
    view: new DataView(tableStream.buffer, tableStream.byteOffset, tableStream.byteLength),
    end: fc + lcb,
  }
  const cbStshi = u16(reader, fc, 'o comprimento do STSHI')
  const stshiStart = fc + 2
  need(reader, stshiStart, cbStshi, 'o STSHI')
  if (cbStshi < MIN_STSHI_BYTES)
    badStsh(`o STSHI tem apenas ${cbStshi} bytes, menor que o mínimo de 18 do Word 97`)
  const cstd = u16(reader, stshiStart, 'a contagem de estilos')
  const cbSTDBase = u16(reader, stshiStart + 2, 'o comprimento base do STD')
  if (cbSTDBase !== WORD_97_STD_BASE && cbSTDBase !== WORD_2000_STD_BASE)
    badStsh(
      `a folha de estilos usa um comprimento base STD não suportado (${cbSTDBase}; esperado 0x0A para Word 97 ou 0x12 para Word 2000+)`,
    )
  if (cstd > MAX_STYLES)
    badStsh(`a folha de estilos declara ${cstd} estilos, mais que os 4096 que o Word pode endereçar`)
  const ftcAt = (at: number): number | undefined =>
    cbStshi >= at + 2 ? reader.view.getUint16(stshiStart + at, true) : undefined
  const defaultFonts: DefaultFontIndices = {
    ...(ftcAt(12) !== undefined ? { ascii: ftcAt(12) } : {}),
    ...(ftcAt(14) !== undefined ? { eastAsia: ftcAt(14) } : {}),
    ...(ftcAt(16) !== undefined ? { other: ftcAt(16) } : {}),
    ...(ftcAt(18) !== undefined ? { cs: ftcAt(18) } : {}),
  }
  const styles = new Map<number, Style>()
  let offset = stshiStart + cbStshi
  for (let istd = 0; istd < cstd; istd++) {
    const cbStd = u16(reader, offset, `o comprimento do estilo ${istd}`)
    offset += 2
    if (cbStd === 0) {
      // empty slots keep their index meaningful for istd references and carry no properties
      styles.set(istd, {
        istd,
        sti: 0,
        kind: 'paragraph',
        name: '',
        papxSprms: EMPTY,
        chpxSprms: EMPTY,
      })
      continue
    }
    const stdStart = offset
    const stdEnd = stdStart + cbStd
    need(reader, stdStart, cbStd, `o estilo ${istd}`)
    if (cbStd < cbSTDBase + 6)
      badStsh(`o estilo ${istd} tem apenas ${cbStd} bytes, curto demais para sua base e cabeçalho de nome`)
    const info1 = u16(reader, stdStart, `info1 do estilo ${istd}`)
    const info2 = u16(reader, stdStart + 2, `info2 do estilo ${istd}`)
    const info3 = u16(reader, stdStart + 4, `info3 do estilo ${istd}`)
    const sti = info1 & 0x0fff
    const kind = styleKind(info2 & 0x000f, istd)
    const istdBase = (info2 >> 4) & 0x0fff
    const cupx = info3 & 0x000f
    const istdNext = (info3 >> 4) & 0x0fff
    const name = readName(reader, stdStart + cbSTDBase, stdEnd)
    let upxOffset = stdStart + cbSTDBase + 2 + name.length * 2 + 2
    const upx: Uint8Array[] = []
    for (let index = 0; index < cupx; index++) {
      const cbUPX = u16(reader, upxOffset, `o comprimento do UPX ${index} do estilo ${istd}`)
      upxOffset += 2
      if (upxOffset + cbUPX > stdEnd)
        badStsh(`o UPX ${index} do estilo ${istd} ultrapassa sua definição de estilo`)
      upx.push(tableStream.subarray(upxOffset, upxOffset + cbUPX))
      upxOffset += cbUPX
      // UPX entries are word-aligned relative to the start of the STD
      if ((upxOffset - stdStart) % 2 === 1) upxOffset++
    }
    let papxSprms: Uint8Array = EMPTY
    let chpxSprms: Uint8Array = EMPTY
    if (kind === 'paragraph') {
      const papx = upx[0]
      if (papx !== undefined) {
        if (papx.length < 2)
          badStsh(`o estilo de parágrafo ${istd} tem um PAPX menor que seu prefixo istd`)
        papxSprms = papx.subarray(2)
      }
      chpxSprms = upx[1] ?? EMPTY
    } else if (kind === 'character') {
      chpxSprms = upx[0] ?? EMPTY
    }
    styles.set(istd, {
      istd,
      sti,
      kind,
      ...(istdBase === ISTD_BASE_NONE ? {} : { basedOn: istdBase }),
      ...(kind === 'paragraph' && istdNext !== ISTD_BASE_NONE ? { next: istdNext } : {}),
      name,
      papxSprms,
      chpxSprms,
    })
    offset = stdEnd
    // STD entries are padded to a 2-byte boundary relative to the STSH start
    // (fcStshf itself is not necessarily even).
    if ((offset - fc) % 2 === 1) offset++
  }
  return { styles, defaultFonts }
}
