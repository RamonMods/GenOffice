import { LegacyDocError } from './errors'

/** One piece of the document's text: a CP range and where its bytes live. */
export interface Piece {
  /** first character position (CP) of the piece */
  cpStart: number
  /** one past the last character position of the piece */
  cpEnd: number
  /** byte offset of the piece's text inside the WordDocument stream */
  fc: number
  /** true: one byte per character (Windows-1252); false: UTF-16LE (two bytes per character) */
  compressed: boolean
}

export interface PieceTable {
  /** character positions of the piece boundaries; cps.length === pieces.length + 1 */
  cps: number[]
  pieces: Piece[]
}

export interface ParseClxOptions {
  /** body-story character count (FibRgLw97.ccpText); the CPs of all stories are chained, the body is [0, ccpText) */
  ccpText: number
  /** WordDocument stream length, for piece-range validation */
  wordDocumentLength: number
}

/** CLX record types (MS-DOC 2.9.38): 0x01 Prc (property modifier), 0x02 Pcdt (piece table) */
const CLX_PRC = 0x01
const CLX_PCDT = 0x02

/** Pcdt payload = (n + 1) CPs + n PCDs of 8 bytes: lcb = 4 * (n + 1) + 8 * n = 12n + 4 */
const PCD_BYTES = 8
const CP_BYTES = 4

const FC_COMPRESSED = 0x40000000
/** bits 0-29 of Pcd.fc carry the offset; bit 31 is reserved and must be zero */
const FC_OFFSET_MASK = 0x3fffffff
const FC_RESERVED = 0x80000000

function badClx(detail: string): never {
  throw new LegacyDocError('bad_clx', detail)
}

function badPieceTable(detail: string): never {
  throw new LegacyDocError('bad_piece_table', detail)
}

/**
 * Parse the CLX (piece table) of a Word 97-2003 document: walk the Prc
 * records, then read the Pcdt's PlcPcd into character ranges with their
 * WordDocument byte ranges. Only the body story ([0, ccpText)) is read later,
 * but every piece is validated so a broken table fails here, not mid-text.
 */
export function parseClx(
  table: Uint8Array,
  fcClx: number,
  lcbClx: number,
  options: ParseClxOptions,
): PieceTable {
  if (lcbClx < CP_BYTES + 1) badClx(`o comprimento do CLX ${lcbClx} é pequeno demais para um registro Pcdt`)
  if (fcClx > table.length || lcbClx > table.length - fcClx)
    badClx(`a faixa do CLX (${fcClx}+${lcbClx}) está fora do fluxo de tabela de ${table.length} bytes`)
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength)
  const end = fcClx + lcbClx

  let offset = fcClx
  let plcStart = -1
  let plcLength = 0
  while (offset < end) {
    const kind = table[offset]!
    if (kind === CLX_PRC) {
      if (offset + 3 > end) badClx('o cabeçalho de um registro Prc termina fora do CLX')
      const cb = view.getUint16(offset + 1, true)
      if (offset + 3 + cb > end) badClx(`um registro Prc de ${cb} bytes termina fora do CLX`)
      offset += 3 + cb
      continue
    }
    if (kind === CLX_PCDT) {
      if (offset + 5 > end) badClx('o cabeçalho do registro Pcdt termina fora do CLX')
      const lcb = view.getUint32(offset + 1, true)
      if (lcb > end - offset - 5) badClx(`o conteúdo Pcdt de ${lcb} bytes termina fora do CLX`)
      plcStart = offset + 5
      plcLength = lcb
      break
    }
    badClx(`tipo de registro CLX desconhecido 0x${kind.toString(16)} no deslocamento ${offset - fcClx}`)
  }
  if (plcStart < 0) badClx('o CLX não contém um registro de tabela de peças Pcdt')
  if (plcLength < CP_BYTES)
    badClx(`o conteúdo PlcPcd tem ${plcLength} bytes, curto demais para um único CP`)
  if ((plcLength - CP_BYTES) % (CP_BYTES + PCD_BYTES) !== 0)
    badClx(`o comprimento do conteúdo PlcPcd ${plcLength} não é 12n+4`)

  const pieceCount = (plcLength - CP_BYTES) / (CP_BYTES + PCD_BYTES)
  if (pieceCount === 0) {
    if (options.ccpText > 0)
      badPieceTable(`a tabela de peças está vazia, mas ccpText é ${options.ccpText}`)
    return { cps: [], pieces: [] }
  }

  const cps: number[] = []
  for (let index = 0; index <= pieceCount; index++)
    cps.push(view.getUint32(plcStart + index * CP_BYTES, true))
  if (cps[0] !== 0) badPieceTable(`o primeiro CP é ${cps[0]}, e não 0`)
  for (let index = 1; index <= pieceCount; index++) {
    if (cps[index]! <= cps[index - 1]!)
      badPieceTable(
        `os limites das peças não são estritamente crescentes (${cps[index - 1]} -> ${cps[index]})`,
      )
  }
  // the CPs chain every story (body, footnotes, headers, ...); this stage reads
  // the body, so the table must at least cover it
  if (cps[pieceCount]! < options.ccpText)
    badPieceTable(
      `a tabela de peças termina no CP ${cps[pieceCount]}, antes do texto principal de ${options.ccpText} caracteres`,
    )

  const pieces: Piece[] = []
  const pcdsStart = plcStart + (pieceCount + 1) * CP_BYTES
  for (let index = 0; index < pieceCount; index++) {
    const pcdStart = pcdsStart + index * PCD_BYTES
    // u16 flags (fNoParaLast / fPaphNil / fCopied: not needed to read text)
    const fcRaw = view.getUint32(pcdStart + 2, true)
    const chars = cps[index + 1]! - cps[index]!
    if ((fcRaw & FC_RESERVED) !== 0) badPieceTable(`peça ${index}: o bit 31 reservado de fc está ativado`)
    const compressed = (fcRaw & FC_COMPRESSED) !== 0
    const masked = fcRaw & FC_OFFSET_MASK
    if (compressed && (masked & 1) !== 0)
      badPieceTable(`peça ${index}: uma peça comprimida tem deslocamento de byte ímpar`)
    const fc = compressed ? masked / 2 : masked
    const byteLength = compressed ? chars : chars * 2
    if (fc > options.wordDocumentLength || byteLength > options.wordDocumentLength - fc)
      badPieceTable(
        `a faixa de texto da peça ${index} (${fc}+${byteLength}) está fora do fluxo WordDocument de ${options.wordDocumentLength} bytes`,
      )
    pieces.push({
      cpStart: cps[index]!,
      cpEnd: cps[index + 1]!,
      fc,
      compressed,
    })
  }
  return { cps, pieces }
}
