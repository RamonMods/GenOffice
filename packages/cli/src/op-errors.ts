import type { ErrorReason } from './result'
import { didYouMean } from './suggest'

export interface OpFailure {
  index: number
  op: string
  error: string
  reason: ErrorReason
  valid_range?: [number, number]
  available?: string[]
  supported?: string[]
  usage?: string
  did_you_mean?: string
}

/**
 * The executors' guided errors are prose written for a model: "no element
 * "e_9" on slide 2. Available: [e_2, e_3]." The same facts are lifted into
 * fields here so a program branches on `reason` and reads the ids from
 * `available` instead of parsing the sentence. The message stays verbatim.
 */
export function classifyOpError(index: number, op: string, error: string): OpFailure {
  const f: OpFailure = { index, op, error, reason: 'op_rejected' }
  const usage = /(?:^|\n)(Usage: .+?)(?: Nothing was applied\b.*)?$/s.exec(error)?.[1]
  if (usage) f.usage = usage.trim()
  const list = (label: string): string[] | undefined => {
    const m = new RegExp(`${label}: \\[([^\\]]*)\\]`).exec(error)
    if (!m) return undefined
    return m[1]
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean)
  }
  const supported = list('Supported ops')
  const available = list('Available(?: groups)?')
  const range = /out of (?:range|bounds) \((-?\d+)-(-?\d+)\)/.exec(error)
  const blocks = /block index invalid or out of range \(the document has (\d+) blocks?\)/.exec(
    error,
  )
  if (/unknown op\b/i.test(error)) {
    f.reason = 'unknown_op'
    if (supported) {
      f.supported = supported
      const guess = didYouMean(op, supported)
      if (guess) f.did_you_mean = guess
    }
  } else if (range) {
    f.reason = 'out_of_range'
    f.valid_range = [Number(range[1]), Number(range[2])]
  } else if (blocks) {
    f.reason = 'out_of_range'
    f.valid_range = [0, Number(blocks[1]) - 1]
  } else if (
    available ||
    /\bno (?:element|slide|group|master\/layout part|matching blocks?)\b/i.test(error)
  ) {
    f.reason = 'target_not_found'
    if (available) {
      f.available = available
      // entries read "sp_0 (e_2)": either id may be what the caller mistyped
      const wanted = /no (?:element|slide|group|master\/layout part) "([^"]+)"/.exec(error)?.[1]
      const guess =
        wanted &&
        didYouMean(
          wanted,
          available.flatMap((a) => a.replace(/[()]/g, '').split(/\s+/)),
        )
      if (guess) f.did_you_mean = guess
    }
  }
  return f
}

export function opSuggestion(f: OpFailure, domain: 'slides' | 'docs' | 'sheets'): string {
  switch (f.reason) {
    case 'unknown_op':
      return f.did_you_mean
        ? `você quis dizer "${f.did_you_mean}"? (\`genoffice guide ${domain}\` lista todas as ops), depois reenvie o lote inteiro`
        : `execute \`genoffice guide ${domain}\` para ver a lista de ops, depois reenvie o lote inteiro`
    case 'out_of_range':
      return `use um índice entre ${f.valid_range![0]} e ${f.valid_range![1]} (\`genoffice ${domain} read <file> --json\` lista os atuais), depois reenvie o lote inteiro`
    case 'target_not_found':
      return f.did_you_mean
        ? `você quis dizer "${f.did_you_mean}"? (detail.failures[].available lista os ids nesse slide), depois reenvie o lote inteiro`
        : `execute \`genoffice ${domain} read <file> --json\` para obter os ids atuais${f.available ? ' (veja detail.failures[].available)' : ''}, depois reenvie o lote inteiro`
    default:
      return f.usage
        ? `corrija a op ${f.index} para corresponder à linha de uso (detail.failures[].usage), depois reenvie o lote inteiro`
        : `corrija a op ${f.index}, depois reenvie o lote inteiro`
  }
}
