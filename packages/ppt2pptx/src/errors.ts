/**
 * Error model for the legacy .ppt reader. Every failure the parser can detect
 * is a LegacyPptError with a stable code, so callers (the import UI in later
 * stages) can react without parsing messages, and a malformed file can never
 * surface as a raw RangeError/TypeError.
 */
export const LEGACY_PPT_ERROR_CODES = [
  'not_cfb',
  'truncated',
  'missing_stream',
  'unsupported_version',
  'encrypted',
  'limits_exceeded',
  'corrupt',
  'bad_persist',
  'bad_slide',
] as const

export type LegacyPptErrorCode = (typeof LEGACY_PPT_ERROR_CODES)[number]

/** What the user can do about each failure; appended to every message. */
const ADVICE: Record<LegacyPptErrorCode, string> = {
  not_cfb: 'O arquivo não parece ser uma apresentação do PowerPoint 97-2003.',
  truncated: 'O arquivo é menor do que a própria estrutura indica.',
  missing_stream: 'O fluxo da apresentação está ausente no arquivo.',
  unsupported_version: 'Apenas arquivos do PowerPoint 97-2003 (.ppt) são compatíveis.',
  encrypted:
    'A apresentação está protegida por senha; remova a senha no PowerPoint antes de convertê-la.',
  limits_exceeded: 'A apresentação é grande demais para ser convertida.',
  corrupt: 'A estrutura da apresentação está danificada e não pode ser analisada.',
  bad_persist: 'O diretório de alterações da apresentação está danificado.',
  bad_slide: 'Um contêiner de slide está danificado.',
}

export class LegacyPptError extends Error {
  readonly code: LegacyPptErrorCode

  constructor(code: LegacyPptErrorCode, detail?: string) {
    const advice = ADVICE[code]
    super(detail ? `${advice} (${detail})` : advice)
    this.name = 'LegacyPptError'
    this.code = code
  }
}
