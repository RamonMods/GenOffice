/**
 * Error model for the legacy .doc reader. Every failure the parser can detect
 * is a LegacyDocError with a stable code, so callers (the import UI in later
 * stages) can react without parsing messages, and a malformed file can never
 * surface as a raw RangeError/TypeError.
 */
export const LEGACY_DOC_ERROR_CODES = [
  'not_cfb',
  'truncated',
  'bad_fib',
  'unsupported_version',
  'encrypted',
  'missing_stream',
  'bad_clx',
  'bad_piece_table',
  'limits_exceeded',
  'corrupt',
  'bad_fkp',
  'bad_sprm',
  'bad_stsh',
  'bad_section',
  'bad_hdd',
  'bad_note',
  'bad_table',
  'bad_picture',
  'bad_list',
  'bad_atn',
] as const

export type LegacyDocErrorCode = (typeof LEGACY_DOC_ERROR_CODES)[number]

/** What the user can do about each failure; appended to every message. */
const HINTS: Record<LegacyDocErrorCode, string> = {
  not_cfb:
    'o arquivo não é um documento composto OLE2; um .docx renomeado para .doc é importado pelo leitor de Office Open XML',
  truncated:
    'o arquivo está incompleto (download ou cópia truncados); obtenha uma cópia íntegra do arquivo .doc',
  bad_fib: 'o cabeçalho do Word é inválido; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  unsupported_version:
    'somente documentos do Word 97-2003 são suportados; abra o arquivo no Word ou no LibreOffice e salve-o como .docx',
  encrypted:
    'o documento está protegido por senha; remova a senha no Word ou no LibreOffice antes de importá-lo',
  missing_stream:
    'o documento não contém um fluxo que todo arquivo do Word precisa; abra-o no Word ou no LibreOffice e salve-o novamente',
  bad_clx:
    'a tabela de texto do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_piece_table:
    'a tabela de texto do documento aponta para fora do arquivo; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  limits_exceeded:
    'o documento é maior que o limite de importação legada; divida-o ou converta-o para .docx no Word ou no LibreOffice',
  corrupt: 'o documento está corrompido; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_fkp:
    'a tabela de formatação do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_sprm:
    'os registros de formatação do documento são inválidos; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_stsh:
    'a folha de estilos do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_section:
    'a tabela de seções do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_hdd:
    'a tabela de cabeçalho/rodapé do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_note:
    'a tabela de notas de rodapé/fim do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_table:
    'as propriedades de tabela do documento são inválidas; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_picture:
    'os dados de imagem do documento são inválidos; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_list:
    'a formatação de lista do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
  bad_atn:
    'a tabela de comentários do documento é inválida; abra o arquivo no Word ou no LibreOffice e salve-o novamente',
}

/** A legacy .doc that cannot be read, with a machine-readable reason code. */
export class LegacyDocError extends Error {
  readonly code: LegacyDocErrorCode

  constructor(code: LegacyDocErrorCode, detail: string) {
    super(`Não foi possível importar este arquivo .doc (${code}): ${detail}. ${HINTS[code]}.`)
    this.name = 'LegacyDocError'
    this.code = code
  }
}

export function isLegacyDocError(error: unknown): error is LegacyDocError {
  return error instanceof LegacyDocError
}
