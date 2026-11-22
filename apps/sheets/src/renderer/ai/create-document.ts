/// AI create_document, renderer side: xlsx/csv serialize one worksheet's
/// display grid (values as shown, like the manual CSV export) and hand the
/// text to the main process to write into the default save folder; docx/pdf/md/html
/// forward AI-authored content untouched (the shell routes them into the
/// docs-owned creation flow).

import {
  cellPatternLookup,
  csvSheetById,
  serializeActiveSheetCsv,
  sheetHasFormulas,
} from '../csv-export'
import { settleFormulas } from '../formula-settle'
import type { LazyWorkbookState, UniverRuntime } from '../univer-state'
import type { CreateDocumentToolOutcome, CreateDocumentToolRequest } from './tools'

export interface AiCreateDocumentContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
}

export async function createAiDocument(
  ctx: AiCreateDocumentContext,
  request: CreateDocumentToolRequest,
): Promise<CreateDocumentToolOutcome> {
  if (
    request.type === 'docx' ||
    request.type === 'pdf' ||
    request.type === 'md' ||
    request.type === 'html'
  ) {
    const result = await window.desktopApi.createDocument(request)
    if (!result.ok) return { ok: false, error: result.error ?? 'falha ao criar o documento' }
    return {
      ok: true,
      name: `${request.title}.${request.type}`,
      ...(result.path ? { path: result.path } : {}),
    }
  }
  const sheet = csvSheetById(ctx.univerRef.current, request.sheetId)
  if (!sheet) {
    return {
      ok: false,
      error: 'A planilha não está disponível (nenhuma pasta de trabalho aberta ou sheetId desconhecido).',
    }
  }
  const state = ctx.lazyWorkbookRef.current
  // File-backed sheets of a not-fully-loaded workbook would serialize only
  // the streamed-in region — refuse instead of silently exporting partial
  // data. Sheets added this session live entirely in the grid and are exempt
  // (same rule as ensureRangeLoaded).
  if (
    state &&
    !state.flags.preloadComplete &&
    state.file.sheets.some((fileSheet) => fileSheet.id === sheet.getSheetId())
  ) {
    return {
      ok: false,
      error:
        'A pasta de trabalho ainda não terminou de carregar, então os dados desta planilha ainda não estão totalmente na grade — tente novamente em instantes. ' +
        'Em pastas de trabalho transmitidas muito grandes, copie as linhas necessárias para uma nova planilha (add_sheet + copy_range) e exporte essa planilha.',
    }
  }
  const runtime = ctx.univerRef.current
  await settleFormulas(runtime)
  if (ctx.lazyWorkbookRef.current !== state) {
    return { ok: false, error: 'A pasta de trabalho mudou enquanto as fórmulas eram calculadas — tente novamente.' }
  }
  const content = serializeActiveSheetCsv(
    sheet,
    state,
    runtime
      ? cellPatternLookup(runtime, sheet.getSheetId(), state?.file.date1904 === true)
      : undefined,
  )
  if (content === 'too-large') {
    return { ok: false, error: 'A planilha é grande demais para ser exportada como arquivo independente.' }
  }
  const sheetName = sheet.getSheetName()
  const title = request.title?.trim() || sheetName
  const result = await window.desktopApi.createDocument({
    type: request.type,
    title,
    content,
    ...(request.type === 'xlsx' ? { sheetName } : {}),
  })
  if (!result.ok) return { ok: false, error: result.error ?? 'falha ao criar o arquivo' }
  return {
    ok: true,
    name: `${title}.${request.type}`,
    sheetName,
    hadFormulas: sheetHasFormulas(sheet),
    ...(result.path ? { path: result.path } : {}),
  }
}
