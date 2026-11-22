import { clipText } from '../preview'
import type { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import {
  applyTrackedChangesXml,
  BLANK_BULLET_NUM_ID,
  BLANK_ORDERED_NUM_ID,
  buildBlankDocx,
  cleanMoveRangeMarkers,
  compareDocuments,
  countTrackedChangesXml,
  findChartWorkbookPath,
  nextNoteId,
  parseChartPartXml,
  parseDocx,
  parsePropertyChanges,
  patchChartPartXml,
  patchChartWorkbookXlsxBase64,
  pendingHeadingLevel,
  planRemoveParts,
  readDocxPartBase64,
  readSections,
  removePartsRefusalMessage,
  removePartsRefusalMessageFor,
  saveDocx,
  type CommentInfo,
  type CompareGranularity,
  type CompareIgnore,
  type CompareSummary,
  type DocumentProperties,
  type HeaderFooter,
  type HfPartInfo,
  type NoteInfo,
  type PropertyChanges,
  type RemovePartsOptions,
  type RemovedPartBinding,
  type RemovePartsPlan,
  type RevisionXmlFilter,
  type SaveOptions,
  type SectionInfo,
  type StyleUpsert,
  type PictureWatermarkSpec,
  type WatermarkSpec,
  isContentControlLock,
  type ContentControlSpec,
} from '@genoffice/docx-engine'
import type { FloatSpec } from '../../../../apps/docs/src/renderer/ai/floating-ops'
import type { AiNotesAccess, NoteKind } from '../../../../apps/docs/src/renderer/ai/note-ops'
import type { AiCommentsAccess } from '../../../../apps/docs/src/renderer/ai/tools'
import { ensureDom } from '../dom'
import type { PathContext } from '../fs'
import type { BatchMode } from '../batch'
import { classifyOpError, opSuggestion, type OpFailure } from '../op-errors'
import {
  signatureFromSchema,
  withFingerprint,
  type JsonSchema,
  type OpCatalog,
  type OpEntry,
} from '../op-catalog'
import { CliError, EXIT } from '../result'
import { readImageSource } from './image-source'
import { imageSize } from './image-size'
import { applySectionEdits, pageSetupAccess } from './docx-sections'

/**
 * The Word editing surface lives in the docs renderer (Tiptap document,
 * restricted-HTML parser, op executor, save plan). Those modules are pure
 * apart from needing a DOM, so genoffice runs them under jsdom. They are loaded
 * lazily, after the DOM exists, and by relative path until they move into a
 * package of their own.
 */
async function docsModules() {
  await ensureDom()
  const [
    { Editor },
    extensions,
    convert,
    protocol,
    ops,
    tools,
    locale,
    comments,
    hfText,
    revisions,
    pageSetup,
    revisionOps,
    floating,
    noteOps,
    tableOps,
    fieldResults,
  ] = await Promise.all([
    import('@tiptap/core'),
    import('../../../../apps/docs/src/renderer/editor/extensions'),
    import('../../../../apps/docs/src/renderer/editor/convert'),
    import('../../../../apps/docs/src/renderer/ai/protocol'),
    import('../../../../apps/docs/src/renderer/ai/ops'),
    import('../../../../apps/docs/src/renderer/ai/tools'),
    import('../../../../apps/docs/src/renderer/i18n/locale'),
    import('../../../../apps/docs/src/renderer/editor/comments'),
    import('../../../../apps/docs/src/renderer/editor/hf-text'),
    import('../../../../apps/docs/src/renderer/editor/revisions'),
    import('../../../../apps/docs/src/renderer/ai/page-setup'),
    import('../../../../apps/docs/src/renderer/ai/revision-ops'),
    import('../../../../apps/docs/src/renderer/ai/floating-ops'),
    import('../../../../apps/docs/src/renderer/ai/note-ops'),
    import('../../../../apps/docs/src/renderer/ai/table-ops'),
    import('../../../../apps/docs/src/renderer/editor/field-results'),
  ])
  locale.setModuleLang('pt')
  return {
    Editor,
    extensions,
    convert,
    protocol,
    ops,
    tools,
    comments,
    hfText,
    revisions,
    pageSetup,
    revisionOps,
    floating,
    noteOps,
    tableOps,
    fieldResults,
  }
}

type Parsed = Awaited<ReturnType<typeof parseDocx>>
type Modules = Awaited<ReturnType<typeof docsModules>>

type HfView = 'default' | 'first' | 'even'
type HfSlot = `${'header' | 'footer'}${'' | 'First' | 'Even'}`

/** One pending header/footer edit for a section other than the last (see side.hfSection). */
interface SectionHfEdit {
  /** docxIndex of the paragraph carrying this section's sectPr; the engine locates the section by it */
  lastBlockIndex: number
  kind: 'header' | 'footer'
  variant: HfView
  hf: HeaderFooter
}

/** Header/footer parts, comments and notes live outside the ProseMirror document; edits are kept here until save. */
interface SideState {
  hf: Partial<Record<HfSlot, HeaderFooter | null>>
  hfDirty: Set<HfSlot>
  /**
   * Pending header/footer edits of the non-last sections of a multi-section document,
   * keyed by sectionHfKey. saveDocument turns them into SaveOptions.sectionHf (the
   * trailing section keeps using the hf/hfDirty slots above, as the app routes them).
   */
  hfSection: Map<string, SectionHfEdit>
  titlePg: boolean
  evenOddHf: boolean
  titlePgDirty: boolean
  evenOddHfDirty: boolean
  comments: CommentInfo[]
  commentsDirty: boolean
  /** pending sectPr rewrites by the docxIndex of the block carrying them (section breaks, trailing) */
  sectPr: Map<number, string>
  /** define_style results by styleId, written into styles.xml on save */
  styleUpserts: Map<string, StyleUpsert>
  /** set_watermark result; undefined = header untouched */
  watermark?: WatermarkSpec | PictureWatermarkSpec | null
  watermarkText: string | null
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  notesDirty: boolean
  /** header / footer / footnote / endnote part XML by zip path, loaded on the first accept_changes / reject_changes */
  revisionParts: Map<string, string> | null
  revisionPartsDirty: Set<string>
  /** set_properties changes; written into the docProps parts on save */
  properties?: PropertyChanges
  /** remove_parts package scrub applied in the final save pass */
  removeParts?: RemovePartsOptions
  /** block-level content controls to write on save: wrap a range of body blocks in one new control, or unwrap the control with a tag */
  contentControls?: SaveOptions['contentControls']
}

export interface OpenDocument {
  parsed: Parsed
  editor: Editor
  numIds: { bullet: string | null; ordered: string | null }
  mods: Modules
  side: SideState
}

function hfFromPart(part: HfPartInfo | null | undefined): HeaderFooter | null {
  if (
    !part ||
    (!part.text && !part.hasPageNumber && part.paras.length === 0 && !part.images?.length)
  )
    return null
  return {
    text: part.text,
    pageNumber: part.hasPageNumber,
    paras: part.paras.length > 0 ? part.paras : undefined,
  }
}

function sideStateOf(parsed: Parsed): SideState {
  const dflt = (kind: 'header' | 'footer'): HeaderFooter | null => {
    const text = kind === 'header' ? parsed.headerText : parsed.footerText
    const pageNumber = kind === 'header' ? parsed.headerHasPageNumber : parsed.footerHasPageNumber
    const paras = kind === 'header' ? parsed.headerParas : parsed.footerParas
    return text || pageNumber || paras?.length
      ? { text: text ?? '', pageNumber, paras: paras ?? undefined }
      : null
  }
  return {
    hf: {
      header: dflt('header'),
      footer: dflt('footer'),
      headerFirst: hfFromPart(parsed.headerFirst),
      footerFirst: hfFromPart(parsed.footerFirst),
      headerEven: hfFromPart(parsed.headerEven),
      footerEven: hfFromPart(parsed.footerEven),
    },
    hfDirty: new Set(),
    hfSection: new Map(),
    titlePg: parsed.titlePg ?? false,
    evenOddHf: parsed.evenAndOddHeaders ?? false,
    titlePgDirty: false,
    evenOddHfDirty: false,
    comments: [...parsed.comments],
    commentsDirty: false,
    sectPr: new Map(),
    styleUpserts: new Map(),
    watermarkText: parsed.watermarkText ?? null,
    footnotes: [...parsed.footnotes],
    endnotes: [...parsed.endnotes],
    notesDirty: false,
    revisionParts: null,
    revisionPartsDirty: new Set(),
  }
}

export async function openDocument(bytes: Uint8Array): Promise<OpenDocument> {
  const mods = await docsModules()
  const parsed = await parseDocx(bytes)
  const editor = new mods.Editor({
    element: document.createElement('div'),
    extensions: mods.extensions.editorExtensions,
  })
  editor.commands.setContent(
    mods.convert.blocksToPmDoc(parsed.blocks, readSections(parsed)) as never,
  )
  editor.storage.listNumbering.styles = parsed.styles
  const numIds = {
    bullet: mods.protocol.findNumId(parsed.blocks, 'bullet') ?? BLANK_BULLET_NUM_ID,
    ordered: mods.protocol.findNumId(parsed.blocks, 'ordered') ?? BLANK_ORDERED_NUM_ID,
  }
  return { parsed, editor, numIds, mods, side: sideStateOf(parsed) }
}

export async function blankDocument(options?: {
  omitGeneratorFields?: boolean
}): Promise<OpenDocument> {
  return openDocument(await buildBlankDocx(options))
}

export async function saveDocument(doc: OpenDocument): Promise<Uint8Array> {
  // caption numbers are saved as results
  doc.mods.fieldResults.settleFieldResults(doc.editor, doc.parsed.blocks)
  const plan = doc.mods.convert.pmDocToSavePlan(doc.editor.getJSON() as never, doc.parsed.blocks)
  const { side } = doc
  const hf = (slot: HfSlot) => (side.hfDirty.has(slot) ? (side.hf[slot] ?? undefined) : undefined)
  const sectionHf = [...side.hfSection.values()].map((edit) => ({
    lastBlockIndex: edit.lastBlockIndex,
    kind: edit.kind,
    variant: edit.variant,
    hf: edit.hf,
  }))
  const sectionEdits = applySectionEdits(doc, plan)
  const options: SaveOptions = {
    ...sectionEdits.options,
    ...(sectionHf.length > 0 ? { sectionHf } : {}),
    trailingSectPrChange: plan.trailingSectPrChange,
    header: hf('header'),
    footer: hf('footer'),
    headerFirst: hf('headerFirst'),
    footerFirst: hf('footerFirst'),
    headerEven: hf('headerEven'),
    footerEven: hf('footerEven'),
    titlePg: side.titlePgDirty ? side.titlePg : undefined,
    evenAndOddHeaders: side.evenOddHfDirty ? side.evenOddHf : undefined,
    comments: side.commentsDirty ? side.comments : undefined,
    styleUpserts: side.styleUpserts.size > 0 ? [...side.styleUpserts.values()] : undefined,
    watermark: side.watermark,
    footnotes: side.notesDirty ? side.footnotes : undefined,
    endnotes: side.notesDirty ? side.endnotes : undefined,
    properties: side.properties,
    removeParts: side.removeParts,
    contentControls: side.contentControls,
    ...(await chartPartPatches(doc, plan.chartPatches)),
  }
  for (const path of side.revisionPartsDirty) {
    options.partXml = { ...options.partXml, [path]: side.revisionParts!.get(path)! }
  }
  return saveDocx(doc.parsed, sectionEdits.saveBlocks, options)
}

export interface DocxCompareOptions {
  /** author recorded on every revision */
  author?: string
  /** ISO timestamp recorded on every revision */
  date?: string
  /** at what granularity a changed paragraph is diffed */
  granularity?: CompareGranularity
  /** what the comparison leaves out */
  ignore?: readonly CompareIgnore[]
}

/**
 * The redline of two documents: the revised one with the differences marked as
 * tracked insertions and deletions, ready to be accepted or rejected.
 */
export async function compareDocxFiles(
  original: Uint8Array,
  revised: Uint8Array,
  options: DocxCompareOptions = {},
): Promise<{ bytes: Uint8Array; summary: CompareSummary }> {
  const [originalDoc, revisedDoc] = await Promise.all([parseDocx(original), parseDocx(revised)])
  const result = await compareDocuments(originalDoc, revisedDoc, options)
  return { bytes: result.bytes, summary: result.summary }
}

const REVISION_PART_RE = /^word\/(?:header\d*|footer\d*|footnotes|endnotes)\.xml$/

async function revisionParts(doc: OpenDocument): Promise<Map<string, string>> {
  if (doc.side.revisionParts) return doc.side.revisionParts
  const zip = await JSZip.loadAsync(doc.parsed.internal.originalBytes)
  const parts = new Map<string, string>()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (REVISION_PART_RE.test(name) && !entry.dir) parts.set(name, await entry.async('string'))
  }
  doc.side.revisionParts = parts
  return parts
}

/** tracked changes outside the body (headers, footers, footnotes, endnotes) */
export async function countRevisionsElsewhere(doc: OpenDocument): Promise<number> {
  let n = 0
  for (const xml of (await revisionParts(doc)).values()) n += countTrackedChangesXml(xml)
  return n
}

/** orphan move-range markers (unpaired starts / ends) in the body and the story parts; the save drops them */
async function countOrphanMoveRangeMarkers(doc: OpenDocument): Promise<number> {
  let n = 0
  const zip = await JSZip.loadAsync(doc.parsed.internal.originalBytes)
  const body = zip.file('word/document.xml')
  if (body) n += cleanMoveRangeMarkers(await body.async('string')).removed
  for (const xml of (await revisionParts(doc)).values()) n += cleanMoveRangeMarkers(xml).removed
  return n
}

/**
 * wrap_content_control / unwrap_content_control: the control is written on
 * save, so the op only records the change; the document body is untouched.
 */
function setContentControl(
  doc: OpenDocument,
  op: Record<string, unknown>,
  name: 'wrap_content_control' | 'unwrap_content_control',
): { error?: string; output?: string } {
  const { side } = doc
  const controls = side.contentControls ?? {}
  if (name === 'unwrap_content_control') {
    const tag = typeof op.tag === 'string' ? op.tag.trim() : ''
    if (!tag) return { error: 'unwrap_content_control precisa de uma tag' }
    const anyShell = doc.parsed.blocks.some((b) =>
      (b.sdtShells ?? (b.sdtShell ? [b.sdtShell] : [])).some((shell) => shell.tag === tag),
    )
    if (!anyShell) {
      // an inline (run-level) control is real and `docs read` lists it under
      // inlineContentControls, so "no content control" would mislead; unwrap
      // is a block-level op (genoffice#2554)
      const inline = doc.parsed.blocks.some((b) =>
        (b.inlineContentControls ?? []).some((c) => c.tag === tag),
      )
      if (inline)
        return {
          error: `o controle de conteúdo com a tag ${tag} é embutido (nível de execução); controles embutidos são reportados, mas não podem ser desembrulhados — unwrap_content_control funciona em controles de nível de bloco`,
        }
      return { error: `nenhum controle de conteúdo tem a tag ${tag}` }
    }
    side.contentControls = { ...controls, unwrap: [...(controls.unwrap ?? []), { tag }] }
    return { output: `desembrulhar o controle de conteúdo com a tag ${tag}` }
  }
  const start = Number(op.startBlockIndex)
  const end = Number(op.endBlockIndex ?? op.startBlockIndex)
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start) {
    return {
      error:
        'wrap_content_control precisa de startBlockIndex e endBlockIndex (endBlockIndex >= startBlockIndex)',
    }
  }
  const root = doc.editor.state.doc
  if (end >= root.childCount) {
    return { error: `o bloco ${end} está além do fim do documento (${root.childCount} blocos)` }
  }
  const docxIndexes: number[] = []
  for (let i = start; i <= end; i++) {
    const docxIndex = root.child(i).attrs.docxIndex
    if (typeof docxIndex !== 'number') {
      return { error: `o bloco ${i} não tem índice original no documento e não pode ser agrupado` }
    }
    docxIndexes.push(docxIndex)
  }
  const inRange = new Set(docxIndexes)
  for (const docxIndex of docxIndexes) {
    const group = doc.parsed.blocks[docxIndex]?.sdtShell?.group
    if (group === undefined) continue
    const straddles = doc.parsed.blocks.some(
      (b, i) => b.sdtShell?.group === group && !inRange.has(i),
    )
    if (straddles) {
      return {
        error: `os blocos ${start}-${end} cortam um controle de conteúdo existente; agrupe tudo ou nada`,
      }
    }
  }
  const spec: ContentControlSpec = {}
  if (typeof op.tag === 'string' && op.tag.trim()) spec.tag = op.tag.trim()
  if (typeof op.title === 'string' && op.title.trim()) spec.title = op.title.trim()
  if (op.lock !== undefined && op.lock !== null) {
    if (!isContentControlLock(op.lock)) {
      return {
        error: `bloqueio desconhecido "${String(op.lock)}": none, sdtLocked, contentLocked ou sdtContentLocked`,
      }
    }
    spec.lock = op.lock
  }
  side.contentControls = {
    ...controls,
    wrap: [...(controls.wrap ?? []), { docxIndexes, spec }],
  }
  return {
    output: `agrupar os blocos ${start}-${end} em um controle de conteúdo${spec.tag ? ` com a tag ${spec.tag}` : ''}`,
  }
}

/**
 * accept_changes / reject_changes: the body goes through the editor; the parts it
 * never loads are rewritten as XML. Positional selectors (ids, blockIndex,
 * blockRange) address the body only.
 */
async function applyRevisionOp(
  doc: OpenDocument,
  op: Record<string, unknown>,
  mode: 'accept' | 'reject',
): Promise<{ output?: string; error?: string }> {
  const { op: _op, ...input } = op
  const ops = doc.mods.revisionOps
  const sel = ops.validateSelector(input)
  if ('error' in sel) return { error: sel.error }
  const bodyOnly = !!sel.ids || sel.blockIndex !== undefined || !!sel.blockRange
  const body = ops.applyRevisionSelection(doc.editor, sel, mode)
  let elsewhere = 0
  if (!bodyOnly) {
    const filter: RevisionXmlFilter = {
      ...(sel.author !== undefined ? { author: sel.author } : {}),
      ...(sel.type !== undefined ? { type: sel.type } : {}),
      ...(sel.before !== undefined ? { before: sel.before } : {}),
    }
    for (const [path, xml] of await revisionParts(doc)) {
      const r = applyTrackedChangesXml(xml, mode, filter)
      if (r.xml === xml) continue
      elsewhere += r.applied
      doc.side.revisionParts!.set(path, r.xml)
      doc.side.revisionPartsDirty.add(path)
    }
  }
  if ('error' in body && elsewhere === 0) {
    // A file whose only tracked-change markup is orphan move-range markers
    // reports no pending changes, but the save already drops those markers
    // (genoffice#2349): accepting / rejecting such a file should succeed and
    // let the save clean them instead of rejecting it (genoffice#2554 case 3).
    const orphans = await countOrphanMoveRangeMarkers(doc)
    if (orphans > 0)
      return {
        output: `Não restam alterações registradas. ${orphans} marcador(es) de intervalo de movimentação órfão(s) será(ão) descartado(s) ao salvar o documento.`,
      }
    return { error: body.error }
  }
  const verb = mode === 'accept' ? 'Aceitas' : 'Rejeitadas'
  const done = 'error' in body ? [] : body.entries
  const parts: string[] = []
  if (done.length)
    parts.push(`${done.length} alteração(ões) registrada(s) (${done.map((e) => e.id).join(', ')})`)
  if (elsewhere) parts.push(`${elsewhere} em cabeçalhos, rodapés e notas`)
  const left = ops.listRevisionEntries(doc.editor.state.doc)
  const leftElsewhere = await countRevisionsElsewhere(doc)
  const still: string[] = []
  if (left.length) still.push(`Ainda ${ops.describePending(left)}`)
  if (leftElsewhere)
    still.push(
      `${left.length ? 'e' : 'Ainda'} ${leftElsewhere} em cabeçalhos, rodapés e notas${bodyOnly ? ' (ids se referem ao corpo; all: true ou author / type / before os alcançam)' : ''}`,
    )
  return {
    output: `${verb} ${parts.join(' e ')}. ${still.length ? `${still.join(' ')}.` : 'Não restam alterações registradas.'}`,
  }
}

/** edit_chart changes live in the chart's own part and its embedded workbook, not in the body XML */
async function chartPartPatches(
  doc: OpenDocument,
  patches: ReturnType<Modules['convert']['pmDocToSavePlan']>['chartPatches'],
): Promise<Pick<SaveOptions, 'partXml' | 'partBinary'>> {
  const partXml: Record<string, string> = {}
  const partBinary: Record<string, string> = {}
  const original = doc.parsed.internal.originalBytes
  for (const { partPath, patch } of patches) {
    const part = doc.parsed.extras.chartParts[partPath]
    if (!part) continue
    const patched = patchChartPartXml(part, patch)
    partXml[partPath] = patched
    const wbPath = await findChartWorkbookPath(original, partPath)
    const wb = wbPath ? await readDocxPartBase64(original, wbPath) : null
    const display = wb ? parseChartPartXml(patched, partPath) : null
    if (wbPath && wb && display) {
      const updated = await patchChartWorkbookXlsxBase64(
        wb,
        display.categories,
        display.series.map((s, i) => ({
          name: s.name ?? `Series${i + 1}`,
          values: s.values as (number | null)[],
        })),
      )
      if (updated) partBinary[wbPath] = updated
    }
  }
  return {
    ...(Object.keys(partXml).length ? { partXml } : {}),
    ...(Object.keys(partBinary).length ? { partBinary } : {}),
  }
}

export function closeDocument(doc: OpenDocument): void {
  doc.editor.destroy()
}

export interface BlockSummary {
  index: number
  type: string
  level?: number
  /** protected blocks: image, chart, field, formula … (targets for setImageProperties / edit_chart) */
  kind?: string
  text: string
  /** native tables: grid size, merged cells and the table style, for the table ops */
  table?: ReturnType<Modules['tableOps']['describeTable']>
  /** the block-level content control this block sits in (w:sdt), when it is in one */
  contentControl?: ContentControlSummary
  /**
   * every block-level control wrapping this block, outermost first; set
   * whenever the block sits in at least one control (so a single-control block
   * carries a one-element array, and contentControl keeps the outermost)
   */
  contentControls?: ContentControlSummary[]
  /** run-level (inline) controls inside this paragraph, read-only (their bytes stay in place) */
  inlineContentControls?: InlineContentControlSummary[]
  /** the text was clipped to the preview length; `--full` or `--max-chars` returns the rest */
  truncated?: true
}

/** a run-level (inline) w:sdt inside a paragraph */
export interface InlineContentControlSummary {
  tag: string
  title: string
  lock: string
  id: string
  /** the plain text inside the control */
  text: string
}

/** what a block-level content control (w:sdt) carries */
export interface ContentControlSummary {
  /** programmatic tag (w:sdtPr/w:tag) */
  tag: string
  /** human-visible title (w:sdtPr/w:alias) */
  title: string
  /** lock state (w:sdtPr/w:lock); empty when the control carries none */
  lock: string
  /** control id (w:sdtPr/w:id); empty when the control carries none */
  id: string
}

export const PREVIEW_CHARS = 200

/**
 * Text a non-text leaf contributes to a block's plain text. A soft line break
 * (Shift+Enter, <w:br/>) reads as "\n" so the two lines never run together;
 * page/column/wrap breaks and every other leaf keep contributing nothing.
 */
function plainLeafText(leaf: { type: { name: string }; attrs: Record<string, unknown> }): string {
  if (leaf.type.name !== 'hardBreak') return ''
  const { pageBreak, colBreak, wrapBreak, mathBreak } = leaf.attrs ?? {}
  return pageBreak || colBreak || wrapBreak || mathBreak ? '' : '\n'
}

/** Block list an agent targets ops at: index, node type, heading level, text (clipped to `maxChars`). */
export function describeDocument(
  doc: OpenDocument,
  range?: [number, number],
  maxChars = PREVIEW_CHARS,
): BlockSummary[] {
  const out: BlockSummary[] = []
  const root = doc.editor.state.doc
  const [start, end] = range ?? [0, root.childCount - 1]
  for (let i = start; i <= Math.min(end, root.childCount - 1); i++) {
    const node = root.child(i)
    // textContent skips leaves without a leafText; read the soft breaks back in
    const { text, truncated } = clipText(
      node.textBetween(0, node.content.size, '', plainLeafText),
      maxChars,
    )
    out.push({
      index: i,
      type: node.type.name.replace(/^doc/, '').toLowerCase(),
      ...(typeof node.attrs.level === 'number' ? { level: node.attrs.level } : {}),
      ...(node.type.name === 'docProtected' ? { kind: protectedKind(node.attrs) } : {}),
      ...(node.type.name === 'docTable' ? { table: doc.mods.tableOps.describeTable(node) } : {}),
      ...(contentControlOf(doc, node) ? { contentControl: contentControlOf(doc, node)! } : {}),
      ...(contentControlsOf(doc, node).length > 0
        ? { contentControls: contentControlsOf(doc, node) }
        : {}),
      ...(inlineControlsOf(doc, node).length > 0
        ? { inlineContentControls: inlineControlsOf(doc, node) }
        : {}),
      text,
      ...(truncated ? { truncated } : {}),
    })
  }
  return out
}

/** table styles defined in styles.xml (setTableStyle accepts the id or the name) */
export function listTableStyles(doc: OpenDocument): Array<{ id: string; name: string }> {
  return [...doc.parsed.styles.values()]
    .filter((s) => s.type === 'table')
    .map((s) => ({ id: s.styleId, name: s.name }))
}

/** the parsed block behind an editor block (the editor block keeps the body index of its original) */
function parsedBlockOf(doc: OpenDocument, block: Editor['state']['doc']) {
  const docxIndex = block.attrs.docxIndex
  if (typeof docxIndex !== 'number') return null
  return doc.parsed.blocks[docxIndex] ?? null
}

/** the content control (w:sdt) a block sits in: the editor block keeps the body index of its original, and the shell rides on it */
function contentControlOf(
  doc: OpenDocument,
  block: Editor['state']['doc'],
): ContentControlSummary | null {
  const shell = parsedBlockOf(doc, block)?.sdtShell
  if (!shell) return null
  return { tag: shell.tag, title: shell.alias, lock: shell.lock, id: shell.id }
}

/** every block-level control the block sits in, outermost first (a single control gives a one-element array) */
function contentControlsOf(
  doc: OpenDocument,
  block: Editor['state']['doc'],
): ContentControlSummary[] {
  const parsed = parsedBlockOf(doc, block)
  const chain = parsed?.sdtShells ?? (parsed?.sdtShell ? [parsed.sdtShell] : [])
  return chain.map((shell) => ({
    tag: shell.tag,
    title: shell.alias,
    lock: shell.lock,
    id: shell.id,
  }))
}

/** run-level controls inside the paragraph (read-only reporting) */
function inlineControlsOf(
  doc: OpenDocument,
  block: Editor['state']['doc'],
): InlineContentControlSummary[] {
  return parsedBlockOf(doc, block)?.inlineContentControls ?? []
}

/** native docx charts are passthrough blocks carrying chartDisplay; edit_chart targets them like AI-made ones */
function protectedKind(attrs: Record<string, unknown>): string {
  if (attrs.chartDisplay) return 'chart'
  return typeof attrs.blockType === 'string' ? attrs.blockType : 'protected'
}

export function blockRangeHtml(doc: OpenDocument, start: number, end: number): string {
  return doc.mods.protocol.serializeRangeToHtml(doc.editor, start, end)
}

export interface ExportHtml {
  html: string
  /** blocks Markdown cannot carry, by kind */
  skipped: { images: number; fields: number }
  /** formulas whose LaTeX was not recoverable; emitted as their token text */
  formulasAsText: number
}

/**
 * The whole document as HTML for a human-facing export. The AI protocol's
 * serializer describes protected blocks (images, fields, formulas without
 * LaTeX) with "[Protected …]" notes meant for the model; here they become a
 * math node, the field's visible text, or nothing, and inline <formula> tags
 * take the markdown editor's math markup.
 */
export function documentHtml(doc: OpenDocument): ExportHtml {
  const top = doc.editor.state.doc
  const parts: string[] = []
  const skipped = { images: 0, fields: 0 }
  let formulasAsText = 0
  let runStart = -1
  const flush = (end: number) => {
    if (runStart >= 0) parts.push(blockRangeHtml(doc, runStart, end))
    runStart = -1
  }
  top.forEach((node, _offset, index) => {
    if (node.type.name !== 'docProtected') {
      if (runStart < 0) runStart = index
      return
    }
    flush(index - 1)
    // pending tracked deletions are dropped, as serializeRangeToHtml drops them
    if (doc.mods.protocol.isTrackedDeleted(node)) return
    const formula = node.attrs.formulaDisplay as { tokens?: string[]; latex?: string } | null
    if (formula?.latex) {
      parts.push(`<div data-type="block-math" data-latex="${escapeHtml(formula.latex)}"></div>`)
    } else if (formula?.tokens?.length) {
      formulasAsText++
      parts.push(`<p>${escapeHtml(formula.tokens.join(' '))}</p>`)
    } else if (node.attrs.blockType === 'image') {
      skipped.images++
    } else {
      const preview = String(node.attrs.previewText ?? '')
        .replace(/\s+/g, ' ')
        .trim()
      if (preview) parts.push(`<p>${escapeHtml(preview)}</p>`)
      else skipped.fields++
    }
  })
  flush(top.childCount - 1)
  const html = parts
    .join('\n')
    // the protocol escapes & < > in formula text but not quotes, which would end the attribute
    .replace(
      /<formula>([\s\S]*?)<\/formula>/g,
      (_m, latex: string) =>
        `<span data-type="inline-math" data-latex="${latex.replace(/"/g, '&quot;')}"></span>`,
    )
  return { html, skipped, formulasAsText }
}

function escapeHtml(s: string): string {
  return s.replace(
    /[&<>"]/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!,
  )
}

/** Replaces the whole body with a restricted-HTML fragment (the create_document contract). */
export function fillFromHtml(doc: OpenDocument, html: string): number {
  const nodes = doc.mods.protocol.parseHtmlFragment(html, doc.numIds)
  if (nodes.length === 0) {
    throw new CliError(EXIT.usage, 'não foi possível extrair nenhum conteúdo do HTML', undefined, {
      reason: 'invalid_argument',
    })
  }
  const last = doc.editor.state.doc.childCount - 1
  doc.mods.protocol.replaceBlockRange(doc.editor, 0, last, nodes)
  return nodes.length
}

export interface DocOpResult {
  index: number
  op: string
  output: string
}

export interface DocOpsOutcome {
  results: DocOpResult[]
  failures: OpFailure[]
  /** ops after the first rejection under stop_on_error, never attempted */
  skipped: number
}

const HTML_TOOLS = new Set(['insert_content', 'replace_blocks'])
/** app tools reachable headless besides the html pair; insert_image is built here (no renderer image bridge) */
const SIDE_TOOLS = new Set([
  'insert_chart',
  'edit_chart',
  'set_header_footer',
  'set_page_setup',
  'insert_section_break',
  'reply_comment',
  'resolve_comment',
  'add_comment',
  'delete_comment',
  'read_comments',
  'read_revisions',
  'accept_changes',
  'reject_changes',
  'define_style',
  'list_styles',
  'set_watermark',
  'insert_text_box',
  'insert_footnote',
  'insert_endnote',
  'delete_note',
  'edit_note',
  'read_notes',
])

/**
 * Block-level content controls (w:sdt). These are CLI-only ops: the app has
 * no AI tool for them, so the catalog carries their schema here.
 */
const CONTENT_CONTROL_OPS = ['wrap_content_control', 'unwrap_content_control'] as const

const CONTENT_CONTROL_SCHEMAS: Record<(typeof CONTENT_CONTROL_OPS)[number], JsonSchema> = {
  wrap_content_control: {
    type: 'object',
    properties: {
      startBlockIndex: {
        type: 'integer',
        description:
          'o primeiro bloco do intervalo a agrupar (base 0, como `genoffice docs read` os numera)',
      },
      endBlockIndex: {
        type: 'integer',
        description: 'o último bloco do intervalo a agrupar; o padrão é startBlockIndex',
      },
      tag: {
        type: 'string',
        description: 'tag programática (w:sdtPr/w:tag): como uma verificação posterior encontra o controle',
      },
      title: { type: 'string', description: 'título visível ao usuário (w:sdtPr/w:alias)' },
      lock: {
        type: 'string',
        enum: ['none', 'sdtLocked', 'contentLocked', 'sdtContentLocked'],
        description: 'estado de bloqueio; none não grava nenhum bloqueio',
      },
    },
    required: ['startBlockIndex'],
  },
  unwrap_content_control: {
    type: 'object',
    properties: {
      tag: {
        type: 'string',
        description:
          'a tag do controle a desembrulhar (o mais externo ou um controle interno aninhado; a tag pode se repetir)',
      },
    },
    required: ['tag'],
  },
}

const CONTENT_CONTROL_NOTES: Record<(typeof CONTENT_CONTROL_OPS)[number], string> = {
  wrap_content_control:
    'coloca um intervalo de blocos em um único controle de conteúdo de nível de bloco (w:sdt); os blocos continuam editáveis e o controle carrega a tag/título/bloqueio; `genoffice docs read` mostra a cadeia de controles em cada bloco do intervalo (contentControls, do mais externo para o mais interno) e os controles de nível de execução como inlineContentControls',
  unwrap_content_control:
    'remove todos os controles de conteúdo com essa tag e deixa seus blocos como blocos comuns do corpo; controles aninhados são desembrulhados nível por nível (desembrulhar uma tag interna mantém a casca externa ao redor do conteúdo); controles embutidos/de nível de execução são reportados, mas não desembrulhados',
}

/**
 * Package-level scrubs (remove_parts). CLI-only like the content controls: the
 * app has no AI tool for them, so the catalog carries their schema here.
 */
const PACKAGE_OPS = ['remove_parts'] as const

const PACKAGE_SCHEMAS: Record<(typeof PACKAGE_OPS)[number], JsonSchema> = {
  remove_parts: {
    type: 'object',
    properties: {
      kinds: {
        type: 'array',
        items: { type: 'string', enum: ['customXml', 'webExtensions'] },
        description:
          'famílias de partes a remover: customXml (armazenamentos de dados de modelos/GED: customXml/item*.xml + itemProps*.xml) ou webExtensions (painéis de tarefas/complementos em word/webextensions/)',
      },
      paths: {
        type: 'array',
        items: { type: 'string' },
        description:
          'caminhos explícitos de partes relativos à raiz do pacote (sem barra inicial), por exemplo "word/activeX/activeX1.bin"; caminhos desconhecidos são ignorados e listados',
      },
      dataBinding: {
        type: 'string',
        enum: ['refuse', 'strip'],
        description:
          'um armazenamento customXml removido pode dar suporte a controles de conteúdo por meio de w:dataBinding: refuse (padrão) falha o lote e nomeia os controles; strip remove a vinculação e mantém o controle e seu texto',
      },
    },
  },
}

const PACKAGE_NOTES: Record<(typeof PACKAGE_OPS)[number], string> = {
  remove_parts:
    'remove partes do pacote que contêm metadados em uma limpeza antes do arquivamento (armazenamentos customXml, partes de extensões web/complementos ou caminhos explícitos do pacote); seus relacionamentos, entradas em [Content_Types].xml e pastas vazias são removidos junto, e todas as outras partes mantêm seus bytes; controles de conteúdo vinculados a um armazenamento customXml removido fazem o lote falhar, a menos que dataBinding: "strip" desvincule esses controles',
}

export interface ApplyOptions {
  ctx?: PathContext
  /** folder of the ops file, for relative image paths */
  baseDir?: string
  /** record content edits as tracked changes by this author */
  trackAuthor?: string
  mode?: BatchMode
}

/**
 * Every entry goes through the app's executeTool: apply_ops entries as an
 * `apply_ops` call, the html and side tools as themselves. That is what keeps
 * the tool layer's stale-document guard in step (it records the document it
 * last touched), so mixed batches behave like one in-app turn.
 *
 * A dry run executes the batch on the in-memory editor and simply never
 * saves: that is the only way index shifts from earlier entries are taken
 * into account when later ones are validated.
 */
export async function applyDocOps(
  doc: OpenDocument,
  ops: Record<string, unknown>[],
  opts: ApplyOptions = {},
): Promise<DocOpsOutcome> {
  const mode = opts.mode ?? 'atomic'
  const results: DocOpResult[] = []
  const failures: OpFailure[] = []
  const comments = commentsAccess(doc)
  const hf = headerFooterAccess(doc)
  const pageSetup = pageSetupAccess(doc)
  const extras = docExtras(doc)
  const notes = notesAccess(doc)
  for (const [index, op] of ops.entries()) {
    const name = typeof op.op === 'string' ? op.op : ''
    // a rejected op left the document as it was: the executor applies each call whole or not at all
    const rejected = (output: string): boolean => {
      const failure = classifyOpError(index, name, output)
      if (mode === 'atomic') {
        throw new CliError(
          EXIT.usage,
          `operação ${index} (${name || '?'}) rejeitada: ${output}`,
          { failures: [failure] },
          { reason: failure.reason, suggestion: opSuggestion(failure, 'docs') },
        )
      }
      failures.push(failure)
      return true
    }
    let failed = false
    if (name === 'insert_image' || name === 'insert_picture') {
      const r = await insertImage(doc, op, opts, name === 'insert_picture')
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'set_watermark' && typeof op.image === 'string') {
      const r = await setPictureWatermark(doc, op, opts)
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'set_properties') {
      const r = setDocumentProperties(doc, op)
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'remove_parts') {
      const r = await setRemoveParts(doc, op)
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'wrap_content_control' || name === 'unwrap_content_control') {
      const r = setContentControl(doc, op, name)
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'set_header_footer' && op.section !== undefined) {
      const r = setSectionHeaderFooter(doc, op)
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else if (name === 'accept_changes' || name === 'reject_changes') {
      const r = await applyRevisionOp(doc, op, name === 'accept_changes' ? 'accept' : 'reject')
      if (r.error) failed = rejected(r.error)
      else results.push({ index, op: name, output: r.output! })
    } else {
      const call = HTML_TOOLS.has(name)
        ? { id: `genoffice-${index}`, name, input: htmlToolInput(doc, name, op) }
        : SIDE_TOOLS.has(name)
          ? { id: `genoffice-${index}`, name, input: sideToolInput(doc, name, op) }
          : { id: `genoffice-${index}`, name: 'apply_ops', input: { ops: [op] } }
      const exec = await doc.mods.tools.executeTool(
        doc.editor,
        call,
        doc.numIds,
        opts.trackAuthor ? { author: opts.trackAuthor } : undefined,
        undefined,
        null,
        comments,
        hf,
        undefined,
        pageSetup,
        extras,
        notes,
      )
      // the in-app executor reports a target that matched nothing as a soft result; for a
      // scripted batch that is a wrong index, so the op counts as rejected
      if (exec.isError || /^No matching blocks/i.test(exec.output)) failed = rejected(exec.output)
      else results.push({ index, op: name, output: exec.output })
    }
    if (failed && mode === 'stop_on_error') {
      return { results, failures, skipped: ops.length - index - 1 }
    }
  }
  return { results, failures, skipped: 0 }
}

/**
 * set_properties: the docProps parts (core, extended, custom) are written on
 * save, so the op only records the change; the document body is untouched.
 */
function setDocumentProperties(
  doc: OpenDocument,
  op: Record<string, unknown>,
): { output?: string; error?: string } {
  let changes: PropertyChanges
  try {
    changes = parsePropertyChanges(op)
  } catch (err) {
    return { error: (err as Error).message }
  }
  doc.side.properties = mergePropertyChanges(doc.side.properties, changes)
  return { output: `definir propriedades do documento: ${describePropertyChanges(changes)}` }
}

/** two set_properties ops in one batch compose: the later value wins, a later clear wipes the earlier values */
function mergePropertyChanges(
  current: PropertyChanges | undefined,
  next: PropertyChanges,
): PropertyChanges {
  if (!current || next.clear) return next
  const merged: PropertyChanges = { ...current }
  if (next.core) merged.core = { ...(current.core ?? {}), ...next.core }
  if (next.app) merged.app = { ...(current.app ?? {}), ...next.app }
  if (next.custom) merged.custom = { ...(current.custom ?? {}), ...next.custom }
  return merged
}

function describePropertyChanges(changes: PropertyChanges): string {
  const parts: string[] = []
  if (changes.clear === 'all') parts.push('limpar tudo')
  for (const section of ['core', 'app', 'custom'] as const) {
    for (const [key, value] of Object.entries(changes[section] ?? {})) {
      parts.push(`${section}.${key}=${value === null ? 'null' : JSON.stringify(value)}`)
    }
  }
  return parts.join(', ')
}

const REMOVE_PARTS_KINDS = ['customXml', 'webExtensions'] as const

const REMOVE_PARTS_FIELDS = new Set(['op', 'paths', 'kinds', 'dataBinding'])

/** Validate a remove_parts op; kinds are de-duplicated and at least one selector is required. */
function parseRemovePartsOp(op: Record<string, unknown>): {
  options?: RemovePartsOptions
  error?: string
} {
  // a silently ignored field would make the scrub quieter than the caller
  // believes: an unknown field is an op error, not a warning (genoffice#2554)
  const unknownFields = Object.keys(op).filter((key) => !REMOVE_PARTS_FIELDS.has(key))
  if (unknownFields.length > 0)
    return {
      error: `remove_parts: campo(s) desconhecido(s) ${unknownFields.map((f) => JSON.stringify(f)).join(', ')} (op, paths, kinds, dataBinding)`,
    }
  const kinds: NonNullable<RemovePartsOptions['kinds']> = []
  if (op.kinds !== undefined) {
    if (!Array.isArray(op.kinds))
      return { error: 'remove_parts precisa de kinds como uma matriz de customXml | webExtensions' }
    for (const kind of op.kinds) {
      if (!REMOVE_PARTS_KINDS.includes(kind as (typeof REMOVE_PARTS_KINDS)[number]))
        return {
          error: `remove_parts: tipo desconhecido ${JSON.stringify(kind)} (customXml | webExtensions)`,
        }
      const known = kind as (typeof REMOVE_PARTS_KINDS)[number]
      if (!kinds.includes(known)) kinds.push(known)
    }
  }
  const paths: string[] = []
  if (op.paths !== undefined) {
    if (!Array.isArray(op.paths))
      return { error: 'remove_parts precisa de paths como uma matriz de caminhos do pacote' }
    for (const path of op.paths) {
      if (typeof path !== 'string' || !path.trim())
        return { error: 'remove_parts: cada entrada de paths deve ser um caminho do pacote não vazio' }
      const trimmed = path.trim()
      if (!paths.includes(trimmed)) paths.push(trimmed)
    }
  }
  if (kinds.length === 0 && paths.length === 0)
    return {
      error:
        'remove_parts precisa de kinds (customXml | webExtensions) ou paths (caminhos de partes relativos ao pacote)',
    }
  if (op.dataBinding !== undefined && op.dataBinding !== 'refuse' && op.dataBinding !== 'strip')
    return { error: 'remove_parts: dataBinding deve ser "refuse" ou "strip"' }
  const options: RemovePartsOptions = {}
  if (kinds.length > 0) options.kinds = kinds
  if (paths.length > 0) options.paths = paths
  if (op.dataBinding === 'refuse' || op.dataBinding === 'strip')
    options.dataBinding = op.dataBinding
  return { options }
}

/** two remove_parts ops in one batch compose: union of kinds/paths, the later explicit dataBinding wins */
function mergeRemoveParts(
  current: RemovePartsOptions | undefined,
  next: RemovePartsOptions,
): RemovePartsOptions {
  const kinds = [...(current?.kinds ?? [])]
  for (const kind of next.kinds ?? []) if (!kinds.includes(kind)) kinds.push(kind)
  const paths = [...(current?.paths ?? [])]
  for (const path of next.paths ?? []) if (!paths.includes(path)) paths.push(path)
  const merged: RemovePartsOptions = {}
  if (kinds.length > 0) merged.kinds = kinds
  if (paths.length > 0) merged.paths = paths
  const dataBinding = next.dataBinding ?? current?.dataBinding
  if (dataBinding) merged.dataBinding = dataBinding
  return merged
}

function describeControlBinding(binding: RemovedPartBinding): string {
  const label = binding.tag ? `"${binding.tag}"` : 'controle sem tag'
  const id = binding.controlId ? ` w:id ${binding.controlId}` : ''
  return `${label}${id} in ${binding.part}`
}

/** One-line plan/report for remove_parts (the dry-run plan keeps the first line). */
function describeRemovePartsPlan(plan: RemovePartsPlan, options: RemovePartsOptions): string {
  const families = options.kinds?.length ? ` [${options.kinds.join(', ')}]` : ''
  const paths =
    plan.removedPaths.length > 0 ? plan.removedPaths.join(', ') : 'nada (já ausente)'
  const action = options.dataBinding ?? 'refuse'
  const binding =
    plan.dataBindings.length > 0
      ? `; dataBinding: ${action} desvincularia ${plan.dataBindings.length} controle(s) (${plan.dataBindings.map(describeControlBinding).join(', ')})`
      : `; dataBinding: ${action} (nenhum controle vinculado)`
  return `remover ${plan.removedPaths.length} parte(s) do pacote${families}: ${paths}${binding}`
}

/**
 * remove_parts: the parts are dropped in the final package pass on save, so the
 * op only records the change; the package is planned against the source bytes
 * here so the dry run lists the parts and a dataBinding refusal fails the batch
 * (nothing is written either way). Paths that miss the package or name parts
 * the package depends on are op errors — a scrub must not half-apply silently
 * (genoffice#2554).
 */
async function setRemoveParts(
  doc: OpenDocument,
  op: Record<string, unknown>,
): Promise<{ output?: string; error?: string }> {
  const parsed = parseRemovePartsOp(op)
  if (!parsed.options) return { error: parsed.error! }
  const merged = mergeRemoveParts(doc.side.removeParts, parsed.options)
  const plan = await planRemoveParts(doc.parsed.internal.originalBytes, merged)
  if (plan.refusedPaths.length > 0)
    return { error: removePartsRefusalMessageFor(plan.refusedPaths) }
  if (plan.unknownPaths.length > 0)
    return {
      error: `remove_parts: não existem essas parte(s) do pacote: ${plan.unknownPaths.join(', ')} (paths são caminhos de arquivo relativos ao pacote; curingas e pastas não são compatíveis)`,
    }
  const strip = merged.dataBinding === 'strip'
  if (plan.dataBindings.length > 0 && !strip)
    return { error: removePartsRefusalMessage(plan.dataBindings) }
  doc.side.removeParts = merged
  return { output: describeRemovePartsPlan(plan, merged) }
}

/** the document properties (core, extended/app, custom) of the open document */
export function listProperties(doc: OpenDocument): DocumentProperties {
  return doc.parsed.properties
}

/** insert_content / insert_chart without a position append at the end; the in-app default ("after the cursor") has no meaning headless. */
function htmlToolInput(
  doc: OpenDocument,
  name: string,
  op: Record<string, unknown>,
): Record<string, unknown> {
  const { op: _op, ...input } = op
  if (name === 'insert_content' && input.afterBlockIndex === undefined) {
    input.afterBlockIndex = doc.editor.state.doc.childCount - 1
  }
  return input
}

function sideToolInput(
  doc: OpenDocument,
  name: string,
  op: Record<string, unknown>,
): Record<string, unknown> {
  const { op: _op, ...input } = op
  if (name === 'insert_chart' && input.afterBlockIndex === undefined) {
    input.afterBlockIndex = doc.editor.state.doc.childCount - 1
  }
  return input
}

const CLI_COMMENT_AUTHOR = 'Assistente de IA'

const commentDate = (): string => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')

function commentsAccess(doc: OpenDocument): AiCommentsAccess {
  const { side } = doc
  return {
    list: () => side.comments,
    add: (range, text, meta) => {
      const id = doc.mods.comments.nextCommentId(side.comments)
      if (!doc.mods.comments.addCommentToRange(doc.editor, range.from, range.to, id)) return null
      side.comments.push({
        id,
        author: meta.author ?? CLI_COMMENT_AUTHOR,
        ...(meta.initials ? { initials: meta.initials } : {}),
        date: commentDate(),
        text,
      })
      side.commentsDirty = true
      return id
    },
    remove: (id) => {
      const victims = new Set([
        id,
        ...side.comments.filter((c) => c.parentId === id).map((c) => c.id),
      ])
      if (!side.comments.some((c) => c.id === id)) return false
      for (const v of victims) doc.mods.comments.removeCommentFromDoc(doc.editor, v)
      side.comments = side.comments.filter((c) => !victims.has(c.id))
      side.commentsDirty = true
      return true
    },
    reply: (parentId, text) => {
      const id = doc.mods.comments.nextCommentId(side.comments)
      if (!doc.mods.comments.addReplyToCommentRange(doc.editor, parentId, id)) return false
      side.comments.push({ id, author: CLI_COMMENT_AUTHOR, date: commentDate(), text, parentId })
      side.commentsDirty = true
      return true
    },
    resolve: (id) => {
      if (!side.comments.some((c) => c.id === id)) return false
      side.comments = side.comments.map((c) =>
        c.id === id || c.parentId === id ? { ...c, done: true } : c,
      )
      side.commentsDirty = true
      return true
    },
  }
}

function notesAccess(doc: OpenDocument): AiNotesAccess {
  const { side } = doc
  const listOf = (kind: NoteKind) => (kind === 'footnote' ? side.footnotes : side.endnotes)
  const setList = (kind: NoteKind, next: NoteInfo[]) => {
    if (kind === 'footnote') side.footnotes = next
    else side.endnotes = next
    side.notesDirty = true
  }
  return {
    list: listOf,
    add: (kind, text) => {
      const id = nextNoteId(listOf(kind))
      setList(kind, [...listOf(kind), { id, text }])
      return id
    },
    remove: (kind, id) => {
      if (!listOf(kind).some((n) => n.id === id)) return false
      setList(
        kind,
        listOf(kind).filter((n) => n.id !== id),
      )
      return true
    },
    replace: (kind, id, next) => {
      if (!listOf(kind).some((n) => n.id === id)) return false
      setList(
        kind,
        listOf(kind).map((n) => (n.id === id ? next : n)),
      )
      return true
    },
    protectedMarkBlock: (kind, id) =>
      doc.mods.noteOps.protectedNoteMarkBlock(
        doc.editor.state.doc,
        (block) => protectedXml(doc, block),
        kind,
        id,
      ),
  }
}

export interface NoteSummary {
  kind: NoteKind
  id: string
  /** display number by document order of the reference marks; absent when the text has no mark */
  num?: number
  blockIndex?: number
  text: string
}

/** Footnotes and endnotes with the block holding their reference mark; ids as delete_note takes them. */
export function listNotes(doc: OpenDocument): NoteSummary[] {
  const anchors = doc.mods.noteOps.noteAnchors(doc.editor.state.doc)
  // marks inside protected blocks are only in the block XML; the parse numbered those
  const inXml = new Map<string, number>()
  doc.editor.state.doc.forEach((block, _offset, blockIndex) => {
    if (block.type.name !== 'docProtected') return
    const xml = protectedXml(doc, block)
    for (const m of xml.matchAll(/<w:(footnote|endnote)Reference\b[^>]*\bw:id="([^"]+)"/g)) {
      inXml.set(`${m[1]}:${m[2]}`, blockIndex)
    }
  })
  const parsedNumbers = doc.parsed.noteNumbers as Record<string, number> | undefined
  const summarize = (kind: NoteKind, notes: NoteInfo[]) =>
    notes.map((n) => {
      const key = `${kind}:${n.id}`
      const a = anchors.get(key)
      const xmlBlock = inXml.get(key)
      const num = a?.num ?? parsedNumbers?.[key]
      const blockIndex = a?.blockIndex ?? xmlBlock
      return {
        kind,
        id: n.id,
        ...(num !== undefined ? { num } : {}),
        ...(blockIndex !== undefined ? { blockIndex } : {}),
        text: n.text,
      }
    })
  return [...summarize('footnote', doc.side.footnotes), ...summarize('endnote', doc.side.endnotes)]
}

export interface FieldSummary {
  blockIndex: number
  /** field keyword: SEQ, REF, DATE, PAGE, TOC … */
  type: string
  /** full instruction text */
  instr: string
  /** cached result as the document shows it */
  result: string
  /** Word recomputes the field when the file opens */
  dirty: boolean
  /** the field sits in a protected block (only Word or updateFields-free rebuilds change it) */
  protected?: true
}

const keywordOf = (instr: string) => /^\s*([A-Za-z]+)/.exec(instr)?.[1]?.toUpperCase() ?? ''

/** complex (fldChar) fields at depth 1 and fldSimple fields in a block's XML */
function xmlFields(xml: string): Array<Pick<FieldSummary, 'instr' | 'result' | 'dirty'>> {
  const out: Array<Pick<FieldSummary, 'instr' | 'result' | 'dirty'>> = []
  const decode = (t: string) =>
    t
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
  let depth = 0
  let instr = ''
  let result: string | null = null
  let dirty = false
  const simple: Array<{ instr: string; dirty: boolean; result: string }> = []
  const re =
    /<w:fldChar\b([^>]*)\/?>|<w:instrText\b[^>]*>([\s\S]*?)<\/w:instrText>|<w:t\b[^>]*>([^<]*)<\/w:t>|<w:fldSimple\b([^>]*)>|<\/w:fldSimple>/g
  for (const m of xml.matchAll(re)) {
    if (m[1] !== undefined) {
      const type = /w:fldCharType="(\w+)"/.exec(m[1])?.[1]
      if (type === 'begin') {
        depth++
        if (depth === 1) {
          instr = ''
          result = null
          dirty = /\bw:dirty="(?:true|1)"/.test(m[1])
        }
      } else if (type === 'separate') {
        if (depth === 1) result = ''
      } else if (type === 'end') {
        if (depth === 1)
          out.push({ instr: decode(instr).trim(), result: decode(result ?? ''), dirty })
        depth = Math.max(0, depth - 1)
      }
    } else if (m[2] !== undefined) {
      if (depth === 1 && result === null) instr += m[2]
    } else if (m[3] !== undefined) {
      if (depth === 1 && result !== null) result += m[3]
      else if (depth === 0 && simple.length) simple[simple.length - 1]!.result += m[3]
    } else if (m[4] !== undefined) {
      simple.push({
        instr: decode(/\bw:instr="([^"]*)"/.exec(m[4])?.[1] ?? '').trim(),
        dirty: /\bw:dirty="(?:true|1)"/.test(m[4]),
        result: '',
      })
    } else {
      const done = simple.pop()
      if (done) out.push({ instr: done.instr, result: decode(done.result), dirty: done.dirty })
    }
  }
  return out
}

function protectedXml(doc: OpenDocument, block: Editor['state']['doc']): string {
  if (typeof block.attrs.genXml === 'string') return block.attrs.genXml
  const docxIndex = block.attrs.docxIndex
  return (typeof docxIndex === 'number' ? doc.parsed.blocks[docxIndex]?.originalXml : null) ?? ''
}

/** Every field with its block: editable inline fields from the editor, the rest from the block XML. */
export function listFields(doc: OpenDocument): FieldSummary[] {
  const out: FieldSummary[] = []
  doc.editor.state.doc.forEach((block, _offset, blockIndex) => {
    if (block.type.name === 'docProtected') {
      for (const f of xmlFields(protectedXml(doc, block))) {
        out.push({ blockIndex, type: keywordOf(f.instr), ...f, protected: true })
      }
      return
    }
    // a result split by formatting is one field: join text nodes that touch and carry the same mark
    type FieldMark = Editor['state']['doc']['marks'][number]
    let prev: { mark: FieldMark; end: number; field: FieldSummary } | null = null
    block.descendants((node, pos) => {
      if (!node.isText) return true
      const instr = node.marks.find((m) => m.type.name === 'instrField')
      const ref = node.marks.find((m) => m.type.name === 'refField')
      const mark = instr ?? ref
      if (!mark) return true
      if (prev && prev.end === pos && prev.mark.eq(mark)) {
        prev.field.result += node.text ?? ''
        prev.end += node.nodeSize
        return true
      }
      const field: FieldSummary = instr
        ? {
            blockIndex,
            type: keywordOf(String(instr.attrs.instr)),
            instr: String(instr.attrs.instr),
            result: node.text ?? '',
            dirty: instr.attrs.dirty === true,
          }
        : {
            blockIndex,
            type: 'REF',
            instr: mark.attrs.instr
              ? String(mark.attrs.instr).trim()
              : `REF ${String(mark.attrs.name)} \\h`,
            result: node.text ?? '',
            dirty: mark.attrs.dirty === true,
          }
      out.push(field)
      prev = { mark, end: pos + node.nodeSize, field }
      return true
    })
  })
  return out
}

const slotOf = (kind: 'header' | 'footer', view: HfView): HfSlot =>
  view === 'default' ? kind : `${kind}${view === 'first' ? 'First' : 'Even'}`

export interface HeaderFooterSectionState {
  index: number
  firstBlock: number
  lastBlock: number
  header: string
  footer: string
  headerFirst: string | null
  footerFirst: string | null
  headerEven: string | null
  footerEven: string | null
}

export interface HeaderFooterState {
  header: string
  footer: string
  headerFirst: string | null
  footerFirst: string | null
  headerEven: string | null
  footerEven: string | null
  titlePg: boolean
  evenOddHf: boolean
  multiSection: boolean
  /**
   * Per-section text, indexed the way set_header_footer{section} addresses sections
   * (0-based, the order `docs read --sections` lists). An edit that targeted one
   * section is visible here, so read and write agree.
   */
  sections: HeaderFooterSectionState[]
  /** the default header's watermark: text, picture, or none; pending set_watermark results included */
  watermark: WatermarkState
}

export type WatermarkState =
  | { kind: 'none'; text: null }
  | { kind: 'text'; text: string }
  | { kind: 'picture'; text: null; widthPt: number; heightPt: number; washout: boolean }

function watermarkState(doc: OpenDocument): WatermarkState {
  const { side, parsed } = doc
  if (side.watermark !== undefined) {
    const wm = side.watermark
    if (wm === null) return { kind: 'none', text: null }
    if ('image' in wm) {
      // the frame is sized against the page on save; report the natural size until then
      const box = { widthPt: (wm.image.widthPx * 72) / 96, heightPt: (wm.image.heightPx * 72) / 96 }
      const f = wm.scale === undefined ? 1 : wm.scale / 100
      return {
        kind: 'picture',
        text: null,
        widthPt: Math.round(box.widthPt * f * 100) / 100,
        heightPt: Math.round(box.heightPt * f * 100) / 100,
        washout: wm.washout !== false,
      }
    }
    return { kind: 'text', text: wm.text }
  }
  if (side.watermarkText) return { kind: 'text', text: side.watermarkText }
  const pic = parsed.watermarkPicture
  if (pic)
    return {
      kind: 'picture',
      text: null,
      widthPt: pic.widthPt,
      heightPt: pic.heightPt,
      washout: pic.washout,
    }
  return { kind: 'none', text: null }
}

/** Key of a pending per-section header/footer edit: its break paragraph (docxIndex) + kind + variant. */
function sectionHfKey(lastBlockIndex: number, kind: 'header' | 'footer', variant: HfView): string {
  return variant === 'default'
    ? `${lastBlockIndex}:${kind}`
    : `${lastBlockIndex}:${kind}:${variant}`
}

/**
 * The current value of one section's header/footer: that section's pending edit,
 * the nearest earlier section's part it inherits (Word's "same as previous"), or the
 * document-level strip of a document without header/footer references.
 */
function resolveSectionHf(
  doc: OpenDocument,
  sections: SectionInfo[],
  si: number,
  kind: 'header' | 'footer',
  variant: HfView,
): HeaderFooter | null {
  const { side, parsed } = doc
  const section = sections[si]!
  const own = side.hfSection.get(sectionHfKey(section.lastBlockIndex, kind, variant))
  if (own) return own.hf
  // the last section's pending edit lives in the trailing-options slots saveDocument reads
  if (si === sections.length - 1) {
    const slot = slotOf(kind, variant)
    if (side.hfDirty.has(slot)) return side.hf[slot] ?? null
  }
  // walking back from si, the nearest section with a pending part or a parsed reference owns it
  for (let i = si; i >= 0; i--) {
    const pending = side.hfSection.get(sectionHfKey(sections[i]!.lastBlockIndex, kind, variant))
    if (pending) return pending.hf
    const ref = (kind === 'header' ? sections[i]!.headerRefs : sections[i]!.footerRefs)[variant]
    if (!ref) continue
    const info = parsed.hfParts?.[ref]
    return info ? hfFromPart(info) : null
  }
  // Word inherits forward only: with a reference somewhere later, this section shows nothing
  const anyRef = sections.some(
    (s) => (kind === 'header' ? s.headerRefs : s.footerRefs)[variant] !== undefined,
  )
  return anyRef ? null : (side.hf[slotOf(kind, variant)] ?? null)
}

function sectionHfText(
  doc: OpenDocument,
  sections: SectionInfo[],
  si: number,
  kind: 'header' | 'footer',
  variant: HfView,
): string {
  const value = resolveSectionHf(doc, sections, si, kind, variant)
  return value ? doc.mods.hfText.hfEditText(value) : ''
}

export function headerFooterState(doc: OpenDocument): HeaderFooterState {
  const { side } = doc
  const textOf = (slot: HfSlot) => {
    const v = side.hf[slot]
    return v ? doc.mods.hfText.hfEditText(v) : ''
  }
  const sections = readSections(doc.parsed)
  return {
    header: textOf('header'),
    footer: textOf('footer'),
    headerFirst: side.titlePg ? textOf('headerFirst') : null,
    footerFirst: side.titlePg ? textOf('footerFirst') : null,
    headerEven: side.evenOddHf ? textOf('headerEven') : null,
    footerEven: side.evenOddHf ? textOf('footerEven') : null,
    titlePg: side.titlePg,
    evenOddHf: side.evenOddHf,
    multiSection: sections.length > 1,
    sections: sections.map((section, index) => ({
      index,
      firstBlock: section.firstBlockIndex,
      lastBlock: section.lastBlockIndex,
      header: sectionHfText(doc, sections, index, 'header', 'default'),
      footer: sectionHfText(doc, sections, index, 'footer', 'default'),
      headerFirst: section.titlePg ? sectionHfText(doc, sections, index, 'header', 'first') : null,
      footerFirst: section.titlePg ? sectionHfText(doc, sections, index, 'footer', 'first') : null,
      headerEven: side.evenOddHf ? sectionHfText(doc, sections, index, 'header', 'even') : null,
      footerEven: side.evenOddHf ? sectionHfText(doc, sections, index, 'footer', 'even') : null,
    })),
    watermark: watermarkState(doc),
  }
}

/** The default variant is written to the trailing section's part, as the app does for single-section documents. */
function headerFooterAccess(doc: OpenDocument): {
  read(): HeaderFooterState
  set(kind: 'header' | 'footer', view: HfView, text: string): string | null
} {
  const { side } = doc
  return {
    read: () => headerFooterState(doc),
    set: (kind, view, text) => {
      if (view === 'first' && !side.titlePg) {
        side.titlePg = true
        side.titlePgDirty = true
      }
      if (view === 'even' && !side.evenOddHf) {
        side.evenOddHf = true
        side.evenOddHfDirty = true
      }
      const slot = slotOf(kind, view)
      side.hf[slot] = doc.mods.hfText.applyHfText(side.hf[slot] ?? null, text)
      side.hfDirty.add(slot)
      return null
    },
  }
}

/**
 * set_header_footer with an explicit `section` (0-based, the numbering `docs read
 * --sections` prints). The edit is routed the way the app routes it: the trailing
 * section keeps using the save's trailing-options slots, every other section goes
 * through SaveOptions.sectionHf keyed by its break paragraph.
 */
function setSectionHeaderFooter(
  doc: OpenDocument,
  op: Record<string, unknown>,
): { output?: string; error?: string } {
  const kind = op.kind
  if (kind !== 'header' && kind !== 'footer') return { error: 'kind deve ser "header" ou "footer"' }
  const view = op.view === undefined ? 'default' : op.view
  if (view !== 'default' && view !== 'first' && view !== 'even')
    return { error: 'view deve ser "default", "first" ou "even"' }
  if (typeof op.text !== 'string') return { error: 'text deve ser uma string' }
  // a doubly escaped newline arrives as the two characters backslash-n
  const text = op.text.includes('\n') ? op.text : op.text.replace(/\\n/g, '\n')
  if (text.length > 2000)
    return { error: 'text é longo demais para um cabeçalho/rodapé (máximo de 2000 caracteres)' }
  const sections = readSections(doc.parsed)
  if (sections.length === 0) return { error: 'o documento ainda não tem seções' }
  const section = op.section
  if (
    typeof section !== 'number' ||
    !Number.isInteger(section) ||
    section < 0 ||
    section >= sections.length
  )
    return {
      error: `section deve ser um inteiro entre 0 e ${sections.length - 1} (o documento tem ${sections.length} seção${sections.length === 1 ? '' : 's'})`,
    }
  const { side } = doc
  const pageSetup = pageSetupAccess(doc)
  // a section break inserted earlier in this same batch exists only in the editor:
  // the engine's per-section save routing addresses parsed sections, so the indexes
  // would not line up
  if (pageSetup.list().length !== sections.length)
    return {
      error:
        'set_header_footer{section} se refere às seções salvas; aplique a quebra de seção em um `genoffice docs apply` separado primeiro e depois edite o cabeçalho/rodapé',
    }
  if (view === 'even' && !side.evenOddHf) {
    side.evenOddHf = true
    side.evenOddHfDirty = true
  }
  const next = doc.mods.hfText.applyHfText(
    resolveSectionHf(doc, sections, section, kind, view),
    text,
  )
  const slot = slotOf(kind, view)
  if (section === sections.length - 1) {
    if (view === 'first' && !sections[section]!.titlePg) {
      side.titlePg = true
      side.titlePgDirty = true
    }
    side.hf[slot] = next
    side.hfDirty.add(slot)
  } else {
    if (view === 'first' && !sections[section]!.titlePg) {
      // a first-page part is only used by a section carrying <w:titlePg/>
      const current = pageSetup.current(section)
      if (!current) return { error: `a seção ${section} não existe` }
      const failure = pageSetup.set(section, { settings: current.settings, titlePg: true })
      if (failure) return { error: failure }
    }
    side.hfSection.set(sectionHfKey(sections[section]!.lastBlockIndex, kind, view), {
      lastBlockIndex: sections[section]!.lastBlockIndex,
      kind,
      variant: view,
      hf: next,
    })
  }
  const where = sections.length > 1 ? ` da seção ${section} (base 0)` : ''
  return {
    output: `Atualizado o ${kind}${view !== 'default' ? ` (variante de página ${view})` : ''}${where}.`,
  }
}

const IMAGE_MIMES = new Set(['image/png', 'image/jpeg', 'image/gif'])

/**
 * The app's insert_image / insert_picture download through the renderer
 * bridge and measure with an <img>; neither exists here, so the node is built
 * from the bytes directly, then the executor is told the document is still
 * the one it saw. `picture` = the sized / floating insert_picture variant.
 */
async function insertImage(
  doc: OpenDocument,
  op: Record<string, unknown>,
  opts: ApplyOptions,
  picture = false,
): Promise<{ output?: string; error?: string }> {
  const url = typeof op.url === 'string' ? op.url.trim() : ''
  if (!url) return { error: 'url deve ser um caminho local, uma URL data: ou uma URL http(s)' }
  if (!opts.ctx) return { error: 'fontes de imagem precisam de um contexto de caminho' }
  const source = await readImageSource(url, opts.ctx, opts.baseDir)
  if (!source) return { error: `imagem não encontrada ou não disponível para download: ${url}` }
  const size = imageSize(source.bytes)
  if (!size || !source.mime || !IMAGE_MIMES.has(source.mime)) {
    return { error: 'formato de imagem não compatível (somente png, jpg e gif podem ser incorporados)' }
  }
  const at = doc.mods.floating.insertPosition(doc.editor, op.afterBlockIndex)
  if ('error' in at) return { error: at.error }
  const base64 = Buffer.from(source.bytes).toString('base64')
  const mime = source.mime as 'image/png' | 'image/jpeg' | 'image/gif'
  let float: FloatSpec | undefined
  if (picture && op.float !== undefined) {
    const r = doc.mods.floating.resolveFloat(op.float)
    if ('error' in r) return { error: r.error }
    float = r.float
  }
  const built = doc.mods.floating.pictureNode({
    base64,
    mime,
    naturalWidth: size.width,
    naturalHeight: size.height,
    label: picture ? 'Figura' : 'Imagem',
    ...(picture
      ? {
          width: op.width,
          height: op.height,
          float,
          ...(typeof op.altText === 'string' ? { altText: op.altText } : {}),
        }
      : { width: `${Math.min(size.width, Number(op.maxWidthPx) || 480)}px` }),
  })
  if ('error' in built) return { error: built.error }
  doc.editor.chain().insertContentAt(at.pos, built.node).run()
  doc.mods.tools.markDocSeen(doc.editor)
  return {
    output: `Inserida a ${float ? 'flutuante ' : ''}${picture ? 'figura' : 'imagem'} (${built.widthPx}x${built.heightPx}px) após o bloco ${at.after}.`,
  }
}

/** set_watermark with an image: the app downloads through the renderer bridge; here the bytes come from a path / data: / http(s) URL. */
async function setPictureWatermark(
  doc: OpenDocument,
  op: Record<string, unknown>,
  opts: ApplyOptions,
): Promise<{ output?: string; error?: string }> {
  const url = String(op.image).trim()
  if (!url) return { error: 'image deve ser um caminho local, uma URL data: ou uma URL http(s)' }
  if (!opts.ctx) return { error: 'fontes de imagem precisam de um contexto de caminho' }
  const source = await readImageSource(url, opts.ctx, opts.baseDir)
  if (!source) return { error: `imagem não encontrada ou não disponível para download: ${url}` }
  const size = imageSize(source.bytes)
  if (!size || !source.mime || !IMAGE_MIMES.has(source.mime)) {
    return { error: 'formato de imagem não compatível (somente png, jpg e gif podem ser incorporados)' }
  }
  const resolved = doc.mods.floating.resolvePictureWatermark(op, {
    base64: Buffer.from(source.bytes).toString('base64'),
    mime: source.mime as 'image/png' | 'image/jpeg' | 'image/gif',
    widthPx: size.width,
    heightPx: size.height,
  })
  if ('error' in resolved) return { error: resolved.error }
  doc.side.watermark = resolved.spec
  return { output: `Marca d'água de imagem definida (origem de ${size.width}x${size.height}px).` }
}

type DocExtras = NonNullable<Parameters<Modules['tools']['executeTool']>[10]>

/** styles.xml catalog plus this batch's define_style entries, and the header watermark. */
function docExtras(doc: OpenDocument): DocExtras {
  return {
    styles: {
      list: () => listStyles(doc),
      upsert: (up) => {
        const prev = doc.side.styleUpserts.get(up.styleId)
        doc.side.styleUpserts.set(
          up.styleId,
          prev
            ? {
                ...prev,
                ...up,
                pPr: up.pPr || prev.pPr ? { ...prev.pPr, ...up.pPr } : undefined,
                rPr: up.rPr || prev.rPr ? { ...prev.rPr, ...up.rPr } : undefined,
              }
            : up,
        )
        return null
      },
    },
    watermark: {
      current: () => watermarkState(doc).text,
      set: (spec) => {
        doc.side.watermark = spec
        return null
      },
    },
  }
}

export interface StyleSummary {
  styleId: string
  name: string
  type: 'paragraph' | 'character' | 'table'
  basedOn?: string
  headingLevel?: number
  /** blocks (paragraphs, headings, list items, runs) carrying the style */
  inUse: number
  /** defined or changed by this batch, not saved yet */
  pending?: boolean
}

/** Every style the document defines (plus pending define_style entries), with usage counts. */
export function listStyles(doc: OpenDocument): StyleSummary[] {
  const counts = new Map<string, number>()
  const headings = new Map<number, number>()
  const bump = (id: string) => counts.set(id, (counts.get(id) ?? 0) + 1)
  doc.editor.state.doc.descendants((node) => {
    if (typeof node.attrs?.styleId === 'string' && node.attrs.styleId) bump(node.attrs.styleId)
    else if (node.type.name === 'docHeading') {
      const level = Number(node.attrs.level) || 1
      headings.set(level, (headings.get(level) ?? 0) + 1)
    }
    for (const mark of node.marks) {
      if (typeof mark.attrs?.styleId === 'string' && mark.attrs.styleId) bump(mark.attrs.styleId)
    }
    return true
  })
  const out = new Map<string, StyleSummary>()
  for (const s of doc.parsed.styles.values()) {
    if (s.linkedCharShell) continue
    let inUse = counts.get(s.styleId) ?? 0
    if (s.headingLevel && /^Heading[1-9]$/.test(s.styleId))
      inUse += headings.get(s.headingLevel) ?? 0
    out.set(s.styleId, {
      styleId: s.styleId,
      name: s.name,
      type: s.type,
      ...(s.basedOn ? { basedOn: s.basedOn } : {}),
      ...(s.headingLevel ? { headingLevel: s.headingLevel } : {}),
      inUse,
    })
  }
  const parsed = (id: string) => doc.parsed.styles.get(id)
  for (const up of doc.side.styleUpserts.values()) {
    const cur = out.get(up.styleId)
    const basedOn = up.basedOn === undefined ? cur?.basedOn : (up.basedOn ?? undefined)
    const headingLevel = pendingHeadingLevel(
      up.styleId,
      (id) => doc.side.styleUpserts.get(id),
      parsed,
    )
    out.set(up.styleId, {
      styleId: up.styleId,
      name: up.name ?? cur?.name ?? up.styleId,
      type: cur?.type ?? up.type ?? 'paragraph',
      ...(basedOn ? { basedOn } : {}),
      ...(headingLevel ? { headingLevel } : {}),
      inUse: cur?.inUse ?? counts.get(up.styleId) ?? 0,
      pending: true,
    })
  }
  return [...out.values()]
}

export interface CommentSummary extends CommentInfo {
  blockIndex?: number
  anchorText?: string
}

/** Every comment with the block its anchor sits in, ids as reply_comment / resolve_comment / delete_comment take them. */
export function listComments(doc: OpenDocument): CommentSummary[] {
  const anchors = doc.mods.protocol.commentAnchors(doc.editor)
  return doc.side.comments.map((c) => {
    const a = anchors.get(c.parentId ?? c.id)
    const base = { ...c, done: c.done === true }
    return a ? { ...base, blockIndex: a.blockIndex, anchorText: a.excerpt } : base
  })
}

export interface RevisionSummary {
  /** positional: r1 = first pending change in document order; renumbered after every edit */
  id: string
  type: 'insertion' | 'deletion' | 'formatting' | 'move'
  kind: string
  author: string
  date?: string
  blockIndex: number
  text: string
  /** formatting changes: `field: old → new` pairs */
  change?: string
}

/** Pending tracked changes in document order, ids as accept_changes / reject_changes take them. */
export function listRevisions(doc: OpenDocument): RevisionSummary[] {
  return doc.mods.revisionOps
    .listRevisionEntries(doc.editor.state.doc)
    .map(({ from: _from, to: _to, ...rest }) => rest)
}

const CONTENT_OPS = [
  'insert_content',
  'replace_blocks',
  'insert_image',
  'insert_picture',
  'insert_text_box',
  'insert_chart',
  'edit_chart',
  'set_header_footer',
  'set_page_setup',
  'insert_section_break',
  'set_watermark',
  'define_style',
  'add_comment',
  'reply_comment',
  'resolve_comment',
  'accept_changes',
  'reject_changes',
  'insert_footnote',
  'insert_endnote',
  'edit_note',
  'delete_note',
  'delete_comment',
  'set_properties',
] as const

const CONTENT_NOTES: Record<(typeof CONTENT_OPS)[number], string> = {
  insert_content: 'novos blocos após um índice (-1 = início do documento; omitido = fim do documento)',
  replace_blocks:
    'reescreve um intervalo de blocos; os novos blocos herdam a formatação de parágrafo e de texto dos blocos substituídos',
  insert_image:
    'url é um caminho local (absoluto, ou relativo ao diretório atual e depois ao arquivo de ops), uma URL data: ou uma URL http(s); png/jpg/gif; reduzido ao limite de maxWidthPx; anexado ao final quando afterBlockIndex é omitido',
  insert_picture:
    'variante dimensionada/flutuante de insert_image (mesmas formas de url); width/height e deslocamentos de float aceitam "2.54cm", "1in", "72pt", "96px" ou pontos; float.anchor paragraph|page|margin, float.wrap square|tight|topAndBottom|behind|inFront',
  insert_text_box:
    'caixa de texto flutuante ancorada após um bloco; x/y/width/height em cm/in/pt/px; âncora paragraph (padrão) ou page; wrap topAndBottom (o texto continua abaixo) ou none',
  set_watermark:
    'marca de água atrás de todas as páginas, gravada no cabeçalho (Word Design > Watermark): texto cinza (color "#RRGGBB", opacity 0-1, diagonal true/false) ou uma imagem (image = caminho local / data: / URL http(s); scale em porcentagem, por padrão cabe nas margens; washout true/false); text null ou image null remove a marca de água',
  define_style:
    'cria ou ajusta um estilo em styles.xml; somente os campos informados mudam; aplique-o a parágrafos com a op de bloco { op: "applyStyle", target, styleId }; `genoffice docs read --styles` lista os ids',
  insert_chart:
    'gráfico nativo do Word; os valores por série correspondem às categorias; anexado ao final quando afterBlockIndex é omitido',
  edit_chart:
    'altera os dados de um bloco de gráfico (`genoffice docs read` lista blocos com kind "chart"); as contagens devem corresponder às originais, null mantém uma posição',
  set_header_footer:
    'texto simples, \\n entre linhas, {PAGE} / {NUMPAGES} viram campos de número de página, "" limpa; view first/even ativa a configuração de primeira página diferente / par-ímpar; section escolhe uma seção (base 0, `genoffice docs read --sections`; omitido = a última seção); `genoffice docs read --header-footer` mostra o texto de cada seção',
  set_page_setup:
    'papel (A3/A4/A5/B5/Letter/Legal/Tabloid) ou width/height, orientação, margens {top,right,bottom,left,header,footer,gutter}, colunas {count,spacing}, titlePg, pageNumberStart/pageNumberFormat; comprimentos como "2.54cm" / "1in" / "72pt" ou twips; section ou blockIndex escolhe uma seção, nenhum dos dois = todas; `genoffice docs read --sections` as lista',
  insert_section_break:
    'os blocos após afterBlockIndex passam a formar uma nova seção (uma cópia da configuração atual, que depois pode ser alterada com set_page_setup section: N); tipo nextPage (padrão) / continuous / evenPage / oddPage',
  add_comment:
    'nova conversa em um bloco, ou em um trecho exato de texto dentro dele (occurrence escolhe uma entre várias correspondências); o texto do documento não é alterado; o autor padrão é "Assistente de IA"',
  reply_comment: 'ids de `genoffice docs read --comments`; respostas são anexadas à raiz da conversa',
  resolve_comment: 'ids de `genoffice docs read --comments`',
  accept_changes:
    'seletor: all: true, ids de `genoffice docs read --revisions` (posicionais, releia após cada edição), ou author / type / blockIndex / blockRange / before; inserções viram texto comum, o texto excluído sai, o texto movido permanece no destino; all / author / type / before também abrangem cabeçalhos, rodapés, notas de rodapé e notas de fim',
  reject_changes:
    'mesmo seletor de accept_changes; o texto inserido sai, o texto excluído volta e a formatação é revertida',
  insert_footnote:
    'uma marca de referência sobrescrita entra em blockIndex (logo após afterText, senão no fim do bloco) e o texto da nota aparece no rodapé da página; `genoffice docs read --notes` lista as notas',
  insert_endnote: 'como insert_footnote, mas o texto da nota é reunido no fim do documento',
  edit_note:
    'findReplace dentro de uma nota de rodapé ou de fim (id de `genoffice docs read --notes`; kind só quando uma nota de rodapé e uma de fim compartilham o id): toda ocorrência de find vira replace, a nota mantém seu id, marca de referência e formatação; matchCase é true por padrão',
  delete_note:
    'kind footnote|endnote e um id de `genoffice docs read --notes`; remove também a marca de referência',
  delete_comment:
    'remove o comentário e sua âncora; uma raiz de conversa com respostas exige withReplies: true; um id de resposta remove apenas essa resposta',
  set_properties:
    'as propriedades do documento do arquivo: core (title, subject, creator, keywords, description, lastModifiedBy, revision, created, modified, category, contentStatus, identifier, language, version), app (template, totalTime, pages, words, characters, application, docSecurity, lines, paragraphs, scaleCrop, company, linksUpToDate, charactersWithSpaces, sharedDoc, hyperlinksChanged, appVersion, manager) e custom (por nome); uma string define um valor, um número ou booleano cria uma propriedade personalizada tipada (i4/bool; uma propriedade i4/i8/r8/bool/filetime/date existente mantém seu tipo vt quando o valor cabe), null remove a propriedade, clear: "all" esvazia as três partes; `genoffice docs read --properties` lista as atuais com seus customTypes',
}

const AFTER_BLOCK_INDEX =
  'insere após este índice de bloco; -1 = início do documento; omitido = fim do documento'

const INSERT_IMAGE_SCHEMA: JsonSchema = {
  type: 'object',
  properties: {
    url: { type: 'string', description: 'caminho local, URL data: ou URL http(s) de um png/jpg/gif' },
    maxWidthPx: { type: 'integer', description: 'largura máxima em px, padrão 480' },
    afterBlockIndex: { type: 'integer', description: AFTER_BLOCK_INDEX },
  },
  required: ['url'],
}

/** set_properties: the text values of a core/app property, by key (a string sets, null removes) */
const PROPERTY_VALUES: JsonSchema = {
  type: 'object',
  additionalProperties: { type: ['string', 'null'] },
}

/** set_properties.custom: text stays lpwstr, a number/boolean keeps or takes a typed vt value */
const CUSTOM_PROPERTY_VALUES: JsonSchema = {
  type: 'object',
  additionalProperties: { type: ['string', 'number', 'boolean', 'null'] },
}

const SET_PROPERTIES_SCHEMA: JsonSchema = {
  type: 'object',
  properties: {
    core: {
      ...PROPERTY_VALUES,
      description:
        'propriedades core por nome amigável (title, subject, creator, keywords, description, lastModifiedBy, …); uma string define o valor, null remove o elemento',
    },
    app: {
      ...PROPERTY_VALUES,
      description:
        'propriedades estendidas (app) por nome amigável (company, template, application, manager, …); uma string define o valor, null remove o elemento',
    },
    custom: {
      ...CUSTOM_PROPERTY_VALUES,
      description:
        'propriedades personalizadas por nome (MatterId, Reviewer, …): uma string grava vt:lpwstr, um número vt:i4 (ou i8/r8 quando fora de int32), um booleano vt:bool; uma propriedade tipada existente mantém seu tipo vt quando o valor cabe nele (i4/i8/r8/ui*/bool/filetime/date); null remove a propriedade',
    },
    clear: {
      type: 'string',
      enum: ['all'],
      description: 'remove todas as propriedades das três partes',
    },
  },
}

/** insert_picture / set_watermark headless: the in-app schema, with the image source widened to local paths */
function insertPictureSchema(tool: JsonSchema, key: 'url' | 'image' = 'url'): JsonSchema {
  const copy = cliToolSchema(tool)
  const src = copy.properties?.[key]
  if (src)
    src.description =
      key === 'url'
        ? 'caminho local, URL data: ou URL http(s) de um png/jpg/gif'
        : 'marca de água de imagem: caminho local, URL data: ou URL http(s) de um png/jpg/gif, centralizada atrás do corpo; null remove a marca de água'
  return copy
}

/** The in-app tool schemas describe the cursor default; headless, an omitted position appends. */
function cliToolSchema(schema: JsonSchema): JsonSchema {
  const copy = JSON.parse(JSON.stringify(schema)) as JsonSchema
  const after = copy.properties?.afterBlockIndex
  if (after) after.description = AFTER_BLOCK_INDEX
  return copy
}

/** The in-app toolSchema has no section selector; headless, one call can target any section (see setSectionHeaderFooter). */
function headerFooterSchema(schema: JsonSchema): JsonSchema {
  const copy = cliToolSchema(schema)
  copy.properties = {
    ...copy.properties,
    section: {
      type: 'integer',
      description:
        'grava esta seção em vez da última: índice a partir de 0 de `genoffice docs read --sections`; uma quebra de seção inserida no mesmo lote precisa ser salva primeiro',
    },
  }
  return copy
}

/** The engine signatures read `{ op: "name", fields }  // note`; the catalog keeps the two apart. */
function splitSignature(signature: string): { signature: string; note?: string } {
  const [sig, note] = signature.split(/\s+\/\/\s+/, 2)
  const fields = sig!.replace(/^\{\s*op:\s*"[^"]+",\s*/, '{ ')
  return note ? { signature: fields, note } : { signature: fields }
}

export async function docsCatalog(): Promise<{ catalog: OpCatalog; htmlRules: string }> {
  const mods = await docsModules()
  const block: OpEntry[] = mods.ops.opCatalog().map((d) => ({
    op: d.name,
    group: 'block',
    ...splitSignature(d.signature),
    target: d.target,
    keys: [...d.keys],
  }))
  const tools = new Map(mods.tools.AGENT_TOOLS.map((t) => [t.name, t]))
  const content: OpEntry[] = CONTENT_OPS.map((op) => {
    const tool = tools.get(op)
    if (!tool && op !== 'insert_image' && op !== 'set_properties') {
      throw new Error(`docs op catalog: no tool named ${op}`)
    }
    const schema =
      op === 'insert_image'
        ? INSERT_IMAGE_SCHEMA
        : op === 'set_properties'
          ? SET_PROPERTIES_SCHEMA
          : op === 'insert_picture'
            ? insertPictureSchema(tool!.inputSchema as JsonSchema)
            : op === 'set_watermark'
              ? insertPictureSchema(tool!.inputSchema as JsonSchema, 'image')
              : op === 'set_header_footer'
                ? headerFooterSchema(tool!.inputSchema as JsonSchema)
                : cliToolSchema(tool!.inputSchema as JsonSchema)
    return {
      op,
      group: 'content',
      signature: signatureFromSchema(schema, []),
      note: CONTENT_NOTES[op],
      schema,
    }
  })
  const controls: OpEntry[] = CONTENT_CONTROL_OPS.map((op) => ({
    op,
    group: 'content-controls',
    signature: signatureFromSchema(CONTENT_CONTROL_SCHEMAS[op]),
    note: CONTENT_CONTROL_NOTES[op],
    schema: CONTENT_CONTROL_SCHEMAS[op],
  }))
  const packages: OpEntry[] = PACKAGE_OPS.map((op) => ({
    op,
    group: 'package',
    signature: signatureFromSchema(PACKAGE_SCHEMAS[op]),
    note: PACKAGE_NOTES[op],
    schema: PACKAGE_SCHEMAS[op],
  }))
  const groups = [
    {
      name: 'block',
      summary: 'edições de blocos existentes, endereçadas por alvo',
      ops: block.map((o) => o.op),
    },
    {
      name: 'content',
      summary:
        'novo conteúdo, figuras, caixas de texto, gráficos, cabeçalho/rodapé, configuração de página, marca de água, estilos, comentários, alterações registradas, notas de rodapé',
      ops: [...CONTENT_OPS],
    },
    {
      name: 'content-controls',
      summary: 'controles de conteúdo de nível de bloco (w:sdt): agrupa blocos em um controle ou os desembrulha',
      ops: [...CONTENT_CONTROL_OPS],
    },
    {
      name: 'package',
      summary:
        'partes do pacote: remove partes custom XML / de extensões web que contêm metadados (limpeza antes do arquivamento)',
      ops: [...PACKAGE_OPS],
    },
  ]
  return {
    catalog: withFingerprint({
      domain: 'docs',
      groups,
      ops: [...block, ...content, ...controls, ...packages],
    }),
    htmlRules: mods.protocol.HTML_RULES,
  }
}
