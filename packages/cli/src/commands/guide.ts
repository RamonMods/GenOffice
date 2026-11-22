import { flagBool } from '../args'
import type { CommandDef } from '../registry'
import { docsCatalog } from '../formats/docx'
import { sheetCatalog, sheetGuideText } from '../formats/xlsx-catalog'
import {
  opLines,
  renderGroups,
  withFingerprint,
  type GuideDomain,
  type OpCatalog,
  type OpEntry,
} from '../op-catalog'
import { CliError, EXIT, type CommandResult } from '../result'
import { didYouMean } from '../suggest'

const DOMAINS: readonly GuideDomain[] = ['slides', 'docs', 'sheets']

/**
 * Every op reference is generated from the definition the executor validates against (the pptx-ops
 * markdown, the sheets zod schema, the docs op registry and tool schemas), so the guide cannot
 * drift from what `apply` accepts; `--json` returns the same catalog with each op's schema.
 */
export const guideCommand: CommandDef = {
  name: 'guide',
  summary: 'Imprime a referência de ops e os guias de design que um agente precisa antes de escrever ops ou specs.',
  usage: 'guide <slides|docs|sheets> [group|op|design|spec] [--index] [--fingerprint]',
  options: [
    { name: 'index', description: 'assinatura de uma linha de cada op em vez de um guia de grupo' },
    {
      name: 'fingerprint',
      description:
        'apenas o fingerprint do catálogo (muda com qualquer op ou campo); sem domínio = todos os três',
    },
  ],
  async run(args) {
    const [domain, topic] = args.positionals
    const json = flagBool(args, 'json')
    if (flagBool(args, 'fingerprint') && (domain === undefined || isDomain(domain))) {
      return fingerprints(domain, json)
    }
    if (!isDomain(domain)) {
      throw new CliError(
        EXIT.usage,
        'guias disponíveis para: slides, docs, sheets',
        { usage: 'genoffice guide slides | genoffice guide docs | genoffice guide sheets' },
        {
          reason: domain === undefined ? 'missing_argument' : 'invalid_argument',
          suggestion: 'execute `genoffice guide slides|docs|sheets`',
        },
      )
    }
    if (domain === 'slides' && (topic === 'design' || topic === 'spec')) {
      const { SLIDES_GUIDES } = await import('@genoffice/pipelines/slides')
      return { summary: SLIDES_GUIDES[topic].content }
    }
    const { catalog, text } = await load(domain, topic, flagBool(args, 'index'))
    if (!topic) return { summary: text(), ...(json ? { detail: { ...catalog } } : {}) }
    const group = catalog.groups.find((g) => g.name === topic)
    if (group) {
      const entries = catalog.ops.filter((e) => e.group === topic)
      return {
        summary: text(topic),
        ...(json
          ? { detail: { domain, fingerprint: catalog.fingerprint, group, ops: entries } }
          : {}),
      }
    }
    const entry = catalog.ops.find((e) => e.op === topic)
    if (!entry) {
      const names = [...catalog.groups.map((g) => g.name), ...catalog.ops.map((e) => e.op)]
      const guess = didYouMean(topic, names)
      throw new CliError(
        EXIT.usage,
        `grupo ou op desconhecido: ${topic}`,
        { groups: catalog.groups.map((g) => g.name) },
        {
          reason: 'invalid_argument',
          suggestion: guess
            ? `você quis dizer ${guess}?`
            : `execute \`genoffice guide ${domain}\` para ver os grupos e as ops`,
        },
      )
    }
    return { summary: opText(entry), ...(json ? { detail: { ...entry } } : {}) }
  },
}

function isDomain(value: string | undefined): value is GuideDomain {
  return (DOMAINS as readonly string[]).includes(value ?? '')
}

interface Loaded {
  catalog: OpCatalog
  text: (group?: string) => string
}

/** The catalog `guide <domain> --json` prints; the MCP op schemas are built from the same object. */
export async function loadCatalog(domain: GuideDomain): Promise<OpCatalog> {
  return (await load(domain, undefined, false)).catalog
}

async function load(
  domain: GuideDomain,
  topic: string | undefined,
  index: boolean,
): Promise<Loaded> {
  if (domain === 'sheets') return { catalog: sheetCatalog(), text: sheetGuideText }
  if (domain === 'docs') {
    const { catalog, htmlRules } = await docsCatalog()
    return { catalog, text: (group) => docsGuideText(catalog, htmlRules, group) }
  }
  const docs = await import('@genoffice/pptx-ops/op-docs')
  const catalog = slidesCatalog(docs)
  return {
    catalog,
    text: (group) => {
      if (index) return docs.opSignatureIndex()
      if (group) return docs.opGuide(group)!
      return [
        'Grupos de op (genoffice guide slides <group> mostra um; genoffice guide slides <op> mostra uma op):',
        docs.opGuideCatalog(),
        '',
        'Criando uma nova apresentação: `genoffice guide slides design` (o fluxo de trabalho por etapas: folha de estilo, esboço, um arquivo de página por vez, construção, QC) e `genoffice guide slides spec` (o esboço e a spec JSON de página para `genoffice slides check` e `genoffice create --type pptx --spec`).',
        '',
        'Toda op: { "op": "<name>", "target": { "slide": <index|"s_n">, "el"?: "e_*" }, ...fields }.',
        'As unidades são EMU (914400 por polegada; sufixos in/cm/mm/pt/px aceitos); tamanhos de fonte são pt. `genoffice slides read <file>` lista ids e geometria.',
        'Vocabulário:',
        docs.opVocabulary(),
      ].join('\n')
    },
  }
}

type SlidesOpDocs = typeof import('@genoffice/pptx-ops/op-docs')

function slidesCatalog(docs: SlidesOpDocs): OpCatalog {
  const ops: OpEntry[] = Object.entries(docs.OP_DOCS)
    .filter(([, d]) => d.aiCallable !== false && !d.pending)
    .map(([op, d]) => ({ op, group: d.group, signature: d.sig, doc: d.body, examples: d.examples }))
  const groups = docs.OP_GROUPS.map((g) => ({
    name: g,
    summary: docs.OP_GUIDES[g].summary,
    ops: ops.filter((e) => e.group === g).map((e) => e.op),
  }))
  return withFingerprint({ domain: 'slides', groups, ops })
}

function docsGuideText(catalog: OpCatalog, htmlRules: string, group?: string): string {
  const lines = [
    'Ops do Word (genoffice docs apply --ops): um array JSON; cada entrada tem "op".',
    'Alvos: { "blockIndexes": [..] } ou { "nodeType": "docHeading"|"docParagraph"|"docListItem"|"image", "headingLevel"? }; obtenha os índices com `genoffice docs read`.',
    'Notação de campos: simples = string, n = número, bool = booleano, ? = opcional, a|b = uma das opções.',
    '',
    ...renderGroups(catalog, group),
  ]
  if (!group || group === 'content') lines.push('', htmlRules)
  return lines.join('\n')
}

function opText(entry: OpEntry): string {
  const lines = opLines(entry)
  if (entry.doc) lines.push('', entry.doc)
  else if (entry.schema) lines.push('', JSON.stringify(entry.schema, null, 2))
  return lines.join('\n')
}

async function fingerprints(
  domain: GuideDomain | undefined,
  json: boolean,
): Promise<CommandResult> {
  const domains = domain ? [domain] : DOMAINS
  const out: Record<string, string> = {}
  for (const d of domains) out[d] = (await load(d, undefined, false)).catalog.fingerprint
  return {
    summary: Object.entries(out)
      .map(([d, fp]) => `${d} ${fp}`)
      .join('\n'),
    ...(json ? { detail: { fingerprints: out } } : {}),
  }
}
