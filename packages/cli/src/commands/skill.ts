import { mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import {
  agentDisplayLabel,
  agentIds,
  agentTarget,
  bundledSkillFrom,
  currentSkillInstalled,
  detectAgents,
  findLegacySkillCopies,
  installSkill,
  LEDGER_KEY,
  ledgerFromSettings,
  readInstallState,
  sweepLegacySkillCopies,
  type AgentId,
  type AgentTarget,
  type BundledSkill,
  type LegacySkillCopy,
  type SkillInstallState,
  type SkillLedger,
} from '../agent-skills'
import { flagBool, flagString } from '../args'
import { genofficeUserDataDir } from '../gui'
import type { CommandContext, CommandDef } from '../registry'
import { bundledSkillPath } from '../resources'
import { CliError, EXIT, type Warning } from '../result'
import { didYouMean } from '../suggest'

/**
 * The headless twin of Settings → Integrations: list the coding agents on this
 * machine and copy the bundled skill into their skills directory. Writes go
 * into the same ledger the app keeps (userData/app-settings.json), so the pane
 * shows a CLI install as "installed", not "foreign".
 */
export const skillCommand: CommandDef = {
  name: 'skill',
  summary:
    'Lista os agentes de codificação encontrados nesta máquina e instala ou atualiza o skill empacotado do GenOffice para eles.',
  usage: 'skill list | skill install <agent|all> [--dir <path>] [--force] | skill path',
  options: [
    { name: 'list', description: 'igual a `skill list`' },
    {
      name: 'dir',
      value: 'path',
      description: 'install: grava neste diretório de skills em vez disso',
    },
    {
      name: 'force',
      description:
        'install: grava também quando o agente não é detectado, a cópia instalada é mais recente ou foi editada manualmente',
    },
  ],
  async run(args, ctx) {
    const sub = flagBool(args, 'list') ? 'list' : (args.positionals[0] ?? 'list')
    switch (sub) {
      case 'list':
        return list(ctx)
      case 'path': {
        const path = skillFile()
        return { summary: path, detail: { path, version: bundled().version } }
      }
      case 'install':
        return install(args.positionals[1], flagString(args, 'dir'), flagBool(args, 'force'), ctx)
      default:
        throw new CliError(
          EXIT.usage,
          `subcomando de skill desconhecido: ${sub}`,
          { supported: ['list', 'install', 'path'] },
          {
            reason: 'invalid_argument',
            suggestion: didYouMean(sub, ['list', 'install', 'path'])
              ? `você quis dizer \`${didYouMean(sub, ['list', 'install', 'path'])}\`?`
              : 'use `skill list`, `skill install <agent|all>` ou `skill path`',
          },
        )
    }
  },
}

interface AgentRow {
  /** null for a folder given with --dir */
  agent: AgentId | null
  label: string
  detected: boolean
  skills_dir: string
  status: SkillInstallState['status']
  installed_version: string | null
  up_to_date: boolean
}

/** a stale per-tool copy the shared `.agents` row replaced */
interface LegacyRow {
  agent: AgentId
  label: string
  skills_dir: string
  path: string
  /** the copy still holds the bytes this app wrote, so it can be removed */
  ours: boolean
}

function legacyRow(copy: LegacySkillCopy): LegacyRow {
  return {
    agent: copy.agent,
    label: copy.label,
    skills_dir: copy.skillsDir,
    path: copy.path,
    ours: copy.ours,
  }
}

function list(ctx: CommandContext) {
  const skill = bundled()
  const ledger = readLedger(ctx.env)
  const detected = new Set(detectAgents(ctx.env, home(ctx.env)).map((a) => a.id))
  const agents: AgentRow[] = agentIds().map((id) => {
    const target = agentTarget(id, ctx.env, home(ctx.env))!
    return row(target, detected.has(id), readInstallState(target.skillsDir, skill, ledger))
  })
  const current = agents.filter((a) => a.detected && a.up_to_date).length
  const found = agents.filter((a) => a.detected)
  const legacy = findLegacySkillCopies(ledger, ctx.env, home(ctx.env)).map(legacyRow)
  return {
    summary: [
      found.length
        ? `${found.length} agente(s) detectado(s), ${current} com o skill ${skill.version}: ${found
            .map((a) => `${a.agent} ${a.installed_version ?? '-'}`)
            .join(', ')}`
        : 'nenhum agente de codificação detectado nesta máquina',
      legacy.length ? `${legacy.length} cópia(s) legada(s) de skill por ferramenta restante(s)` : '',
    ]
      .filter(Boolean)
      .join('; '),
    detail: { bundled_version: skill.version, skill_path: skillFile(), agents, legacy },
  }
}

function install(
  name: string | undefined,
  dir: string | undefined,
  force: boolean,
  ctx: CommandContext,
) {
  if (!name && !dir) {
    throw new CliError(
      EXIT.usage,
      'skill install precisa de um nome de agente, `all` ou --dir',
      undefined,
      {
        reason: 'missing_argument',
        suggestion: `agentes: ${agentIds().join(', ')}`,
      },
    )
  }
  const skill = bundled()
  const ledger = readLedger(ctx.env)
  const detected = new Set(detectAgents(ctx.env, home(ctx.env)).map((a) => a.id))
  const targets: Array<{ target: AgentTarget | null; skillsDir: string; detected: boolean }> = []
  if (dir) {
    targets.push({ target: null, skillsDir: resolve(ctx.cwd, dir), detected: true })
  } else if (name === 'all') {
    for (const id of agentIds()) {
      if (!detected.has(id)) continue
      const t = agentTarget(id, ctx.env, home(ctx.env))!
      targets.push({ target: t, skillsDir: t.skillsDir, detected: true })
    }
    if (!targets.length) {
      throw new CliError(EXIT.usage, 'nenhum agente de codificação detectado nesta máquina', undefined, {
        reason: 'unsupported',
        suggestion: 'instale em uma pasta com --dir <diretório de skills>',
      })
    }
  } else {
    const t = agentTarget(name as AgentId, ctx.env, home(ctx.env))
    if (!t) {
      const guess = didYouMean(name!, agentIds())
      throw new CliError(
        EXIT.usage,
        `agente desconhecido: ${name}`,
        { supported: [...agentIds(), 'all'] },
        {
          reason: 'unsupported',
          suggestion: guess
            ? `você quis dizer \`${guess}\`?`
            : `agentes: ${agentIds().join(', ')}, ou all`,
        },
      )
    }
    if (!detected.has(t.id) && !force) {
      throw new CliError(
        EXIT.usage,
        `${agentDisplayLabel(t)} não está instalado nesta máquina (${t.skillsDir} seria criado)`,
        { agent: t.id, skills_dir: t.skillsDir },
        { reason: 'unsupported', suggestion: 'repita com --force para criar a pasta mesmo assim' },
      )
    }
    targets.push({ target: t, skillsDir: t.skillsDir, detected: detected.has(t.id) })
  }

  // every target is checked before the first write, so a refusal never leaves a half-done batch
  for (const t of targets) {
    const before = readInstallState(t.skillsDir, skill, ledger)
    const blocked = refusal(before, force)
    if (blocked) {
      throw new CliError(
        EXIT.file,
        `${t.target?.id ?? t.skillsDir}: ${blocked} (${before.path})`,
        { agent: t.target?.id ?? null, skills_dir: t.skillsDir, status: before.status },
        { reason: 'output_exists', suggestion: refusalHint(before.status) },
      )
    }
  }
  const warnings: Warning[] = []
  const written: string[] = []
  const rows: Array<AgentRow & { action: 'written' | 'unchanged' }> = []
  let dirty = false
  for (const t of targets) {
    const label = t.target?.id ?? t.skillsDir
    let action: 'written' | 'unchanged' = 'unchanged'
    if (readInstallState(t.skillsDir, skill, ledger).status === 'installed') {
      warnings.push({
        code: 'already_up_to_date',
        message: `${label}: skill ${skill.version} já instalado`,
      })
    } else {
      installSkill(t.skillsDir, skill, ledger, 'cli')
      dirty = true
      action = 'written'
      written.push(label)
    }
    rows.push({
      ...row(t.target, t.detected, readInstallState(t.skillsDir, skill, ledger)),
      action,
    })
  }
  // the shared `.agents` copy is the one every legacy tool reads now; once it
  // is in place the per-tool copies it replaced are stale duplicates, so they
  // go. Without it they are all the user has, so they stay.
  const sweep = currentSkillInstalled(ledger, ctx.env, home(ctx.env))
    ? sweepLegacySkillCopies(ledger, ctx.env, home(ctx.env))
    : { removed: [], kept: [] }
  for (const c of sweep.removed) {
    warnings.push({
      code: 'legacy_copy_removed',
      message: `${c.label}: removida a cópia obsoleta em ${c.path}`,
    })
  }
  for (const c of sweep.kept) {
    warnings.push({
      code: 'legacy_copy_kept',
      message: `${c.label}: a cópia em ${c.path} foi editada manualmente; mantida no lugar`,
      suggestion: 'exclua-a você mesmo se a ferramenta ainda a usar',
    })
  }
  if (dirty || sweep.removed.length) writeLedger(ctx.env, ledger)
  return {
    summary: written.length
      ? `skill ${skill.version} instalado para ${written.join(', ')}`
      : `skill ${skill.version} já está atualizado`,
    ...(warnings.length ? { warnings } : {}),
    detail: { bundled_version: skill.version, agents: rows },
  }
}
function refusal(state: SkillInstallState, force: boolean): string | null {
  if (state.status === 'occupied') return 'a pasta genoffice ali contém outra coisa'
  if (force) return null
  if (state.status === 'newer')
    return `um skill mais recente ${state.installedVersion} está instalado; sem rebaixamento`
  if (state.status === 'modified') return 'o skill instalado foi editado manualmente; sem sobrescrever'
  return null
}

function refusalHint(status: SkillInstallState['status']): string {
  if (status === 'occupied') return 'mova essa pasta para outro lugar ou instale com --dir em outro diretório'
  return 'repita com --force para substituí-lo'
}

function row(target: AgentTarget | null, detected: boolean, state: SkillInstallState): AgentRow {
  return {
    agent: target?.id ?? null,
    label: target ? agentDisplayLabel(target) : state.path,
    detected,
    skills_dir: target?.skillsDir ?? state.path.replace(/[\\/]genoffice[\\/]SKILL\.md$/, ''),
    status: state.status,
    installed_version: state.installedVersion ?? null,
    up_to_date: state.status === 'installed' || state.status === 'newer',
  }
}

/** test seam: the home the agent dotfolders are probed under */
function home(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_HOME || homedir()
}

function skillFile(): string {
  const path = bundledSkillPath()
  if (!path)
    throw new CliError(EXIT.app, 'o skill empacotado (skills/genoffice/SKILL.md) não foi encontrado')
  return path
}

function bundled(): BundledSkill {
  return bundledSkillFrom(readFileSync(skillFile()))
}

function settingsPath(env: NodeJS.ProcessEnv): string {
  return join(genofficeUserDataDir(env), 'app-settings.json')
}

function readSettings(env: NodeJS.ProcessEnv): Record<string, unknown> {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath(env), 'utf8'))
    if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw as Record<string, unknown>
  } catch {}
  return {}
}

function readLedger(env: NodeJS.ProcessEnv): SkillLedger {
  return ledgerFromSettings(readSettings(env))
}

/** Same file and key the app uses; replaced atomically so a concurrent app write never sees a half file. */
function writeLedger(env: NodeJS.ProcessEnv, ledger: SkillLedger): void {
  const path = settingsPath(env)
  const next = { ...readSettings(env), [LEDGER_KEY]: ledger }
  const tmp = `${path}.${process.pid}.tmp`
  mkdirSync(dirname(path), { recursive: true })
  try {
    writeFileSync(tmp, JSON.stringify(next, null, 2), { encoding: 'utf8', flag: 'wx' })
    renameSync(tmp, path)
  } catch (err) {
    try {
      unlinkSync(tmp)
    } catch {}
    throw err
  }
}
