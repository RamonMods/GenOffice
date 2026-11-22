import { existsSync } from 'node:fs'
import { homedir } from 'node:os'
import { resolve } from 'node:path'
import {
  mcpAgentIds,
  mcpConfigFile,
  mcpSnippet,
  readMcpEntry,
  removeMcpEntry,
  writeMcpEntry,
  type McpEntryState,
} from '../agent-mcp'
import { agentLabel, agentPresent, type AgentId } from '../agent-skills'
import { flagBool, flagString } from '../args'
import type { CommandContext, CommandDef } from '../registry'
import { CliError, EXIT, type Warning } from '../result'
import { didYouMean } from '../suggest'
import { mcpLaunchFromLauncher, withCompactSchemas, type McpLaunch } from '../mcp-launch'
import { launcherPath } from './install'

const SUBCOMMANDS = ['install', 'uninstall', 'list']

export const mcpCommand: CommandDef = {
  name: 'mcp',
  summary:
    "Serve todos os comandos como ferramentas do Model Context Protocol: via stdio para um cliente MCP nesta máquina (`claude mcp add --transport stdio genoffice -- genoffice mcp`, Cursor, Claude Desktop), ou com --http como um servidor Streamable HTTP que clientes em outras máquinas acessam por URL. Via HTTP, arquivos também trafegam: PUT /files/<name> envia um arquivo e toda ferramenta aceita URLs http(s) no lugar de caminhos; as saídas voltam como URLs de download e, quando pequenas, como recursos incorporados. Ops, specs e Markdown são passados inline de qualquer forma; uma nova apresentação passa por deck_start, deck_page e deck_build. `mcp install <agent|all>` registra o servidor stdio na configuração MCP de um agente de código (Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf); `mcp uninstall <agent>` remove; `mcp list` mostra a situação de cada agente. Os parâmetros ops, cells e data das ferramentas apply e create são anunciados com o schema tipado por operação de `guide <domain> --json`; --compact-schemas (ou GENOFFICE_MCP_COMPACT_SCHEMAS=1) os anuncia como arrays simples para clientes com orçamento de contexto pequeno.",
  usage:
    'mcp [--http <port> [--host <addr>] [--token <secret>]] [--compact-schemas] | mcp install <agent|all> [--dir <path>] [--force] [--compact-schemas] | mcp uninstall <agent|all> [--dir <path>] [--force] | mcp list',
  quiet: (args) => args.positionals.length === 0,
  options: [
    {
      name: 'http',
      value: 'port',
      description: 'serve Streamable HTTP nesta porta em vez de stdio (0 escolhe uma porta livre)',
    },
    {
      name: 'host',
      value: 'addr',
      description: 'endereço para vincular com --http (padrão: 127.0.0.1; 0.0.0.0 para outras máquinas)',
    },
    {
      name: 'token',
      value: 'secret',
      description:
        'com --http, exige este token bearer em toda requisição (padrão: $GENOFFICE_MCP_TOKEN; senão, nenhum)',
    },
    {
      name: 'compact-schemas',
      description:
        'anuncia ops, cells e data como arrays sem tipo em vez do schema por operação (tools/list menor; também GENOFFICE_MCP_COMPACT_SCHEMAS=1); `mcp install --compact-schemas` grava a mesma opção na entrada do agente',
    },
    {
      name: 'dir',
      value: 'path',
      description:
        "instalar/desinstalar: trata esta pasta como o diretório de configuração do agente (apenas um agente)",
    },
    {
      name: 'force',
      description:
        'instalar/desinstalar: também grava quando o agente não é detectado ou a entrada do genoffice inicia outro programa',
    },
  ],
  async run(args, ctx) {
    const sub = args.positionals[0]
    if (sub === undefined) {
      for (const flag of ['dir', 'force']) {
        if (flagBool(args, flag)) {
          throw new CliError(
            EXIT.usage,
            `--${flag} precisa de \`mcp install\` ou \`mcp uninstall\``,
            undefined,
            {
              reason: 'invalid_argument',
            },
          )
        }
      }
      return serve(args, ctx)
    }
    for (const flag of ['http', 'host', 'token']) {
      if (flagBool(args, flag)) {
        throw new CliError(
          EXIT.usage,
          `--${flag} serve MCP; não pode ser combinado com \`mcp ${sub}\``,
          undefined,
          {
            reason: 'invalid_argument',
          },
        )
      }
    }
    // compact schemas apply when serving (no subcommand) or when the entry records
    // the toggle; list/uninstall only inspect what is already there
    if (flagBool(args, 'compact-schemas') && sub !== 'install') {
      throw new CliError(
        EXIT.usage,
        `--compact-schemas serve MCP ou o registra; não pode ser combinado com \`mcp ${sub}\``,
        undefined,
        {
          reason: 'invalid_argument',
        },
      )
    }
    switch (sub) {
      case 'list':
        return list(ctx)
      case 'install':
      case 'uninstall':
        return edit(
          sub,
          args.positionals[1],
          flagString(args, 'dir'),
          flagBool(args, 'force'),
          flagBool(args, 'compact-schemas'),
          ctx,
        )
      default:
        throw new CliError(
          EXIT.usage,
          `subcomando mcp desconhecido: ${sub}`,
          { supported: SUBCOMMANDS },
          {
            reason: 'invalid_argument',
            suggestion: didYouMean(sub, SUBCOMMANDS)
              ? `você quis dizer \`${didYouMean(sub, SUBCOMMANDS)}\`?`
              : 'execute `mcp` sozinho para servir, ou `mcp install <agent|all>`, `mcp uninstall <agent>`, `mcp list`',
          },
        )
    }
  },
}

async function serve(args: Parameters<CommandDef['run']>[0], ctx: CommandContext) {
  const http = flagString(args, 'http')
  const compactSchemas =
    flagBool(args, 'compact-schemas') || envFlag(ctx.env.GENOFFICE_MCP_COMPACT_SCHEMAS)
  if (http === undefined) {
    // lazy: the server imports runCli, which registers this command
    const { serveStdio } = await import('../mcp/server')
    await serveStdio({ cwd: ctx.cwd, env: ctx.env, log: ctx.log, compactSchemas })
    return { summary: 'sessão mcp encerrada' }
  }
  const port = Number(http)
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    throw new CliError(EXIT.usage, `--http precisa de um número de porta, recebido: ${http}`, undefined, {
      reason: 'invalid_argument',
    })
  }
  const tokenFlag = flagString(args, 'token')
  // `--token ""` — e.g. `--token "$MY_TOKEN"` with the variable unset —
  // would reach authorized() as "no auth configured": the operator believes
  // the lock is on while every request passes. Fail closed instead.
  if (tokenFlag !== undefined && tokenFlag.trim() === '') {
    throw new CliError(
      EXIT.usage,
      '--token precisa de um valor não vazio (omita a opção ou defina GENOFFICE_MCP_TOKEN)',
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  const hostFlag = flagString(args, 'host')
  // an empty host skips the loopback default and binds every interface
  if (hostFlag !== undefined && hostFlag.trim() === '') {
    throw new CliError(
      EXIT.usage,
      '--host precisa de um valor não vazio (omita a opção para vincular 127.0.0.1)',
      undefined,
      { reason: 'invalid_argument' },
    )
  }
  const { serveHttp } = await import('../mcp/http')
  await serveHttp({
    cwd: ctx.cwd,
    env: ctx.env,
    log: ctx.log,
    port,
    host: hostFlag,
    token: tokenFlag ?? (ctx.env.GENOFFICE_MCP_TOKEN?.trim() || undefined),
    compactSchemas,
  })
  return { summary: 'servidor mcp interrompido' }
}

function envFlag(value: string | undefined): boolean {
  return /^(1|true|yes|on)$/i.test(value?.trim() ?? '')
}

type InstallStatus = 'installed' | 'updated' | 'unchanged' | 'occupied' | 'manual' | 'not_detected'
type UninstallStatus = 'removed' | 'absent' | 'occupied' | 'manual' | 'not_detected'

interface AgentRow {
  agent: AgentId
  label: string
  detected: boolean
  config: string
  registered: boolean
  /** what the existing genoffice entry starts, when there is one */
  command: string | null
  /** the registered entry asks for compact tool schemas (mcp install --compact-schemas) */
  compactSchemas: boolean
}

interface Target {
  id: AgentId
  label: string
  file: string
  detected: boolean
}

function home(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_HOME || homedir()
}

/** The entry every config gets: the absolute launcher, or on Windows the app run as Node on the bundle. */
function launcher(compactSchemas = false): { launcher: string; launch: McpLaunch } {
  const path = launcherPath()
  if (!path) throw new CliError(EXIT.app, 'não foi possível localizar o launcher do genoffice')
  return {
    launcher: path,
    launch: mcpLaunchFromLauncher(path, {
      platform: process.platform,
      exists: existsSync,
      compactSchemas,
    }),
  }
}

function row(t: Target, state: McpEntryState): AgentRow {
  return {
    agent: t.id,
    label: t.label,
    detected: t.detected,
    config: t.file,
    registered: state.status === 'registered' || state.status === 'stale',
    command: state.command,
    compactSchemas: state.compactSchemas === true,
  }
}

function list(ctx: CommandContext) {
  const { launcher: bin, launch } = launcher()
  const detected = mcpDetected(ctx)
  const agents = mcpAgentIds().map((id) => {
    const t = target(id, undefined, detected, ctx)
    const state = readMcpEntry(id, t.file, launch)
    return { ...row(t, state), status: state.status }
  })
  const found = agents.filter((a) => a.detected)
  return {
    summary: found.length
      ? `${found.length} agente(s) detectado(s), ${found.filter((a) => a.registered).length} com genoffice registrado: ${found
          .map((a) => `${a.agent} ${a.status}${a.compactSchemas ? ' (compacto)' : ''}`)
          .join(', ')}`
      : 'nenhum agente de código detectado nesta máquina',
    detail: { launcher: bin, launch, agents },
  }
}

function target(
  id: AgentId,
  dir: string | undefined,
  detected: Set<AgentId>,
  ctx: CommandContext,
): Target {
  const opts = { env: ctx.env, home: home(ctx.env), ...(dir ? { dir: resolve(ctx.cwd, dir) } : {}) }
  return {
    id,
    label: agentLabel(id),
    file: mcpConfigFile(id, opts)!,
    detected: dir !== undefined || detected.has(id),
  }
}

/** MCP detection is its own table: an agent counts when its dotfolder exists, whether or not it installs a skill row. */
function mcpDetected(ctx: CommandContext): Set<AgentId> {
  const env = ctx.env
  return new Set(mcpAgentIds().filter((id) => agentPresent(id, env, home(env))))
}

function edit(
  verb: 'install' | 'uninstall',
  name: string | undefined,
  dir: string | undefined,
  force: boolean,
  compactSchemas: boolean,
  ctx: CommandContext,
) {
  const ids = mcpAgentIds()
  if (!name) {
    throw new CliError(EXIT.usage, `mcp ${verb} precisa de um nome de agente ou \`all\``, undefined, {
      reason: 'missing_argument',
      suggestion: `agentes: ${ids.join(', ')}`,
    })
  }
  if (name === 'all' && dir !== undefined) {
    throw new CliError(EXIT.usage, '--dir se aplica a um agente, não a `all`', undefined, {
      reason: 'invalid_argument',
      suggestion: 'informe o agente cuja configuração fica nessa pasta',
    })
  }
  if (name !== 'all' && !ids.includes(name as AgentId)) {
    const guess = didYouMean(name, ids)
    throw new CliError(
      EXIT.usage,
      `agente desconhecido: ${name}`,
      { supported: [...ids, 'all'] },
      {
        reason: 'unsupported',
        suggestion: guess ? `você quis dizer \`${guess}\`?` : `agentes: ${ids.join(', ')}, ou all`,
      },
    )
  }
  const { launcher: bin, launch } = launcher()
  /** the form an install asks for: reads always compare against the plain launch */
  const wanted = verb === 'install' && compactSchemas ? withCompactSchemas(launch) : launch
  const detected = mcpDetected(ctx)
  const targets = (name === 'all' ? ids : [name as AgentId]).map((id) =>
    target(id, dir, detected, ctx),
  )
  const single = targets.length === 1 ? targets[0]! : null
  if (single && !single.detected && !force) {
    throw new CliError(
      EXIT.usage,
      `${single.label} não está instalado nesta máquina (${
        verb === 'install'
          ? `${single.file} seria gravado`
          : `nada a remover em ${single.file}`
      })`,
      { agent: single.id, config: single.file },
      {
        reason: 'unsupported',
        suggestion: `repita com --force para ${verb === 'install' ? 'gravar' : 'editar'} o arquivo mesmo assim`,
      },
    )
  }
  const forceHint =
    verb === 'install'
      ? 'repita com --force para substituí-lo'
      : 'repita com --force para remover a entrada estranha'
  const manualHint =
    verb === 'install'
      ? 'cole detail.snippet no arquivo manualmente'
      : 'remova a entrada do genoffice do arquivo manualmente'

  const warnings: Warning[] = []
  const rows: Array<AgentRow & { status: InstallStatus | UninstallStatus; snippet?: string }> = []
  for (const t of targets) {
    const before = readMcpEntry(t.id, t.file, launch)
    let status: InstallStatus | UninstallStatus
    if (!t.detected && !force) status = 'not_detected'
    else if (before.status === 'manual') status = 'manual'
    else if (before.status === 'occupied' && !force) status = 'occupied'
    else if (verb === 'install') {
      // an entry that already asks for the requested schema form stays untouched;
      // a plain entry reports no compactSchemas field, so compare the resolved flag
      const sameForm =
        before.status === 'registered' && (before.compactSchemas === true) === compactSchemas
      if (sameForm) status = 'unchanged'
      else {
        writeMcpEntry(t.id, t.file, wanted)
        status = before.status === 'absent' ? 'installed' : 'updated'
      }
    } else status = removeMcpEntry(t.id, t.file) ? 'removed' : 'absent'
    const after = status === 'manual' ? before : readMcpEntry(t.id, t.file, launch)
    const r = { ...row(t, after), status }
    rows.push(
      status === 'manual' && verb === 'install' ? { ...r, snippet: mcpSnippet(t.id, wanted) } : r,
    )
    if (status === 'occupied' || status === 'manual') {
      warnings.push({
        code: status === 'occupied' ? 'entry_occupied' : 'manual_edit_needed',
        message:
          status === 'occupied'
            ? `${t.id}: a entrada do genoffice em ${t.file} inicia ${before.command ?? 'outra coisa'}`
            : `${t.id}: ${t.file} não pôde ser editado; ${verb === 'install' ? 'adicione' : 'remova'} a entrada manualmente`,
        suggestion:
          status === 'occupied'
            ? forceHint
            : verb === 'install'
              ? 'veja detail.agents[].snippet'
              : manualHint,
      })
    }
  }
  if (single && (rows[0]!.status === 'occupied' || rows[0]!.status === 'manual')) {
    const r = rows[0]!
    throw new CliError(
      EXIT.file,
      r.status === 'occupied'
        ? `${single.id}: a entrada do genoffice em ${single.file} inicia ${r.command ?? 'outra coisa'}; não ${verb === 'install' ? 'substituindo' : 'removendo'}`
        : `${single.id}: ${single.file} não pôde ser editado com segurança`,
      {
        agent: single.id,
        config: single.file,
        status: r.status,
        command: r.command,
        ...(r.snippet ? { snippet: r.snippet } : {}),
      },
      { reason: 'output_exists', suggestion: r.status === 'occupied' ? forceHint : manualHint },
    )
  }
  const changed = rows.filter(
    (r) => r.status === 'installed' || r.status === 'updated' || r.status === 'removed',
  )
  const skipped = rows.filter((r) => r.status === 'not_detected').length
  const blocked = rows.filter((r) => r.status === 'occupied' || r.status === 'manual').length
  const tail = [
    skipped ? `${skipped} agente(s) não detectado(s)` : '',
    blocked ? `${blocked} agente(s) deixado(s) de lado (veja os avisos)` : '',
  ]
    .filter(Boolean)
    .join(', ')
  const nothing = verb === 'install' ? 'nada registrado' : 'nada removido'
  const summary = changed.length
    ? `${verb === 'install' ? 'genoffice mcp registrado para' : 'genoffice mcp removido de'} ${changed
        .map((r) => `${r.agent} (${r.status})`)
        .join(', ')}${tail ? `; ${tail}` : ''}`
    : rows.some((r) => r.status === 'unchanged')
      ? `genoffice mcp já registrado${tail ? `; ${tail}` : ''}`
      : rows.some((r) => r.status === 'absent')
        ? `genoffice mcp não estava registrado${tail ? `; ${tail}` : ''}`
        : blocked
          ? `${nothing}: ${tail}`
          : `${nothing}: nenhum agente de código detectado nesta máquina`
  return {
    summary,
    ...(warnings.length ? { warnings } : {}),
    detail: { launcher: bin, launch, agents: rows },
  }
}
