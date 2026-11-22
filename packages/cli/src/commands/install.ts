import { dirname, join } from 'node:path'
import { installCliLink } from '../install'
import type { CommandDef } from '../registry'
import { packagedResourcesDir, repoRoot } from '../resources'
import { CliError, EXIT } from '../result'

/** The launcher the terminal should run: the one shipped beside this bundle, or the checkout's. */
export function launcherPath(): string | null {
  const packaged = packagedResourcesDir()
  if (packaged)
    return join(packaged, 'cli', process.platform === 'win32' ? 'genoffice.cmd' : 'genoffice')
  const root = repoRoot()
  return root ? join(root, 'packages', 'cli', 'bin', 'genoffice') : null
}

export const installCommand: CommandDef = {
  name: 'install-cli',
  summary:
    'Coloca o genoffice no PATH (link simbólico em /usr/local/bin ou ~/.local/bin, ou no PATH do usuário no Windows).',
  usage: 'install-cli',
  async run() {
    const launcher = launcherPath()
    if (!launcher) throw new CliError(EXIT.app, 'não foi possível localizar o executável do genoffice')
    const r = installCliLink({ launcher })
    const detail = {
      status: r.status,
      launcher,
      ...(r.location ? { location: r.location } : {}),
      ...(r.pathHint ? { pathHint: r.pathHint } : {}),
    }
    const onPath = (verb: string) =>
      r.pathHint
        ? `o genoffice ${verb} está vinculado em ${r.location}, mas ${dirname(r.location ?? '')} não está no seu PATH; adicione com: ${r.pathHint}`
        : `o genoffice ${verb} está no PATH (${r.location})`
    switch (r.status) {
      case 'linked':
        return { summary: onPath('agora'), detail }
      case 'present':
        return { summary: onPath('já'), detail }
      case 'occupied':
        throw new CliError(EXIT.app, `${r.location} é outro programa; não será substituído`, {
          ...detail,
          manual: r.manual,
        })
      case 'unwritable':
        throw new CliError(EXIT.app, `${r.location} não é gravável; execute o comando manual`, {
          ...detail,
          manual: r.manual,
        })
      default:
        throw new CliError(EXIT.app, `a configuração do PATH não é suportada em ${process.platform}`, detail)
    }
  },
}
