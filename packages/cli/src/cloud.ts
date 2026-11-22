import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { setGskProxyUrl } from '@genoffice/ai-search'
import { genofficeUserDataDir } from './gui'
import { packagedResourcesDir } from './resources'

/**
 * The cloud commands (search / image / media) reuse the editors' provider
 * routing: RamonMods when signed in (~/.genoffice/auth.json) and cloud tools
 * are on, otherwise the BYOK provider chosen in the app's AI settings. That
 * settings file lives in the shell's Electron userData directory, which genoffice
 * has to locate without Electron.
 */
export function aiSettingsPath(env: NodeJS.ProcessEnv): string {
  return env.GENOFFICE_AI_SETTINGS || join(genofficeUserDataDir(env), 'ai-settings.json')
}

/** First http(s) proxy in the usual environment variables, as the app's main process reads them. */
export function proxyUrlFromEnv(env: NodeJS.ProcessEnv): string | null {
  return (
    [
      env.HTTPS_PROXY,
      env.https_proxy,
      env.HTTP_PROXY,
      env.http_proxy,
      env.ALL_PROXY,
      env.all_proxy,
    ].find((v) => v && /^https?:\/\//.test(v)) ?? null
  )
}

let prepared = false

/** Once per process: proxy for fetch and the ramonmods children, and the packaged ramonmods CLI location. */
export async function prepareCloud(env: NodeJS.ProcessEnv): Promise<void> {
  if (prepared) return
  prepared = true
  const proxy = proxyUrlFromEnv(env)
  if (proxy) {
    setGskProxyUrl(proxy)
    const { ProxyAgent, setGlobalDispatcher } = await import('undici')
    setGlobalDispatcher(new ProxyAgent(proxy))
  }
  const packaged = packagedResourcesDir()
  if (packaged && !env.RAMONMODS_CLI_PATH) {
    const entry = join(packaged, 'ramonmods', 'node_modules', '@genspark', 'cli', 'dist', 'index.js')
    if (existsSync(entry)) process.env.RAMONMODS_CLI_PATH = entry
  }
}
