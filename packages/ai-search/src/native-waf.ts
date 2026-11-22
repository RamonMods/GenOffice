type Fetcher = (url: string, init: RequestInit) => Promise<Response>

/** Legacy native WAF compatibility stub (backend removed). */
export function nativeWafHosts(_env: NodeJS.ProcessEnv = process.env): Set<string> {
  return new Set<string>()
}

export class NativeWaf {
  clear(): void {}

  async prepare(_url: string, init: RequestInit, _fetcher: Fetcher): Promise<RequestInit> {
    return { ...init, headers: new Headers(init.headers) }
  }
}

export const nativeWaf = new NativeWaf()
