import { describe, expect, it } from 'vitest'
import { httpBodyDetail } from '../src/http-error'

describe('httpBodyDetail', () => {
  it('passes API error bodies through, truncated', () => {
    expect(httpBodyDetail('{"error":"bad"}')).toBe('{"error":"bad"}')
    expect(httpBodyDetail('x'.repeat(600))).toHaveLength(500)
  })

  it('names the block page instead of dumping markup', () => {
    const page = `<!DOCTYPE html><html><head><title>\n  Attention Required! | Cloudflare\n</title></head>
<body><span>Cloudflare Ray ID: <strong class="font-semibold">a37f69f31ff7a28e</strong></span>
<span>error code: <span>1020</span></span></body></html>`
    expect(httpBodyDetail(page)).toBe(
      'o serviço retornou uma página da web ("Attention Required! | Cloudflare", código de erro 1020, ray a37f69f31ff7a28e) em vez de uma resposta de API (provavelmente um bloqueio temporário de rede ou do gateway) — verifique sua conexão e tente novamente',
    )
  })

  it('keeps the generic note for a bare HTML shell', () => {
    expect(httpBodyDetail('<html><body></body></html>')).toBe(
      'o serviço retornou uma página da web em vez de uma resposta de API (provavelmente um bloqueio temporário de rede ou do gateway) — verifique sua conexão e tente novamente',
    )
  })
})
