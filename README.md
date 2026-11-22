<p align="center">
  <a href="https://github.com/RamonMods/GenOffice">
    <picture>
      <source srcset="docs/assets/readme/hero-dark.svg" media="(prefers-color-scheme: dark)">
      <img src="docs/assets/readme/hero.svg" alt="GenOffice â€” suÃ­te de escritÃ³rio com IA open-source: Docs, Sheets, Slides, PDF, Markdown e HTML com painel de IA integrado" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>A primeira suÃ­te de escritÃ³rio com IA open-source e completa.</b><br>
Arquivos Word, Excel, PowerPoint e PDF, editados por vocÃª e pela sua IA, salvos de volta nos formatos reais.</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/RamonMods/GenOffice" alt="LicenÃ§a: Apache-2.0"></a>
  <a href="https://github.com/RamonMods/GenOffice/releases/latest"><img src="https://img.shields.io/github/v/release/RamonMods/GenOffice" alt="Ãšltima versÃ£o"></a>
  <a href="https://github.com/RamonMods/GenOffice/releases"><img src="https://img.shields.io/github/downloads/RamonMods/GenOffice/total" alt="Downloads"></a>
  <a href="https://github.com/RamonMods/GenOffice/stargazers"><img src="https://img.shields.io/github/stars/RamonMods/GenOffice?style=flat" alt="Estrelas no GitHub"></a>
  <br><br>
  <a href="https://pixgg.com/ramon"><img src="https://img.shields.io/badge/Pix-pixgg.com%2Framon-32BCAD?style=for-the-badge&logo=pix&logoColor=white" alt="Apoiar o projeto"></a>
</p>

<p align="center">
  <a href="#download"><b>Download</b></a> Â·
  <a href="#recursos"><b>Recursos</b></a> Â·
  <a href="#linha-de-comando-e-skill-de-agente"><b>CLI</b></a> Â·
  <a href="#servidor-mcp"><b>MCP</b></a> Â·
  <a href="#privacidade"><b>Privacidade</b></a> Â·
  <a href="CONTRIBUTING.md"><b>Contribuir</b></a>
</p>

GenOffice Ã© uma alternativa gratuita e open-source ao Microsoft Office para macOS, Windows e Linux. Abre e salva arquivos nativos `.docx`, `.xlsx` e `.pptx`, edita PDF, Markdown e HTML, e coloca um agente de IA ao lado de cada documento â€” nÃ£o um chat preso na lateral, mas um editor que lÃª o arquivo, faz a alteraÃ§Ã£o e mostra exatamente o que tocou.

- **Formatos reais, preservaÃ§Ã£o de bytes.** SÃ³ o que vocÃª edita Ã© reescrito. Todo o resto do arquivo sobrevive byte a byte, para que os documentos continuem funcionando no Word, Excel e PowerPoint.
- **IA que vocÃª pode revisar.** EdiÃ§Ãµes chegam como alteraÃ§Ãµes rastreadas e diffs com rollback em um clique. Planilhas recebem fÃ³rmulas vivas, nÃ£o nÃºmeros colados. ApresentaÃ§Ãµes e pÃ¡ginas sÃ£o geradas no canvas e permanecem totalmente editÃ¡veis.
- **Local por design.** Arquivos abrem, editam, salvam e convertem na sua mÃ¡quina. PDF â†’ Word / Excel / PowerPoint, Markdown â†’ Word e HTML â†’ Word rodam localmente. SÃ³ as chamadas de IA saem da mÃ¡quina, para o provedor que vocÃª escolher.
- **Busque arquivos pelo que eles dizem.** A tela inicial pesquisa nomes, pastas e texto completo dos seus arquivos `.docx`, `.xlsx`, `.pptx`, PDF, Markdown e HTML a partir de um Ã­ndice SQLite local, incluindo CJK.
- **Suas chaves ou nenhuma.** Entre com RamonMods e pule chaves, ou traga sua prÃ³pria chave para Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter e qualquer endpoint compatÃ­vel com OpenAI, inclusive servidores locais.
- **ScriptÃ¡vel e pronto para agentes.** O app inclui a CLI `genoffice` e uma skill para Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode e Windsurf, para que um agente de cÃ³digo possa criar, converter, ler e editar arquivos Office reais na sua mÃ¡quina sem abrir janela.

## Download

**Baixe agora:** [macOS](https://github.com/RamonMods/GenOffice/releases/latest) (Apple Silicon e Intel) Â· [Windows](https://github.com/RamonMods/GenOffice/releases/latest) (x64 e Arm) Â· [Linux](https://github.com/RamonMods/GenOffice/releases/latest) (deb, rpm, AppImage).

| Plataforma | Requisitos | Arquivo |
|---|---|---|
| **macOS** â€” Apple Silicon (arm64) | macOS 11+ | [`.dmg` mais recente (arm64)](https://github.com/RamonMods/GenOffice/releases/latest) |
| **macOS** â€” Intel (x64) | macOS 11+ | [`.dmg` mais recente (x64)](https://github.com/RamonMods/GenOffice/releases/latest) |
| **Windows** (x64) | Windows 10+, Intel/AMD | [Instalador `-x64.exe` mais recente](https://github.com/RamonMods/GenOffice/releases/latest) |
| **Windows** on Arm (ARM64) | Windows 11 on Arm | [Instalador `-arm64.exe` mais recente](https://github.com/RamonMods/GenOffice/releases/latest) |
| **Linux** â€” Debian / Ubuntu | x86_64, glibc 2.34+ | [`.deb` mais recente](https://github.com/RamonMods/GenOffice/releases/latest) |
| **Linux** â€” Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ | [`.rpm` mais recente](https://github.com/RamonMods/GenOffice/releases/latest) |
| **Linux** â€” outras distros | x86_64, glibc 2.34+, FUSE 2 | [`.AppImage` mais recente](https://github.com/RamonMods/GenOffice/releases/latest) |

VersÃµes anteriores estÃ£o na [pÃ¡gina de Releases](https://github.com/RamonMods/GenOffice/releases).

## Recursos

### Docs â€” edite `.docx` com IA revisÃ¡vel
Abre o arquivo como o Word organiza: seÃ§Ãµes em duas colunas, imagens, tabelas sombreadas, cabeÃ§alhos e rodapÃ©s, paginaÃ§Ã£o nas mÃ©tricas de linha do Word. Estilos, comentÃ¡rios, alteraÃ§Ãµes rastreadas, equaÃ§Ãµes, tinta e controles de conteÃºdo fazem round-trip sem tocar no que nÃ£o foi editado.

### Sheets â€” `.xlsx` com fÃ³rmulas vivas e grÃ¡ficos
O agente cria planilhas com `SUMIF`s reais, grÃ¡ficos e alteraÃ§Ãµes em lote desfazÃ­veis. Perguntas sobre a pasta de trabalho voltam com o raciocÃ­nio e as cÃ©lulas exatas usadas como citaÃ§Ãµes clicÃ¡veis.

### Slides â€” de um prompt a um deck `.pptx`
Gere apresentaÃ§Ãµes no canvas, edite totalmente e exporte como `.pptx` real.

### PDF, Markdown e HTML
Edite PDF, Markdown e HTML com o mesmo painel de IA. ConversÃµes locais: PDF â†’ Word/Excel/PowerPoint, Markdown â†’ Word, HTML â†’ Word.

## Linha de comando e skill de agente

O app inclui a CLI `genoffice` para criar, converter, ler e editar arquivos Office reais na sua mÃ¡quina.

```bash
genoffice --help
```

## Servidor MCP

GenOffice tambÃ©m expÃµe ferramentas via MCP para agentes compatÃ­veis.

## Privacidade

Arquivos abrem, editam, salvam e convertem localmente. Apenas chamadas de IA saem da mÃ¡quina, para o provedor que vocÃª escolher. Veja [PRIVACY.md](PRIVACY.md).

## Contribuindo

ContribuiÃ§Ãµes sÃ£o bem-vindas. Veja [CONTRIBUTING.md](CONTRIBUTING.md).

## SeguranÃ§a

Para reportar vulnerabilidades, veja [SECURITY.md](SECURITY.md).

## LicenÃ§a

[Apache-2.0](LICENSE). GenOffice Ã© mantido por [RamonMods](https://github.com/RamonMods).

<p align="center">
  <a href="https://pixgg.com/ramon"><img src="https://img.shields.io/badge/Apoiar%20via%20Pix-pixgg.com%2Framon-32BCAD?style=for-the-badge&logo=pix&logoColor=white" alt="Apoiar via Pix"></a>
</p>

