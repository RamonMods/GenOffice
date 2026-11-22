# GenOffice — como rodar (português)

Projeto baixado em `C:\Users\ramon\Downloads\genoffice`.
Idioma da interface: **português do Brasil (pt-BR)** por padrão.

## Antes de tudo

O Node está instalado em `C:\Program Files\nodejs`, mas a política do PowerShell
bloqueia o `npm.ps1`. Use sempre o `npm.cmd`:

```powershell
cd C:\Users\ramon\Downloads\genoffice
$env:PATH = "C:\Program Files\nodejs;$env:PATH"
& "C:\Program Files\nodejs\npm.cmd" --version
```

## Instalar dependências

```powershell
& "C:\Program Files\nodejs\npm.cmd" install --no-audit --no-fund
```

## Rodar

Opção 1 — modo desenvolvimento (7 editores + shell Electron):

```powershell
& "C:\Program Files\nodejs\npm.cmd" run dev
```

Opção 2 — build completo e abrir o Electron:

```powershell
& "C:\Program Files\nodejs\npm.cmd" run shell
```

Opção 3 — só o editor de documentos no navegador:

```powershell
& "C:\Program Files\nodejs\npm.cmd" run dev:docs
```

## Trocar o idioma na interface

Se quiser outro idioma, use o seletor **Idioma / Language** nas configurações.
O padrão foi alterado no código para `pt` (português). Para voltar ao comportamento
original, edite `apps/shell/src/main/index.ts` e restore a linha:

```ts
uiLang ??= normalizeLang(app.getLocale())
```

## O que foi alterado em relação ao repositório original

- Idioma inicial: `pt` no shell, no pacote `@genoffice/i18n` e nos bootstraps dos 7 apps.
- `<html lang="pt-BR">` nos arquivos HTML dos renderers.
- Os dicionários `pt` já vinham no projeto (não foi preciso traduzir do zero).

## Documentação original

- `README.pt-BR.md` — README em português.
- `docs/` — documentação completa.
