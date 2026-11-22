# Criando um deck a partir de uma spec

O mesmo fluxo de trabalho que o app GenOffice executa quando gera uma apresentação, escrito para um agente que faz o próprio raciocínio e deixa o motor fazer a construção. No app, cada etapa é uma chamada separada ao modelo; aqui, cada etapa é um arquivo separado, e a CLI verifica cada arquivo antes de a próxima etapa poder começar. Trabalhe de cima para baixo; nunca comece posicionando elementos.

Arquivos de um deck, em uma pasta própria (`deck/`):

| Etapa                  | Arquivo                                                      | Verificação                                                                                                                                       |
| ---------------------- | ------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1 folha de estilo      | `deck/style.md`                                              | relia antes de cada página                                                                                                                        |
| 2 outline              | `deck/outline.json`                                          | `genoffice slides check deck/outline.json --json`                                                                                                 |
| 3 páginas              | `deck/pages/01.json`, `02.json`, … um arquivo por página do outline | `genoffice slides check deck/pages/NN.json --json` após escrever cada um (também em relação à entrada do outline e à folha de estilo)            |
| 4 build                | `deck/deck.pptx`                                             | `genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/deck.pptx --json`                                          |
| 5 QC                   | `deck/shots/*.png`                                           | `genoffice slides render`, `genoffice slides audit`, `genoffice slides replace`                                                                   |

Nunca coloque mais de uma página em um arquivo, e nunca escreva uma página antes de sua entrada no outline existir: o build se recusa a executar até que toda página do outline tenha seu arquivo, e uma página escrita em uma única etapa focada, com a folha de estilo e seu brief à sua frente, é o que mantém um deck de 20 páginas consistente. Mantenha os três arquivos juntos: `slides check`, `create --spec` e `slides replace` procuram por `outline.json` e `style.md` ao lado dos arquivos de página e uma pasta acima, verificam cada página em relação à sua entrada no outline (o arquivo `NN.json` é a entrada `pages[N-1]`) e às cores da folha de estilo, e dizem isso em `detail.notes` quando qualquer um dos arquivos está faltando.

## 0. Verifique o que a máquina consegue fazer

Execute `genoffice capabilities --json` uma vez. Ele lê as próprias configurações do GenOffice e informa, sem chamada de rede, se a busca na web, a busca de imagens, a geração de imagens e a análise de mídia estão configuradas (um login do Genspark com as ferramentas em nuvem ativadas, ou uma chave que o usuário inseriu em Settings). Somente quando um recurso está configurado o deck pode usá-lo; quando nada está configurado, trabalhe com o material que você tem e use tipografia, blocos de cor e formas em vez de fotos. Nunca peça ao usuário para configurar uma chave apenas para um deck.

## 1. Primeiro a folha de estilo (uma por deck)

Decida o sistema de design antes de qualquer conteúdo e escreva-o em `deck/style.md`; releia esse arquivo antes de cada página. Use este formato, com valores concretos:

- **Fundos**: fundo principal; um por tipo de página (cover, content, data, closing). Respeite um tom explícito (tema escuro, cores da marca) antes de recorrer a um neutro claro. As páginas de conteúdo compartilham um fundo.
- **Cores**: texto principal, acento primário, acento secundário, fundo de cartão, borda, escritos como valores `#RRGGBB`. Um sistema de acentos para todo o deck: mesmo ao comparar várias empresas ou opções, não dê a cada uma sua própria cor. Nunca mais do que os dois acentos. A verificação de página relata toda cor que uma página usa e que a folha de estilo não nomeia (`detail.style.offPalette`); branco e preto são sempre permitidos.
- **Fontes**: fonte de título (Latina e CJK se o deck misturar escritas), fonte de corpo, faixa de tamanho de título, faixa de tamanho de corpo. Use fontes que existam na máquina de destino; em caso de dúvida, omita `font` e deixe o padrão do motor ser aplicado.
- **Biblioteca de layouts** (pelo menos duas variantes por tipo de página):
  - cover: `cover_typography_hero` (tipografia enorme, sem imagem), `cover_dark_minimal`, `cover_split_color` (blocos de cor 60/40), `cover_full_image_overlay` (foto mais sobreposição escura), `cover_magazine`, `cover_split_image` (texto à esquerda, imagem à direita)
  - content: `left_text_right_image`, `three_column_cards`, `hero_big_number`, `two_column_comparison`, `timeline_horizontal`, `full_image_text_overlay`
  - data: `kpi_cards_row`, `chart_with_insight`, `two_by_two_grid`
  - closing: `closing_cta`, `closing_thank_you`
- **Estilo geral**: uma frase descrevendo a linguagem de design.

## 2. Outline, página por página

Escreva `deck/outline.json`:

```json
{
  "topic": "…",
  "core_hook": "one sentence with tension, a number or a counter-intuitive contrast",
  "pages": [
    {
      "title": "…",
      "type": "cover | content | data | closing",
      "layout": "one variant from the style sheet's library for that type",
      "brief": "what goes in each region, with the real facts and figures",
      "image_queries": ["English scene query or http(s) URL", "…"]
    }
  ]
}
```

Em seguida, `genoffice slides check deck/outline.json --json`: erros (layout desconhecido, um layout de content repetido em páginas consecutivas, texto placeholder, core hook ausente) saem com código 1 e devem ser corrigidos no arquivo; avisos são apenas recomendações. Regras que a verificação impõe e as que ela não consegue impor:

- Um core hook para todo o deck: uma frase com tensão, um número ou um contraste contraintuitivo.
- As páginas de content em um deck não devem repetir a mesma variante de layout em páginas consecutivas, e o deck deve usar pelo menos três variantes diferentes.
- Escolha a variante a partir do conteúdo: três pontos paralelos → `three_column_cards`; um número-chave → `hero_big_number`; comparação → `two_column_comparison` ou `two_by_two_grid`; uma sequência → `timeline_horizontal`; métricas → `kpi_cards_row`.
- Sem placeholders ("XX%", "lorem"): todo número e nome vem do material que você tem. Com a busca na web configurada, pesquise os fatos e números que o brief precisa antes de escrever as páginas (`genoffice search "<query>" --json`). Sem material e sem busca, diga ao usuário quais números são ilustrativos; nunca apresente números inventados como fatos.
- Fotos: com a busca de imagens configurada, dê a cada espaço de foto uma consulta em inglês para uma cena concreta ("harbor cranes at dawn", não "logistics") e busque candidatos com `genoffice search --images "<query>" --max 6 --json`; coloque o `imageUrl` de um resultado diretamente no elemento de imagem da página (o builder faz o download), e guarde seu `sourceUrl` (a página de onde a imagem veio) para os créditos da página de closing: o `imageUrl` costuma ser um proxy do provedor de busca que não diz nada sobre origem ou licença. Dê à caixa da imagem aproximadamente a mesma proporção da foto; uma foto vertical comprimida em uma faixa achatada é recortada no centro em uma fatia fina e relatada por `slides check`. Com a geração de imagens configurada, `genoffice image "<prompt>" --aspect 16:9 --out assets/p3.jpg` cria uma foto ou ilustração na paleta do deck. Sem nenhum dos dois, planeje páginas que não precisem de foto.

## 3. Escreva a spec de cada página, um arquivo por vez

Para a página N: releia `deck/style.md` e a entrada do outline da página, resolva suas `image_queries` em arquivos ou URLs, escreva `deck/pages/NN.json` (dois dígitos, um objeto de página, veja `genoffice guide slides spec`) repetindo `title`, `type` e `layout` da entrada no topo do objeto, e execute `genoffice slides check deck/pages/NN.json --json`. Ele constrói apenas essa página e retorna os elementos descartados, as imagens ilegíveis e a auditoria de geometria (fora dos limites, texto mais alto que sua caixa, sobreposições), então verifica a página em relação à entrada N do outline (`detail.outline.findings`: um campo repetido que diverge ou texto placeholder é erro e sai com código 1; um título ausente no texto da página ou uma foto planejada sem elemento de imagem é um aviso) e em relação à folha de estilo (`detail.style.offPalette`). Corrija o arquivo até que `audit`, `findings` e `offPalette` estejam vazios, então passe para a página N+1. Não agrupe páginas: uma página por etapa, verificada, é o objetivo principal.

A tela é de 1280 × 720 px, origem no canto superior esquerdo, apenas inteiros; nada pode cruzar as bordas. Os elementos são desenhados na ordem do array: primeiro o fundo e as formas de decoração, depois as imagens, e o texto por último.

**Dimensionando caixas de texto** (as caixas têm zero preenchimento interno; o canto superior esquerdo é onde o primeiro glifo começa):

- uma linha tem cerca de `sizePt × 1.8` px de altura com `lineSpacingPct` 110
- um caractere CJK tem cerca de `sizePt × 1.35` px de largura, um caractere latino cerca de `sizePt × 0.7` px
- o texto quebra na largura da caixa: conte as linhas quebradas, dê à caixa essa altura mais uma linha extra
- o builder mede novamente cada caixa de texto simples e a aumenta para caber; ele nunca encolhe, então caixas pequenas demais empurram o que está abaixo delas. Dimensione corretamente em vez de depender da correção.

**Espaçamento**: ≥ 8 px entre o texto e a borda de um cartão, ≥ 20 px entre um título e seu subtítulo, ≥ 5 px entre blocos de texto empilhados. Verifique cada par antes de gerar a saída.

**Escala tipográfica (pt)**: títulos grandes de 32 a 48, subtítulos de 18 a 24, corpo de 12 a 15, números hero até 80.

**Preencha a página**: espalhe o conteúdo por toda a tela; nada de conteúdo amontoado na metade superior acima de um campo vazio. Faça textos e imagens tão grandes quanto o layout permitir.

**Visuais**: fotos apenas de fontes que você realmente tem (arquivos locais, resultados de busca, imagens geradas), no máximo três elementos de imagem por página, e apenas uma foto que realmente combine com o conteúdo da página: uma foto de banco de imagens irrelevante ou genérica é pior do que nenhuma. Sem fotos, trabalhe com tipografia, blocos de cor e formas; nunca finja uma foto com uma caixa cinza. Decorações semelhantes a ícones usam as formas permitidas, no máximo quatro ou cinco por página, relacionadas ao conteúdo. Sem emoji. Visuais de dados são compostos por formas `rect`, `donut` e `line` com tamanhos proporcionais aos valores reais. Apenas cores sólidas (alpha permitido).

## 4. Anti-padrões (trate como defeitos)

- Uma barra de acento vertical fina à esquerda de cada cartão, uma barra colorida no topo dos cartões, uma barra pequena à esquerda dos títulos. Expresse a hierarquia com contraste de fundo, peso e tamanho.
- Cartões arco-íris: um acento diferente por item.
- Blocos de canto decorativos ou linhas curtas; decoração que muda de lugar de página para página.
- Todas as páginas como uma lista "forma + subtítulo em negrito + descrição". A cover não deve ser apenas um título simples mais subtítulo: dê a ela um âncora visual (um bloco de cor grande, uma composição geométrica, um número enorme, uma imagem hero).
- Uma barra de cabeçalho em todo o deck com o título da página em todas as páginas. Varie a composição.
- Cartões irmãos cujos números hero misturam métricas (uma participação de mercado ao lado de um preço ao lado de uma contagem): uma linha, uma unidade, uma fonte. O rótulo de um número deve dizer exatamente o que a fonte mediu (participação de condução óssea não é participação de fones esportivos).
- Duas linhas que dizem a mesma coisa: um subtítulo que repete a legenda do gráfico, uma conclusão que repete um marcador.

## 5. Construa, olhe, audite, corrija

1. `genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/deck.pptx --json` (`--outline` pode ser omitido quando o arquivo fica ao lado da pasta). Os arquivos de página são lidos em ordem de nome; o build se recusa a executar enquanto uma página do outline não tiver arquivo ou um arquivo de página divergir de sua entrada. Leia `detail.issues`, `detail.imageFailures`, `detail.outline.findings` e `detail.style.offPalette`: cada um nomeia o arquivo da página.
2. `genoffice slides render deck/deck.pptx --out deck/shots --json` e olhe cada PNG (precisa ter o GenOffice instalado). Verifique transbordamento, colisões, contraste, recorte, hierarquia, espaços em branco.
3. `genoffice slides audit deck/deck.pptx --json` para obter os achados de geometria que o olho não percebe. Os ids de elemento nos achados são os que `genoffice slides read` e `genoffice slides apply` usam.
4. Corrija uma página em seu arquivo spec e recoloque-a com `genoffice slides replace deck/deck.pptx --slide <n> --spec deck/pages/NN.json --json` (o slide n é o arquivo de página n+1; a página é verificada em relação à entrada n do outline primeiro); os outros slides não são alterados. Pequenos ajustes também podem passar pelas ops de `genoffice slides apply` (`setTransform`, `setText`, `setTextStyle`, `setFill` …). No máximo duas rodadas de correção; então relate o que resta.

Corrija apenas defeitos objetivos e um acabamento contido: mantenha o texto, os fatos, a paleta e a composição. Não redesenhe uma página que já está limpa.
