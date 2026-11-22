# Spec de deck para `genoffice create --type pptx --spec`

O fluxo de deck com design (`genoffice guide slides design`) mantém uma página por arquivo:

```
deck/pages/01.json   { "title": "The 40% squeeze", "type": "cover", "layout": "cover_typography_hero", "background": "#0E1A2B", "elements": [ ... ] }
deck/pages/02.json   { "title": "Where the margin went", "type": "content", "layout": "three_column_cards", "background": "#FFFFFF", "elements": [ ... ] }
```

`title`, `type` e `layout` repetem a entrada do outline da página; o builder os ignora, as verificações os comparam. `genoffice slides check deck/pages/01.json` constrói e audita um arquivo, então o verifica em relação a `outline.json` e `style.md` (procurados ao lado do arquivo e uma pasta acima, ou `--outline <file>`): o arquivo `NN.json` é a entrada `pages[N-1]` do outline; um campo repetido que diverge ou texto placeholder sai com código 1, um título ausente ou uma foto planejada sem imagem é um aviso, e cores que a folha de estilo não nomeia são listadas em `detail.style.offPalette`. `genoffice create --type pptx --spec deck/pages [--outline deck/outline.json] --out deck.pptx` pega todo `*.json` da pasta em ordem de nome (`01`, `02`, … `10`), um slide cada, executa as mesmas verificações em cada arquivo e se recusa a construir quando a contagem de páginas difere do outline ou uma página diverge de sua entrada. `genoffice slides replace deck.pptx --slide n --spec deck/pages/NN.json` reconstrói o slide n a partir de seu arquivo após verificá-lo em relação à entrada n.

`--spec` também aceita um único arquivo contendo o deck inteiro, `{ "pages": [ {...}, {...} ] }`, um array simples de páginas, ou um único objeto de página. Até 60 páginas. Cada página se torna um slide 16:9; a tela é de **1280 × 720 px**, origem no canto superior esquerdo, todos `x` `y` `w` `h` inteiros em px. Os elementos são desenhados na ordem do array (primeiro = inferior). Cores são `#RRGGBB` ou `#RRGGBBAA` (`AA` alpha, `00` transparente).

## Outline

`deck/outline.json` é o plano a partir do qual as páginas são escritas; `genoffice slides check` o valida:

```json
{
  "topic": "Regional logistics review 2024",
  "core_hook": "Margins fell 40% while volume grew",
  "pages": [
    {
      "title": "The 40% squeeze",
      "type": "cover",
      "layout": "cover_typography_hero",
      "brief": "Main title, subtitle with the period, presenter line; dark background from the style sheet.",
      "image_queries": []
    },
    {
      "title": "Where the margin went",
      "type": "content",
      "layout": "three_column_cards",
      "brief": "Three cards: fuel +31%, wages +18%, spot rates -22% (2024 vs 2022), one source line each.",
      "image_queries": ["container port cranes at dawn"]
    }
  ]
}
```

`type` é `cover`, `content`, `data` ou `closing`; `layout` é uma das variantes desse tipo na biblioteca de layouts do guia de design. `image_queries` contém uma consulta de cena em inglês por espaço de foto ou uma URL `http(s)` pronta; `[]` para uma página sem fotos. Erros da verificação: tipo ou layout desconhecido, título, brief, core hook ou entradas de image_queries ausentes, texto placeholder (`XX%`, `lorem`, `TBD`), um layout de content ou data repetido em páginas consecutivas. Avisos: um brief muito fino, uma consulta que não está em inglês, menos de três variantes de layout nas páginas de corpo, uma primeira página que não é cover, uma última página que não é closing.

## Tipos de elemento

**Forma**

```json
{
  "type": "shape",
  "shape": "roundRect",
  "x": 80,
  "y": 120,
  "w": 360,
  "h": 200,
  "fill": "#1F3A5F",
  "stroke": { "color": "#FFFFFF", "widthPt": 1 },
  "valign": "middle",
  "paragraphs": [
    { "align": "center", "runs": [{ "text": "Label", "sizePt": 14, "color": "#FFFFFF" }] }
  ]
}
```

`shape` é um de `rect`, `roundRect`, `ellipse`, `triangle`, `rightArrow`, `leftArrow`, `upArrow`, `downArrow`, `chevron`, `diamond`, `parallelogram`, `trapezoid`, `hexagon`, `pentagon`, `pie`, `donut`, `star5`, `heart`, `cloud`, `line`, `lineArrow`. `line` e `lineArrow` desenham a diagonal de sua caixa do canto superior esquerdo ao inferior direito e precisam de um `stroke`; adicione `"flipV": true` para desenhá-la do canto inferior esquerdo ao superior direito, ou `"flipH": true` para espelhá-la horizontalmente; uma régua horizontal é uma caixa com `h: 1`. As duas flags espelham qualquer forma — um `rightArrow` com `"flipH": true` aponta para a esquerda. Uma forma precisa de um `fill` ou um `stroke` (ou ambos). `paragraphs` é texto de rótulo opcional, centralizado verticalmente a menos que `valign` diga o contrário.

**Texto**

```json
{
  "type": "text",
  "x": 80,
  "y": 60,
  "w": 800,
  "h": 90,
  "valign": "top",
  "paragraphs": [
    {
      "align": "left",
      "lineSpacingPct": 110,
      "spaceAfterPt": 6,
      "bullet": false,
      "runs": [
        {
          "text": "Headline",
          "sizePt": 36,
          "bold": true,
          "color": "#0E1A2B",
          "font": "Helvetica Neue"
        }
      ]
    }
  ]
}
```

Campos de run: `text` (obrigatório), `sizePt` (6 a 160), `bold`, `italic`, `color`, `font`. Campos de parágrafo: `runs` (obrigatório), `align` (`left` `center` `right` `justify`), `lineSpacingPct` (60 a 300), `spaceBeforePt` e `spaceAfterPt` (0 a 96), `bullet` (true = um marcador • com recuo pendente). Um parágrafo pode misturar runs de diferentes tamanhos, pesos e cores (um run de número grande ao lado de um run de unidade pequena). Caixas de texto têm zero preenchimento interno e quebram na largura; o builder mede cada caixa de texto simples e a aumenta para baixo quando o conteúdo é mais alto (ele nunca encolhe uma caixa).

**Imagem**

```json
{ "type": "image", "url": "./assets/hero.jpg", "x": 660, "y": 80, "w": 540, "h": 560 }
```

`url` é um caminho de arquivo local (absoluto, ou relativo ao diretório atual e depois à pasta do arquivo spec), uma URL `data:`, ou uma URL `http(s)` que o genoffice baixa. A imagem é recortada no centro para preencher sua caixa (object-fit: cover), então dê à caixa aproximadamente a proporção da imagem: uma caixa que mantém menos da metade da imagem (uma foto de produto vertical em uma faixa achatada) é posicionada, mas relatada como aviso. Até 8 imagens por página. Uma imagem que não pode ser lida é descartada e listada em `detail.imageFailures`.

## Validação

Elementos sem geometria numérica, totalmente fora da tela ou de um tipo desconhecido são descartados com aviso; caixas que ultrapassam as bordas são limitadas à tela. Elementos de texto sem nenhum texto e formas sem fill ou stroke são descartados. Duas caixas de texto em uma página com texto idêntico ou quase idêntico (um subtítulo que repete a legenda do gráfico) são mantidas, mas geram aviso, assim como uma caixa de imagem que recorta mais da metade de sua imagem. Uma página sem nenhum elemento restante falha; o deck falha apenas quando nenhuma página sobrevive. Tudo o que foi descartado ou gerou aviso é relatado em `detail.issues` por página. Até 48 elementos por página.

## Após a construção

Elementos recebem ids duráveis (`e_*`) e os slides `s_<n>`; `genoffice slides read deck.pptx --json` os lista para ops posteriores de `genoffice slides apply`. `genoffice slides audit` relata elementos fora dos limites, transbordamento de texto e sobreposição por slide com os mesmos ids; `genoffice slides render` escreve um PNG por slide para revisão visual; `genoffice slides replace --slide n --spec <page.json>` troca um slide por uma página reconstruída (os ids de seus elementos são novos).
