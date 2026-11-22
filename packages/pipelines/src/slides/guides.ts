import designMd from './guides/design.md?raw'
import specMd from './guides/spec.md?raw'

/** Agent-facing guides for building a deck from a spec; the app and the genoffice CLI print the same text. */
export const SLIDES_GUIDES: Readonly<
  Record<'design' | 'spec', { description: string; content: string }>
> = {
  design: {
    description:
      'Fluxo de design de apresentações: primeiro a folha de estilo, depois um esboço página a página com variantes de layout, regras de dimensionamento de texto, antipadrões e o ciclo de controle de qualidade',
    content: designMd,
  },
  spec: {
    description:
      'O JSON de especificação da apresentação que `create --type pptx --spec` transforma em um pptx: tela, tipos de elemento, campos e limites',
    content: specMd,
  },
}
