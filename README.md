# Monitor TSI - Modularizado

Este workspace contém o userscript `monitor-tsi.user.js` e um conjunto de módulos extraídos em `src/` para facilitar a manutenção.

## Estrutura

- `monitor-tsi.user.js`
  - Arquivo final do userscript.
  - Pode ser gerado automaticamente a partir dos módulos em `src/`.

- `src/`
  - Contém módulos individuais extraídos do script principal.
  - Há três agregados principais para edição rápida:
    - `src/01-core.js`
    - `src/02-ui.js`
    - `src/03-features.js`
  - Os outros arquivos JS são módulos menores que podem ser editados individualmente.

## Como editar

Opções recomendadas:

1. Editar diretamente um módulo específico em `src/`.
   - Exemplo: `src/39-envio-em-lote.js`, `src/42-gerar-relatorio-whatsapp.js`, `src/56-estilos.js`.

2. Editar em um dos agregados se quiser trabalhar em áreas amplas:
   - `src/01-core.js` — inicialização, carregamento de dados, utilitários gerais.
   - `src/02-ui.js` — layout do painel, modais, temas e UI.
   - `src/03-features.js` — funcionalidades específicas como relatórios, integração, módulos extras.

> Se você editar um agregado, mantenha em mente que ele é gerado automaticamente a partir dos módulos menores. Para evitar perda de mudanças, prefira editar os módulos menores sempre que possível.

## Como reconstruir `monitor-tsi.user.js`

1. Abra o terminal na pasta `monitorbrendo`.
2. Execute:

```powershell
node build.js
```

Isso irá:
- ler o `monitor-tsi.user.js` atual,
- extrair as seções em módulos dentro de `src/`,
- e gerar novamente `monitor-tsi.user.js` a partir das seções encontradas.

## Notas

- O script `build.js` foi criado para dividir o arquivo principal em módulos.
- Se você quiser um fluxo mais seguro, use `src/01-core.js`, `src/02-ui.js` e `src/03-features.js` como pontos de entrada para edição rápida.
- Depois de editar, rode `node build.js` para atualizar `monitor-tsi.user.js`.
