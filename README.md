# Catálogo Zenith Imports

Catálogo minimalista da Zenith Imports, separado do projeto principal do site (`zenith-imports-catalog`) e conectado em tempo real ao catálogo oficial.

## Sincronização automática

- A fonte de verdade de produtos, preços, categorias e disponibilidade é `https://zenithpharmaimports.com/api/products`.
- A rota local `/api/catalog` consulta a fonte oficial sem cache, valida os registros e expõe ao navegador somente os campos usados pelo catálogo.
- A página atualiza os dados ao abrir, a cada 15 segundos e sempre que a aba volta ao foco.
- Produtos com status `out_of_stock` continuam visíveis com o selo **Indisponível**; itens removidos ou inativos não aparecem.
- Alterações de preço, categoria, marca, nome, apresentação, inclusão e remoção também são refletidas automaticamente.
- Se a fonte estiver temporariamente fora do ar, a página mantém o último conteúdo carregado ou usa a versão gerada pela planilha como segurança.

O endereço da fonte pode ser substituído no servidor com `ZENITH_CATALOG_SOURCE_URL`. Não é necessário configurar essa variável no ambiente oficial.

## Estrutura

- `public/catalogo-zenith.html` — o catálogo final, com fontes e imagens embutidas e o cliente de sincronização ao vivo. É o que fica no ar.
- `public/Tabela final atualizada .xlsx` — versão de segurança usada para gerar o conteúdo inicial enquanto a fonte oficial é carregada.
- `src/app/api/catalog/route.ts` — proxy somente de leitura que valida e reduz os dados públicos do catálogo oficial.
- `public/assets/CAPAS/` — capas fotográficas de cada marca (1600×400px).
- `public/assets/MARCAS LOGO/SVG MARCAS/` — logos em SVG usados como alternativa quando não existe capa fotográfica.
- `public/assets/CAPA/HEAD.png` (dentro de `CAPAS/`) — imagem do topo (hero).
- `build-assets/fonts/` — fontes (Barlow, Big Shoulders Display) usadas na geração do HTML.
- `scripts/build-catalog.mjs` — script que lê a planilha + assets e gera `public/catalogo-zenith.html`.
- `src/middleware.ts` — bloqueia o acesso direto a `/catalogo-zenith.html`, forçando o uso do caminho `/py`.

## Como atualizar a aparência ou o fallback

As alterações rotineiras de preço, categoria e disponibilidade devem ser feitas no catálogo oficial; não é preciso gerar nem publicar este projeto novamente.

Para alterar imagens, estilos ou atualizar a versão de segurança:

1. Atualize a planilha `public/Tabela final atualizada .xlsx` (mesmas colunas: GRUPO, MARCA, PRODUTO, APRESENTAÇÃO, VALOR FINAL), se necessário.
2. Para trocar/adicionar a capa de uma marca, coloque o PNG em `public/assets/CAPAS/` (1600×400px) e adicione a marca em `PHOTO_FILE_MAP` no topo de `scripts/build-catalog.mjs`.
3. Rode:
   ```bash
   npm install
   npm run build:catalog
   ```
4. Confira `public/catalogo-zenith.html` localmente (`npm run dev` e abra `http://localhost:3000/py`).
5. Commit e push.

## Acesso

O catálogo fica em `/py`:

```
https://seudominio.com/py
```

O nome real do arquivo (`/catalogo-zenith.html`) e a raiz (`/`) retornam 404 — só `/py` funciona. Não depende de nenhuma variável de ambiente.

## Frete

A tabela de frete por UF em `scripts/build-catalog.mjs` deve ser mantida em sincronia manual com `src/data/freight.ts` do site principal da Zenith, caso os valores mudem lá.
