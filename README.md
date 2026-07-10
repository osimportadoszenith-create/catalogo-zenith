# Catálogo Zenith Imports

Site estático de catálogo (preços, marcas e frete) da Zenith Imports, separado do projeto principal do site (`zenith-imports-catalog`).

## Estrutura

- `public/catalogo-zenith.html` — o catálogo final, um único HTML autocontido (fontes e imagens embutidas em base64). É o que fica no ar.
- `public/Tabela final atualizada .xlsx` — planilha de preços, fonte de verdade dos produtos.
- `public/assets/CAPAS/` — capas fotográficas de cada marca (1600×400px).
- `public/assets/MARCAS LOGO/SVG MARCAS/` — logos em SVG usados como alternativa quando não existe capa fotográfica.
- `public/assets/CAPA/HEAD.png` (dentro de `CAPAS/`) — imagem do topo (hero).
- `build-assets/fonts/` — fontes (Barlow, Big Shoulders Display) usadas na geração do HTML.
- `scripts/build-catalog.mjs` — script que lê a planilha + assets e gera `public/catalogo-zenith.html`.
- `src/middleware.ts` — bloqueia o acesso direto a `/catalogo-zenith.html`, forçando o uso do caminho `/py`.

## Como atualizar o catálogo

1. Atualize a planilha `public/Tabela final atualizada .xlsx` (mesmas colunas: GRUPO, MARCA, PRODUTO, APRESENTAÇÃO, VALOR FINAL).
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
