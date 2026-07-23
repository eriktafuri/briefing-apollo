# Briefing Apollo

Formulário de briefing estilo Typeform (uma pergunta por tela), com a identidade
visual da Apollo. As respostas e os arquivos enviados vão para uma Google Sheet
e uma pasta do Google Drive — sem precisar de servidor ou banco de dados.

## Arquivos

- `index.html` — o formulário (estático, hospedável em qualquer lugar, ex.: GitHub Pages).
  É **único e genérico** — o mesmo link serve pra qualquer cliente novo.
- `painel.html` — painel interno, protegido por senha, com a lista de todas as
  respostas recebidas. **Uso só seu** — não é pra mandar pro cliente.
- `resposta.html` — abre a partir do painel: mostra uma resposta específica
  formatada por inteiro, no mesmo estilo visual da proposta (em vez da linha
  crua da planilha).
- `Code.gs` — o backend, roda no Google Apps Script.
- `apollo-logo.png` — logo usado no cabeçalho.

## Como publicar (10 minutos, só uma vez)

1. **Crie uma Google Sheet nova** (qualquer nome, ex.: "Briefings Apollo").
2. Nela, vá em **Extensões → Apps Script**. Vai abrir um editor de código vazio.
3. Apague o conteúdo padrão e cole o conteúdo do arquivo `Code.gs` deste projeto.
4. Clique em **Implantar → Nova implantação**.
   - Tipo: **App da Web**
   - Executar como: **Eu (seu e-mail)**
   - Quem pode acessar: **Qualquer pessoa**
   - Clique em **Implantar** e autorize as permissões pedidas (é o script acessando
     sua própria planilha e seu próprio Drive).
5. Copie a **URL do app da Web** gerada (algo como `https://script.google.com/macros/s/AKfycb.../exec`).
6. Abra `index.html` e cole essa URL na constante `CONFIG.endpoint`, no topo do `<script>`:
   ```js
   var CONFIG = {
     endpoint: 'COLE_AQUI_A_URL_DO_APPS_SCRIPT', // ← troque por a URL copiada
     ...
   };
   ```
7. Suba `index.html` + `painel.html` + `resposta.html` + `apollo-logo.png` pra onde
   for hospedar (ex.: GitHub Pages, igual à proposta e ao relatório).

Pronto. Cada envio cria uma linha na aba **Respostas** da planilha.

## Painel interno (`painel.html` + `resposta.html`)

Depois que os clientes começam a responder, abra `painel.html` (senha padrão:
`apollo123` — troque em `PAINEL_SENHA` no `Code.gs`). Ele lista todas as
respostas recebidas; clicar em "Ver briefing" abre `resposta.html` com aquela
resposta inteira, formatada e legível — nome, e-mail, objetivo, público,
concorrência, tom de marca, arquivos enviados etc.

Esse painel é **só pra uso interno seu** — não gera nem precisa de link
personalizado por cliente. O formulário (`index.html`) continua sendo um
único link padrão, igual para todo mundo.

## Onde ficam os arquivos

No seu **Google Drive** (o mesmo Drive da conta que publicou o Apps Script),
na raiz ("Meu Drive"), é criada uma pasta **"Briefings Apollo — Arquivos"**
(uma vez só, no primeiro envio). Dentro dela, cada envio que tiver ao menos
um arquivo ganha a sua própria subpasta, nomeada `Empresa — dd-mm-aaaa hh-mm`
(ex.: `Grupo BrasilServ — 23-07-2026 14-30`) — os arquivos de identidade
visual e de conteúdo do mesmo envio caem juntos nessa mesma subpasta, com o
nome original de cada arquivo.

Na planilha, a coluna **pastaDrive** traz o link direto para essa subpasta
(abre tudo de uma vez); as colunas **identidadeVisual** e **conteudoArquivos**
trazem os links de cada arquivo individualmente, caso você queira abrir só um.

## Se precisar alterar as perguntas

Todas as perguntas ficam no array `BLOCKS`, dentro do `<script>` de `index.html`.
Cada bloco é uma seção (ex.: "Sobre a empresa") com uma lista de perguntas.
Tipos disponíveis: `text`, `email`, `tel`, `textarea`, `tags` (múltipla escolha
em chips) e `upload` (arquivos). Adicionar, remover ou reordenar perguntas não
exige mexer em mais nada — o motor (barra de progresso, navegação, envio) se
adapta sozinho à lista.

Se adicionar ou remover uma pergunta do tipo `text`/`email`/`tel`/`textarea`/`tags`,
lembre de também ajustar a lista `HEADERS` em `Code.gs` (senão a coluna nova não
aparece na planilha) e o objeto `payload` dentro de `doSubmit()` em `index.html`.

## Limites

- Até 5 arquivos por campo de upload, 15MB cada (ajustável em `CONFIG.maxFileSizeMB`
  e `CONFIG.maxFilesPerField`, no `index.html`).
- O Apps Script tem um teto de tamanho de requisição — 15MB por arquivo é seguro;
  não vale a pena subir muito esse número.
