# QMC Vistoria — app web instalável (Android)

Checklist de vistoria As-built / HOF no celular ou tablet → relatório PPTX no modelo QMC.

## O que o app faz
Dois tipos de vistoria (botões na tela inicial):

**Pré-HOF**
- Cadastro: ID do Site, Empreendimento, Nº do projeto, Empresa implantadora, Operadora(s).
- Automático: data/hora de início e fim, localização GPS.
- Foto do empreendimento: enviada por você; vai para a capa do PPTX. Logos Vivo/Claro/TIM vêm do app.
- 18 grupos de checklist (OK / Não conf. / N/A; foto obrigatória quando há "Não conf."), observação por grupo, amperagens e medições.
- **Itens adicionais** em qualquer grupo (fora do padrão, ex.: HBTS em plataforma elevada: escada fixa, guarda-corpo, linha de vida): um item por linha, cada um com OK / Não conf. / N/A. Entram na conformidade, na regra de foto obrigatória e no PPTX.
- Fotos por grupo, cada uma com observação própria. PPTX: capa + (checklist + evidências) por grupo, no máximo 2 fotos por slide.

**Acompanhamento** (instalado × projetado)
- Cabeçalho igual + quantidades de projeto: setores, IDFs e Nanos (5G + 4G).
- Grupos HBTS, IDF e Nanos (N3RU): para cada item informa-se Projeto, Instalado, Fusionado e/ou Conectado; o app calcula o % de cada etapa, o avanço do item, do grupo e o geral.
- Avanço do item = média das etapas (instalado/projeto, fusionado/projeto, conectado/projeto). Avanço do grupo = média dos itens. Avanço geral = média dos grupos.
- "Criar próxima vistoria" gera a v2.0, v3.0… com o cabeçalho já preenchido e compara com a anterior (item, grupo e geral).
- PPTX: capa (com quantidades de projeto) → **slide de Conclusão** (avanço geral + observação) → por grupo: tabela + fotos.

Nos dois: funciona offline; sincroniza com o OneDrive de quem está logado (cada pessoa vê só as suas).
Os itens do acompanhamento ficam em `data/acompanhamento.json`.

## Publicar no GitHub Pages
1. Crie um repositório (ex.: `qmc-vistoria`) e envie **todo o conteúdo desta pasta** para a raiz dele.
2. Settings → Pages → *Deploy from a branch* → branch `main`, pasta `/ (root)`.
3. O endereço será `https://SEU-USUARIO.github.io/qmc-vistoria/` (com a barra no final).

## Login Microsoft (uma vez)
1. Acesse https://entra.microsoft.com (ou portal.azure.com) com a conta Microsoft.
2. **Identidade → Registros de aplicativo → Novo registro**.
   - Nome: `QMC Vistoria`
   - Tipos de conta: **Contas em qualquer diretório organizacional e contas pessoais da Microsoft**
   - URI de redirecionamento: plataforma **Aplicativo de página única (SPA)** →
     `https://SEU-USUARIO.github.io/qmc-vistoria/` (exatamente igual ao endereço do app, com a `/` final)
3. Em **Permissões de API** confirme: Microsoft Graph, delegadas → `User.Read` e `Files.ReadWrite`.
4. Copie o **ID do aplicativo (cliente)** e cole em `config.js` (`clientId`).
5. Publique de novo. No celular: abra o endereço no Chrome → menu ⋮ → **Instalar app**.

> Conta corporativa da QMC: o TI pode precisar autorizar o aplicativo (consentimento do administrador).
> Conta pessoal: o próprio usuário aceita na primeira entrada.

## Onde ficam os dados
No OneDrive de cada usuário: pasta **QMC Vistorias / `<ID-do-site>_<data>_<código>`**
contendo `vistoria.json`, `fotos/` e, se salvar, o `.pptx`. Nada passa por servidor da QMC nem do GitHub.

## Atualizar o app
Troque os arquivos no repositório e aumente `VERSION` em `sw.js` (ex.: `v1.0.1`) para os aparelhos
baixarem a versão nova.

## Alterar o checklist
Edite `data/checklist.json` (itens, grupos, campos de amperagem). Blocos repetidos ficam em
`blocos_reutilizaveis`. Não é preciso mexer no código.
