# Painel Bombeira de Garra — XI ENBOM 2026

Painel público de acompanhamento da competição operacional "Bombeira de Garra", servido pelo Google Apps Script HTML Service e alimentado por uma planilha do Google Sheets.

**Fluxo:**
- Fiscais registram penalidades/tempo em pranchetas de papel.
- Gestor digita os dados na planilha.
- Gestor clica em `🏆 Painel ENBOM → 📢 Atualizar painel público` (menu da planilha).
- Público acompanha ranking em tempo real pelo celular (QR Code com a URL do Web App).

---

## 📁 Arquivos deste workspace

| Arquivo | Para que serve |
|---|---|
| `Setup.gs` | **Script de inicialização** — cria todas as abas, validações, fórmulas e dados de referência automaticamente |
| `Codigo.gs` | Backend Apps Script — copiar para o projeto Apps Script |
| `Index.html` | Frontend (página mobile) — copiar para o projeto Apps Script |
| `templates/*.csv` | Dados de referência (caso prefira popular manualmente em vez de usar `Setup.gs`) |
| `Edital Bombeira de Garra - ENBOM 2026..pdf` | Edital original |

---

## 🚀 Setup rápido (modo automático)

### Opção A — Script automático (Recomendado, ~30 segundos)

1. Crie uma planilha em branco no Google Drive (ou use a que você já tem).
2. **Extensões → Apps Script**.
3. Cole `Setup.gs`, `Codigo.gs` e `Index.html` no projeto:
   - Renomeie o `Code.gs` padrão para `Codigo.gs` e cole o conteúdo.
   - Crie novo arquivo `Setup.gs` (➕ → Script).
   - Crie novo arquivo `Index` do tipo HTML (➕ → HTML).
4. No editor, com `Setup.gs` aberto, selecione a função **`inicializarPlanilha`** no menu suspenso (ao lado do botão Executar) e clique em **▶ Executar**.
5. Autorize as permissões pedidas (é uma planilha sua, é seguro).
6. Em ~10 segundos, todas as 8 abas estarão criadas com cabeçalhos, validações, fórmulas e dados de referência.
7. **(Opcional)** Para testar com dados fictícios: rode `popularDadosExemplo` da mesma forma. Adiciona 4 equipes, 8 participantes, 4 resultados e 1 penalidade.
8. Volte na planilha, recarregue (F5) — o menu `🏆 Painel ENBOM` aparecerá.
9. **Implante o Web App** (próxima seção abaixo).

> Depois de rodar `inicializarPlanilha` com sucesso, você pode deletar o arquivo `Setup.gs` do projeto — ele não é mais necessário em produção.

---

## 🛠️ Setup manual (alternativo, passo a passo)

> Use esta seção apenas se preferir configurar as abas à mão em vez de usar o `Setup.gs`.

### 1. Criar a planilha

1. No Google Drive, **Novo → Planilha**. Nomeie: `Painel ENBOM 2026`.
2. Criar **8 abas** (renomeie/adicione conforme necessário):

#### Aba `Equipes`
Cabeçalhos (linha 1):
```
nome_equipe | pais | corporacao | unidade_subnacional | bandeira_emoji | capita | observacoes
```
Validação na coluna `pais`: **Dados → Validação de dados → Intervalo: `RefPaises!A2:A`**.

#### Aba `Participantes`
```
nome | nome_de_guerra | posto_graduacao | pais | corporacao | equipe | eh_capita
```
- Validação na coluna `pais` (D): `RefPaises!A2:A`.
- Validação na coluna `equipe` (F): `Equipes!A2:A`.
- Coluna `eh_capita` (G): tipo **caixa de seleção** (Inserir → Caixa de seleção).
- `nome_de_guerra`: ex.: "SOUZA", "AKAUANY" — exibido em destaque no painel quando a equipe é expandida.

#### Aba `Resultados` ⭐ (a que você mais vai usar)
```
equipe | pista | ordem | hora_inicio | status | tempo_bruto_mmss | qtd_penalidades | segundos_penalidade | tempo_total_seg | motivo_dsq | observacoes
```
- Validação `equipe`: `Equipes!A2:A`.
- Validação `pista`: lista direta `1, 2`.
- Validação `status`: lista direta `Aguardando, Em Prova, Concluída, Desclassificada`.
- Validação `motivo_dsq`: `RefDSQ!A2:A` (números 1 a 7).
- **Fórmula na coluna `segundos_penalidade`** (a partir de H2): `=SE(G2="";0;G2*20)`.
- **Fórmula na coluna `tempo_total_seg`** (a partir de I2): `=SE(F2="";0;(VALOR(ESQUERDA(F2;LOCALIZAR(":";F2)-1))*60+VALOR(EXT.TEXTO(F2;LOCALIZAR(":";F2)+1;2)))+H2)`.
- Formato `tempo_bruto_mmss`: **Texto simples** (para aceitar `mm:ss`).
- Formato `hora_inicio`: **Hora**.

#### Aba `Penalidades`
```
equipe | item | descricao | fase | hora | fiscal | observacoes
```
- Validação `equipe`: `Equipes!A2:A`.
- Validação `item`: `RefPenalidades!A2:A` (números 1 a 10).
- Validação `fase`: lista direta `I, II, III, IV, V`.
- **Fórmula em `descricao`** (a partir de C2): `=SEERRO(PROCV(B2;RefPenalidades!A:B;2;FALSO);"")`.

#### Aba `RefPenalidades`
Copie de `templates/RefPenalidades.csv` (10 itens conforme Anexo D do edital).

#### Aba `RefPaises`
Copie de `templates/RefPaises.csv`.

#### Aba `RefDSQ`
Copie de `templates/RefDSQ.csv` (7 motivos conforme Anexo D do edital).

#### Aba `Config`
Copie de `templates/Config.csv`.

---

### 2. Criar o projeto Apps Script

1. Na planilha, **Extensões → Apps Script**. Vai abrir um novo projeto.
2. Renomeie o projeto: `Painel ENBOM 2026`.
3. **Renomeie `Code.gs` para `Codigo.gs`** (clique nos três pontos do arquivo no painel esquerdo) e cole o conteúdo de [`Codigo.gs`](Codigo.gs) deste repositório.
4. Clique no ➕ ao lado de "Arquivos" → **HTML** → nome: `Index`. Cole o conteúdo de [`Index.html`](Index.html).
5. Salve (`Ctrl+S`).
6. Volte na planilha e **recarregue a aba** (F5). Você deverá ver o menu `🏆 Painel ENBOM` aparecer no topo (após autorizar pela primeira vez).

---

## 📡 Implantar como Web App

1. No editor do Apps Script, clique em **Implantar → Nova implantação**.
2. Engrenagem → **Aplicativo da Web**.
3. Configure:
   - **Descrição:** `Painel ENBOM v1`
   - **Executar como:** `Eu (seu-email@gmail.com)`
   - **Quem tem acesso:** `Qualquer pessoa` (para público acessar sem login)
4. Clique em **Implantar**. Autorize permissões.
5. **Copie a URL gerada** (algo como `https://script.google.com/macros/s/.../exec`).
6. Na planilha: `🏆 Painel ENBOM → 🔗 Mostrar URL do painel` — janela com o link e botão para gerar QR Code.

> 🔄 **Sempre que alterar o código `.gs` ou `.html`**, é preciso fazer **nova implantação** (ou Gerenciar Implantações → editar → versão "Nova versão") para que o público veja a mudança. A URL muda a cada nova implantação se você criar nova implantação — use "Editar implantação existente" para manter a mesma URL.

---

## 🧪 Testar

1. Se ainda não fez, rode `popularDadosExemplo` do `Setup.gs` para inserir dados fictícios — OU cadastre manualmente algumas equipes/resultados.
2. Clique em `🏆 Painel ENBOM → 📢 Atualizar painel público`.
3. Abra a URL do painel no celular — deve mostrar pódio (se 3+ concluídas), em prova, aguardando, concluídas ordenadas, desclassificadas.

---

## 🧮 Como o cálculo é feito

- **Tempo bruto** é digitado pelo gestor no formato `mm:ss` (ex.: `12:34`).
- **Penalidades** somam **+20 segundos cada** (Anexo D, item 18 do edital).
- **Tempo total** = tempo bruto + (nº penalidades × 20).
- **Desempate** (Art. 16 do edital): menor tempo total → menor nº de penalidades → menor tempo bruto.

---

## 🌎 Suporte a equipes internacionais

O painel suporta equipes de fora do Brasil. Em `Equipes`:
- `pais`: escolher da lista `RefPaises` (já vem Brasil + países sul-americanos).
- `corporacao`: nome da corporação de bombeiros (ex.: `CBMAC`, `Bomberos de Buenos Aires`).
- `bandeira_emoji`: opcional — preencher manualmente ou deixar em branco (puxa de `RefPaises` automaticamente).

> ⚠️ **Atenção:** o edital original (Art. 3º–4º) prevê apenas Corpos de Bombeiros Militares brasileiros. A inclusão de equipes estrangeiras depende de autorização expressa da Comissão Organizadora ENBOM 2026 — confirmar antes da divulgação pública.

---

## 🛠️ Solução de problemas

| Sintoma | Como resolver |
|---|---|
| Menu `🏆 Painel ENBOM` não aparece na planilha | Recarregue a aba (F5). Se ainda não aparecer, abra o Apps Script, rode `onOpen` manualmente uma vez para autorizar permissões. |
| Painel mostra dados antigos depois de eu alterar a planilha | Clique em `🏆 Painel ENBOM → 📢 Atualizar painel público`. O cache fica 6h sem isso. |
| Página retorna erro de permissão | Confirme em **Implantar → Gerenciar implantações** que `Quem tem acesso = Qualquer pessoa`. |
| Tempo total dá errado | Confirme formato `mm:ss` em `tempo_bruto_mmss` (texto, não hora). Verifique fórmulas em `segundos_penalidade` e `tempo_total_seg`. |
| Mudei o código e o painel não atualizou | Faça nova implantação ou edite a implantação existente (Apps Script → Implantar → Gerenciar implantações). |
