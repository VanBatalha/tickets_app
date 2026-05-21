# Chamados TI - Smartsheet → Rocket.Chat

Aplicação Flask para consultar chamados da planilha do Smartsheet, filtrar por responsável/Não atribuído e gerar mensagens prontas para Rocket.Chat.

## O que o app faz

- Consulta a planilha de chamados pelo Smartsheet API.
- Também roda localmente usando um arquivo `.xlsx` exportado, útil para testes sem API.
- Lista chamados por responsável na coluna **Atribuído a**, inclusive células com múltiplos contatos.
- Permite selecionar mais de um responsável no filtro.
- Permite filtrar por **Tipo de demanda**: **Todas**, **Interna** ou **Externa**.
- A lista de responsáveis é montada com base nos chamados **Não resolvidos**, para reduzir opções antigas/fechadas.
- Considera como demanda **Interna** qualquer categoria que seja ou comece com `TI Demandas Internos do Setor`; todo o restante entra como **Externa**.
- Lista no filtro **Não atribuído** apenas chamados cujo campo **Atribuído a** esteja preenchido explicitamente como “Não atribuído”. Campo vazio não entra nessa contagem nem nesse filtro.
- Chamados abertos marcados como **Não atribuído** aparecem no topo da lista para qualquer responsável selecionado.
- Exibe os campos principais:
  - Prioridade
  - Ticket
  - Enviado por
  - Categoria
  - Descrição da solicitação ou problema
  - Passos de reprodução do problema
  - Unidade
  - Centro de Custo
  - Número do Patrimônio do Equipamento
  - Criado em
- Mostra campos longos em pop-up pelo botão **Ver detalhes**.
- Abre uma tela de mensagens para o ticket com os botões de cópia do fluxo atual.
- A tela inicial abre por padrão com o filtro **Não resolvidos**.
- O botão **Abrir no Smartsheet** usa o link fixo/filtro configurado em `SMARTSHEET_WEB_URL`, separado do ID/token usado pela API.
- A tela de mensagens preserva os filtros da tela anterior ao usar **Voltar para chamados**.
- A tela de mensagens reaproveita o cache já carregado, evitando reler o Smartsheet a cada clique em **Aplicar**.
- Aceita **Referência da discussão** e **Link da pesquisa de satisfação** por chamado.
- Trata links longos do Smartsheet para não quebrar o Markdown do Rocket.Chat.

## Ajustar tamanho dos campos na tela

Na listagem de chamados, todos os blocos usam o mesmo tamanho. Para regular largura, altura e quantidade de linhas exibidas antes de cortar o texto, altere estas variáveis no arquivo `templates/index.html`:

```css
--ticket-field-min-width: 150px;
--ticket-field-height: 88px;
--ticket-value-lines: 3;
```

Os textos longos continuam disponíveis completos no botão **Ver detalhes**.

## Rodar localmente

No Windows, você pode executar:

```bat
run_local.bat
```

Ou, manualmente:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
python app.py
```

No Linux/macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
python app.py
```

Acesse: http://127.0.0.1:5000


## Filtro Interna/Externa

O filtro **Tipo de demanda** usa a coluna **Categoria**:

- **Interna**: quando a categoria é `TI Demandas Internos do Setor`, `TI Demandas Internos do Setor >` ou começa com esse prefixo, por exemplo `TI Demandas Internos do Setor > Auditoria e Análise`.
- **Externa**: qualquer categoria diferente desse padrão.

Se o nome dessa categoria mudar no Smartsheet, ajuste a função `is_internal_demand_category()` no arquivo `app.py`.

## Configurar `.env`

### Opção 1: Smartsheet API

Preencha:

```env
SMARTSHEET_ACCESS_TOKEN=seu_token
SMARTSHEET_SHEET_ID=id_da_planilha
DEFAULT_RESPONSIBLE=vanderson.batalha@grupocertare.com
SMARTSHEET_WEB_URL=https://app.smartsheet.com/sheets/4mxRwjJvcm57HJ6hwxcgVFhF799qcHM6FxMxw2C1?view=grid&newview=true&filterId=1922378880733060
```

O `DEFAULT_RESPONSIBLE` é opcional. Ele abre o app já filtrado no responsável informado. Para abrir com mais de um responsável selecionado por padrão, separe os valores por ponto e vírgula, por exemplo: `Vanderson Batalha;vanderson.batalha@grupocertare.com`.

### Opção 2: Teste local por XLSX

Deixe o token e o sheet ID vazios e preencha:

```env
XLSX_FILE_PATH=C:\caminho\para\Abertura de chamado de TI.xlsx
```

## Deploy no Render

### Opção A: usando `render.yaml`, recomendada para teste

1. Crie um repositório no GitHub com todos estes arquivos na raiz do projeto.
2. Faça commit e push.
3. No Render, crie um novo serviço usando **Blueprint** ou **New + Web Service**, apontando para o repositório.
4. Se usar o `render.yaml`, o Render já lê:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
5. Depois de criado, vá em **Environment** e preencha as variáveis marcadas como secret/sync false:
   - `SMARTSHEET_ACCESS_TOKEN`
   - `SMARTSHEET_SHEET_ID`
   - `DEFAULT_RESPONSIBLE`, opcional, exemplo: `Vanderson Batalha;vanderson.batalha@grupocertare.com`
6. Confirme também:
   - `SMARTSHEET_WEB_URL=https://app.smartsheet.com/sheets/4mxRwjJvcm57HJ6hwxcgVFhF799qcHM6FxMxw2C1?view=grid&newview=true&filterId=1922378880733060`
   - `CACHE_TTL_SECONDS=300`
   - `ENABLE_SMARTSHEET_WRITE=false`
7. Faça **Manual Deploy** ou aguarde o deploy automático após o push.

### Opção B: Web Service manual

Use estas configurações:

- Runtime: `Python`
- Build Command: `pip install -r requirements.txt`
- Start Command: `gunicorn app:app`
- Health Check Path: `/healthz`

Variáveis de ambiente mínimas:

```env
APP_SECRET_KEY=gere-uma-chave-grande
SMARTSHEET_ACCESS_TOKEN=seu_token
SMARTSHEET_SHEET_ID=id_da_planilha
SMARTSHEET_WEB_URL=https://app.smartsheet.com/sheets/4mxRwjJvcm57HJ6hwxcgVFhF799qcHM6FxMxw2C1?view=grid&newview=true&filterId=1922378880733060
DEFAULT_RESPONSIBLE=Vanderson Batalha;vanderson.batalha@grupocertare.com
CACHE_TTL_SECONDS=300
ENABLE_SMARTSHEET_WRITE=false
SMARTSHEET_BASE_URL=https://api.smartsheet.com/2.0
```

Para teste no Render, mantenha `ENABLE_SMARTSHEET_WRITE=false`. Assim o app apenas lê a planilha e gera mensagens, sem alterar nada no Smartsheet.

## Observação sobre atualização da planilha

O app foi entregue em modo seguro de leitura. Existe uma base no código para atualização de linhas via API, mas a interface não executa alterações no Smartsheet enquanto `ENABLE_SMARTSHEET_WRITE=false`.

Quando for habilitar ações como concluir chamado ou gravar referência da discussão, use preferencialmente um token pessoal do usuário responsável pela ação ou uma conta de serviço aprovada, porque a auditoria do Smartsheet tende a refletir o usuário/token que executou a chamada.

## Saúde da aplicação

Rota de health check:

```text
/healthz
```
