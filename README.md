# Painel Automático de Protocolos — RI Indaial

Consulta automaticamente todos os protocolos listados em `protocolos.json` no
site do Registro de Imóveis de Indaial, todos os dias às 8h (horário de
Brasília), e:

- Gera uma planilha `.xlsx` com o status de cada protocolo
- Atualiza uma aba do Google Sheets com os mesmos dados
- Envia um e-mail com a planilha e os PDFs de exigência anexados

Tudo isso roda sozinho no GitHub (gratuito), sem precisar de servidor.

## Passo a passo para deixar funcionando

### 1. Criar o repositório no GitHub
1. Crie uma conta no [github.com](https://github.com) se ainda não tiver.
2. Crie um repositório novo, **privado** (esses dados são sensíveis).
3. Suba todos os arquivos desta pasta para esse repositório.

### 2. Cadastrar seus protocolos
Edite o arquivo `protocolos.json` e liste todos os protocolos que quer
monitorar, no formato:
```json
[
  { "protocolo": "178024", "senha": "28175710", "referencia": "Lotes 09 e 10 PA3" }
]
```

### 3. Criar a Conta de Serviço do Google (para o Google Sheets)
1. Acesse [console.cloud.google.com](https://console.cloud.google.com).
2. Crie um projeto novo (qualquer nome).
3. Ative a **API do Google Sheets** (menu "APIs e Serviços" → "Ativar APIs").
4. Vá em "Credenciais" → "Criar Credenciais" → "Conta de Serviço".
5. Depois de criada, entre nela → aba "Chaves" → "Adicionar Chave" → tipo JSON.
   Isso baixa um arquivo `.json` — guarde-o, você vai usar no passo 5.
6. Copie o "e-mail" da conta de serviço (algo como
   `xxx@xxx.iam.gserviceaccount.com`).
7. Crie uma planilha nova no Google Sheets, e **compartilhe** ela com esse
   e-mail da conta de serviço (como Editor).
8. Pegue o ID da planilha — é o trecho da URL entre `/d/` e `/edit`.

### 4. Configurar o e-mail de envio
Se usar Gmail: ative a "verificação em duas etapas" na sua conta e gere uma
["senha de app"](https://myaccount.google.com/apppasswords) — não use sua
senha normal.

### 5. Cadastrar os "Secrets" no GitHub
No repositório: Settings → Secrets and variables → Actions → New repository
secret. Cadastre um por um:

| Nome | Valor |
|---|---|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | conteúdo inteiro do arquivo `.json` baixado no passo 3.5 |
| `GOOGLE_SHEET_ID` | o ID da planilha (passo 3.8) |
| `EMAIL_HOST` | `smtp.gmail.com` (se usar Gmail) |
| `EMAIL_PORT` | `587` |
| `EMAIL_USER` | seu e-mail remetente |
| `EMAIL_PASS` | a senha de app gerada no passo 4 |
| `EMAIL_TO` | e-mail(s) que devem receber o relatório, separados por vírgula |

### 6. Testar
No GitHub, vá em "Actions" → "Consultar Protocolos RI" → "Run workflow" para
rodar manualmente e conferir se tudo funciona. Depois disso, ele roda sozinho
todo dia às 8h.

## Adicionando/removendo protocolos no dia a dia

Basta editar `protocolos.json` direto pelo site do GitHub (não precisa
instalar nada) — clique no arquivo, no ícone de lápis, edite e clique em
"Commit changes".

## Se algo der errado

Vá em "Actions" no GitHub e abra a última execução — o log mostra exatamente
em qual protocolo/etapa travou. Pode colar esse log numa conversa comigo que
eu te ajudo a resolver.
