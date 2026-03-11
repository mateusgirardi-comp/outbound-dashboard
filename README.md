# Outbound Dashboard — Comp

Dashboard em tempo real para tracking de campanhas de outbound.

**Live:** https://mateusgirardi-comp.github.io/outbound-dashboard/

---

## Como funciona

```
Google Sheets (abas de campanha)
    → Google Apps Script (calcula métricas → JSON API)
        → Dashboard HTML (GitHub Pages)
```

O dashboard faz `fetch()` no Apps Script a cada 5 minutos. Para atualizar os dados, basta editar a planilha — sem precisar republicar nada.

---

## Setup inicial (uma vez)

### 1. Instalar clasp
```bash
npm install -g @google/clasp
clasp login   # abre o browser para autenticação Google
```

### 2. Criar o projeto Apps Script na planilha
```bash
cd apps-script
clasp create --type sheets \
  --parentId 1evy8peuLyilrhTndKHMbUgqDpWl55p8pCARax24TxGA \
  --title "Outbound Dashboard API"
```
Isso cria um `.clasp.json` com o `scriptId`. Atualize o `.clasp.json` na raiz do repo com esse `scriptId`.

### 3. Fazer push do código
```bash
clasp push
```

### 4. Deployar como Web App
```
No Apps Script (script.google.com):
  Implantar → Nova implantação
  Tipo: App da Web
  Executar como: Eu (owner)
  Quem tem acesso: Qualquer pessoa, mesmo anônimo
  → Implantar → copiar a URL
```

### 5. Atualizar o dashboard com a URL
Em `index.html`, linha ~9:
```js
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/SEU_ID/exec';
```

### 6. Push para GitHub
```bash
git add .
git commit -m "setup: configurar URL do Apps Script"
git push
```

GitHub Pages serve automaticamente o `index.html` como dashboard.

---

## Workflow de atualização de código

```bash
# editar Code.gs aqui
cd apps-script && clasp push   # deploy no Apps Script
git add . && git commit -m "feat: ..." && git push   # versiona no GitHub
```

---

## Atualizar sprints (início de cada mês)

Em `apps-script/Code.gs`, atualize o `SPRINT_CONFIG` no topo:

```js
var SPRINT_CONFIG = [
  { name: 'S1', start: new Date('2026-04-01'), end: new Date('2026-04-07') },
  { name: 'S2', start: new Date('2026-04-08'), end: new Date('2026-04-14') },
  { name: 'S3', start: new Date('2026-04-15'), end: new Date('2026-04-21') },
  { name: 'S4', start: new Date('2026-04-22'), end: new Date('2026-04-30') }
];
```

Depois: `clasp push` + `git push`.

---

## Estrutura do repo

```
outbound-dashboard/
  ├── apps-script/
  │   ├── Code.gs           ← lógica da API (Apps Script)
  │   └── appsscript.json   ← config do projeto
  ├── index.html            ← dashboard (GitHub Pages)
  ├── .clasp.json           ← scriptId do projeto
  └── README.md
```

---

## Como uma nova campanha é detectada

O Apps Script detecta automaticamente abas de campanha com as colunas `Nome`, `BDR` e `Empresa`. Basta criar uma nova aba na planilha com esse formato — ela aparecerá no dashboard no próximo refresh.
