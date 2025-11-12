# Avaliações — v4 (Google Apps Script)
Menu único + unificação **robusta** de nomes para consolidar médias de avaliações (membros e projetos) em planilhas do Google Forms/Sheets.

## ✨ Principais recursos
- **Menu “Avaliações → Atualizar médias”** (UI segura; funciona também por gatilho).
- **Unificação de nomes** (remove artigos “a/o/as/os”, acentos e stopwords; usa chave curta `primeiro + último`) para evitar duplicidades como  
  “a Manuella da Silva Padilha” ≡ “Manuella da Silva Padilha”.
- **Classificação automática de colunas**:
  - Membro: títulos tipo `Avalie o Diretor Fulano:` (padrões configuráveis).
  - Projeto: `Avalie o Projeto X:` etc.
  - Ignora campos de texto/feedback e metadados comuns do Forms.
- **Saídas** em três abas:
  - `Médias Individuais`
  - `Médias por Projeto`
  - `Resumo` (média global + relatório de unificações)
- **Gatilho** `onFormSubmit` instalado automaticamente (opcional).

## 🧩 Como usar (rápido)
1. Crie/abra sua planilha que recebe respostas do Forms.
2. `Extensões → App Script` e **cole** o conteúdo de `src/Code.js`.
3. Ajuste nomes das abas em `CFG` se necessário (ex.: `SHEET_RESPOSTAS`).
4. Execute `runAtualizarMedias` uma vez (vai pedir autorizações).
5. Use o menu **Avaliações → Atualizar médias** a qualquer momento.
6. (Opcional) O gatilho `onFormSubmit` será criado automaticamente.

## 🛠️ Padrões e configurações
Veja o objeto `CFG` no topo do código:
- `MEMBER_COL_PATTERNS` e `PROJECT_COL_PATTERNS` para casar títulos.
- `EXCLUDE_KEYWORDS` para ignorar textos/feedbacks.
- `COLOR_SCALE` para cores por faixa de média.

## 🔎 Requisitos de dados
- Notas numéricas na linha das perguntas (1–4, 0–10 etc.).  
- Campos textuais (feedback, justificativa…) serão ignorados.

## 📦 Instalação com CLASP (opcional)
```bash
npm i -g @google/clasp
clasp login
clasp create --type sheets --title "Avaliações v4" --rootDir ./src
# Em seguida: cole o Code.js e suba
clasp push
clasp open
