# 📋 Painel de Triagem Automatizada de Paletes

Aplicação web para triagem operacional de materiais em paletes, com detecção automática de itens bloqueados a partir de uma planilha (Excel ou XML exportado do Excel).

## ✨ Funcionalidades

- **Importação de planilhas**: aceita arquivos `.xlsx`, `.xls` e `.xml` (formato SpreadsheetML do Excel), com múltiplas abas representando paletes diferentes.
- **Classificação automática**: separa os itens em duas tabelas — *Materiais Bloqueados* e *Itens Liberados* — com base em uma lista de códigos configurada em `script.js`.
- **Controle de progresso**: marque itens como concluídos com um clique; contadores de pendentes/concluídos são atualizados em tempo real.
- **Cópia rápida do código**: botão dedicado para copiar o código do material para a área de transferência.
- **Cronômetro de operação**: mede o tempo desde o carregamento do arquivo até a limpeza do painel.
- **Impressão de placas de bloqueio**: gera uma página de impressão com uma placa de identificação por material bloqueado, pronta para ser fixada fisicamente no palete.

## 🗂️ Estrutura do projeto

```
├── index.html   # Estrutura da página
├── style.css    # Estilos visuais
├── script.js    # Lógica da aplicação (leitura de arquivo, renderização, interações)
└── README.md
```

## 🚀 Como executar

Basta abrir o `index.html` em um navegador — não há dependências de build.

No **StackBlitz**, importe o repositório e utilize o preview estático (HTML/CSS/JS).

## 🛠️ Tecnologias

- HTML5, CSS3 e JavaScript puro (sem frameworks)
- [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) via CDN, para leitura das planilhas `.xlsx`/`.xls`

## 📌 Observações

O código foi organizado em seções comentadas dentro de `script.js` (leitura de arquivo, processamento, renderização, interações, impressão e cronômetro) para facilitar manutenção e leitura.
