/* ==========================================================================
   Painel de Triagem Automatizada de Paletes
   script.js — toda a lógica da aplicação

   Seções:
   1. Estado global
   2. Upload e leitura do arquivo (XLSX/XLS/XML)
   3. Processamento das linhas (palete + material)
   4. Renderização da tabela
   5. Interações do usuário (copiar, concluir, limpar)
   6. Impressão das placas de bloqueio
   7. Cronômetro de operação
   ========================================================================== */

/* --------------------------------------------------------------------------
   1. Estado global
   -------------------------------------------------------------------------- */

let todosOsItens = [];

// Lista pré-definida de fábrica: todos os códigos de bloqueio configurados
const LISTA_BLOQUEADOS = [
    "3150697", "1900169", "3231166", "3231133", "3021287",
    "3150622", "3230931", "3231628", "3320223", "3230432",
    "3150698", "3231494", "3001184", "3231571", "3231462",
    "3231443", "3231181", "3231152", "3231103", "3231426",
    "3150642", "3231132", "3231430", "3231200",
    "3001043", "3001529", "3231158", "3231434", "3231162",
    "3000726", "3001614", "3001527", "3231152", "3001176",
    "3231430", "3231181", "3231613", "3021631",
];

// Variáveis do cronômetro de operação
let intervaloCronometro = null;
let segundosOperacao = 0;

/* --------------------------------------------------------------------------
   2. Upload e leitura do arquivo (XLSX/XLS/XML)
   -------------------------------------------------------------------------- */

document.getElementById('input-xml').addEventListener('change', function (e) {
    let arquivos = e.target.files;
    if (!arquivos || arquivos.length === 0) return;

    let arquivo = arquivos[0];
    document.getElementById('nome-arquivo').innerText = `Arquivo carregado: ${arquivo.name}`;

    let extensao = arquivo.name.split('.').pop().toLowerCase();
    todosOsItens = [];

    if (extensao === "xlsx" || extensao === "xls") {
        // Leitura binária necessária para arquivos .xlsx/.xls (formato compactado)
        let leitor = new FileReader();
        leitor.onload = function (evento) {
            let dadosBinarios = evento.target.result;
            let workbook = XLSX.read(dadosBinarios, { type: 'array' });
            processarWorkbookXlsx(workbook);
        };
        leitor.readAsArrayBuffer(arquivo);
    } else if (extensao === "xml") {
        // Leitura de texto para o formato XML antigo do Excel (SpreadsheetML)
        let leitor = new FileReader();
        leitor.onload = function (evento) {
            let textoXml = evento.target.result;
            processarTextoXmlSemFiltros(textoXml);
        };
        leitor.readAsText(arquivo, 'UTF-8');
    } else {
        alert("Formato de arquivo não suportado. Selecione um arquivo .xlsx, .xls ou .xml.");
    }
});

/* --------------------------------------------------------------------------
   3. Processamento das linhas (palete + material)
   -------------------------------------------------------------------------- */

// Função compartilhada: varre os valores de uma linha (independente da origem ser
// XML ou XLSX) procurando o marcador de palete e os códigos de material válidos.
// "estado" carrega o palete/número atuais e é atualizado por referência entre linhas.
function processarValoresCelulas(valoresCelulas, estado) {
    for (let c = 0; c < valoresCelulas.length; c++) {
        let token = valoresCelulas[c];
        if (!token) continue;

        if (token.toLowerCase().includes("pallet") || token.toLowerCase().includes("palete")) {
            if (valoresCelulas[c + 1]) {
                let numTexto = valoresCelulas[c + 1];
                estado.paleteNo = "Palete " + numTexto;
                estado.numeroPaleteAba = parseInt(numTexto.replace(/\D/g, "")) || estado.numeroPaleteAba;
            }
        }

        if (token.length >= 4 && !isNaN(token) && !token.includes(".")) {
            let desc = valoresCelulas[c + 1] || "---";
            let qtdTexto = valoresCelulas[c + 2] || "0";
            let qtd = parseInt(qtdTexto) || 0;

            if (qtd > 0 && token !== String(qtd)) {
                todosOsItens.push({
                    palete: estado.paleteNo,
                    numPalete: estado.numeroPaleteAba,
                    material: token,
                    descricao: desc,
                    quantidade: qtd,
                    concluido: false,
                    ordemOriginal: todosOsItens.length
                });
                break;
            }
        }
    }
}

// Após processar todas as abas (seja XML ou XLSX), decide se mostra o painel
function finalizarProcessamento(origem) {
    if (todosOsItens.length > 0) {
        reordenarTabela();
        document.getElementById("painel-operacional").style.display = "block";
        iniciarCronometro(); // Liga o tempo de operação
    } else {
        alert(`Aviso: O arquivo ${origem} foi lido com sucesso, mas a estrutura interna não foi processada.`);
    }
}

function processarTextoXmlSemFiltros(textoCompleto) {
    let blocosAbas = textoCompleto.split(/<Worksheet/i);
    blocosAbas.shift();

    blocosAbas.forEach(function (blocoTextual, indiceAba) {
        let nomeMatch = blocoTextual.match(/ss:Name\s*=\s*"([^"]+)"/i) || blocoTextual.match(/Name\s*=\s*"([^"]+)"/i);
        let nomeAba = nomeMatch ? nomeMatch[1] : "Palete_" + (indiceAba + 1);

        let estado = {
            numeroPaleteAba: parseInt(nomeAba.replace(/\D/g, "")) || (indiceAba + 1)
        };
        estado.paleteNo = "Palete " + estado.numeroPaleteAba;

        let linhas = blocoTextual.split(/<Row/i);
        linhas.shift();

        linhas.forEach(function (linhaTexto) {
            let dadosMatch = linhaTexto.match(/<Data[^>]*>([\s\S]*?)<\/Data>/gi);
            if (!dadosMatch) return;

            let valoresCelulas = dadosMatch.map(function (tagData) {
                return tagData.replace(/<[^>]*>/g, "").trim();
            });

            processarValoresCelulas(valoresCelulas, estado);
        });
    });

    finalizarProcessamento("XML");
}

// Processa um workbook .xlsx/.xls lido pela biblioteca SheetJS, reaproveitando
// a mesma lógica de detecção de palete/material usada para o XML.
function processarWorkbookXlsx(workbook) {
    workbook.SheetNames.forEach(function (nomeAba, indiceAba) {
        let planilha = workbook.Sheets[nomeAba];

        // header: 1 => cada linha vira um array de células, igual ao formato do XML
        let matrizLinhas = XLSX.utils.sheet_to_json(planilha, { header: 1, defval: "", raw: false });

        let estado = {
            numeroPaleteAba: parseInt(nomeAba.replace(/\D/g, "")) || (indiceAba + 1)
        };
        estado.paleteNo = "Palete " + estado.numeroPaleteAba;

        matrizLinhas.forEach(function (linha) {
            if (!linha || linha.length === 0) return;

            let valoresCelulas = linha.map(function (celula) {
                return String(celula).trim();
            });

            processarValoresCelulas(valoresCelulas, estado);
        });
    });

    finalizarProcessamento("Excel");
}

function reordenarTabela() {
    todosOsItens.sort(function (a, b) {
        if (a.numPalete !== b.numPalete) {
            return a.numPalete - b.numPalete;
        }
        return a.ordemOriginal - b.ordemOriginal;
    });
    atualizarTela();
}

/* --------------------------------------------------------------------------
   4. Renderização da tabela
   -------------------------------------------------------------------------- */

function atualizarTela() {
    let tbodyBloqueados = document.getElementById("corpo-tabela-bloqueados");
    let tbodyLiberados = document.getElementById("corpo-tabela-liberados");
    let pendentes = 0;
    let concluidos = 0;
    let htmlBloqueados = "";
    let htmlLiberados = "";

    // Controla a troca de palete dentro de cada tabela, para desenhar a linha
    // divisória visual sempre que um novo palete começar na lista.
    let ultimoPaleteBloqueados = null;
    let ultimoPaleteLiberados = null;

    todosOsItens.forEach(function (item, index) {
        let classeConcluido = item.concluido ? "concluido" : "";
        let checkIcon = item.concluido ? "✅" : "⬜";

        if (item.concluido) { concluidos++; } else { pendentes++; }

        let estaBloqueado = LISTA_BLOQUEADOS.includes(item.material);

        // Classe de alerta para os materiais retidos (contraste reforçado via CSS)
        let classeBloqueio = (estaBloqueado && !item.concluido) ? "linha-bloqueada" : "";

        // Ícone fixo de bloqueio: garante a leitura da informação mesmo sem depender só da cor
        let iconeBloqueio = estaBloqueado ? '<span class="icone-bloqueio" title="Material bloqueado">⛔</span>' : "";

        // Marca a primeira linha de cada novo palete dentro desta tabela específica
        let classeNovoPalete = "";
        if (estaBloqueado) {
            if (ultimoPaleteBloqueados !== null && ultimoPaleteBloqueados !== item.numPalete) {
                classeNovoPalete = "novo-palete";
            }
            ultimoPaleteBloqueados = item.numPalete;
        } else {
            if (ultimoPaleteLiberados !== null && ultimoPaleteLiberados !== item.numPalete) {
                classeNovoPalete = "novo-palete";
            }
            ultimoPaleteLiberados = item.numPalete;
        }

        let linhaHtml = `
            <tr class="linha-item ${classeConcluido} ${classeBloqueio} ${classeNovoPalete}">
                <td style="text-align: center; font-size: 16px;" onclick="alternarStatus(${index})">${checkIcon}</td>
                <td onclick="alternarStatus(${index})"><span class="badge-palete">${item.palete}</span></td>
                <td>
                    <div class="bloco-copiar">
                        <button class="btn-copiar" onclick="copiarCodigoDireto('${item.material}', ${index}, event)">📋</button>
                        ${iconeBloqueio}<strong>${item.material}</strong>
                    </div>
                </td>
                <td onclick="alternarStatus(${index})">${item.descricao}</td>
                <td onclick="alternarStatus(${index})"><strong>${item.quantidade}</strong> un.</td>
            </tr>`;

        // Cada item cai só numa das duas tabelas, mantendo a ordem original de palete
        if (estaBloqueado) {
            htmlBloqueados += linhaHtml;
        } else {
            htmlLiberados += linhaHtml;
        }
    });

    tbodyBloqueados.innerHTML = htmlBloqueados || `<tr><td colspan="5" style="text-align: center; color: #64748b; padding: 18px;">Nenhum material bloqueado neste arquivo.</td></tr>`;
    tbodyLiberados.innerHTML = htmlLiberados || `<tr><td colspan="5" style="text-align: center; color: #64748b; padding: 18px;">Nenhum item liberado.</td></tr>`;

    document.getElementById("qtd-pendentes").innerText = pendentes;
    document.getElementById("qtd-concluidos").innerText = concluidos;

    // Estado comemorativo: destaca a caixa de concluídos quando não sobra nenhum item pendente
    let caixaConcluidos = document.getElementById("caixa-concluidos");
    let textoConcluidos = document.getElementById("texto-concluidos");
    let tudoConcluido = todosOsItens.length > 0 && pendentes === 0;

    caixaConcluidos.classList.toggle("tudo-concluido", tudoConcluido);
    textoConcluidos.innerText = tudoConcluido ? "🎉 Tudo Concluído:" : "✅ Itens Concluídos:";
}

/* --------------------------------------------------------------------------
   5. Interações do usuário (copiar, concluir, limpar)
   -------------------------------------------------------------------------- */

function copiarCodigoDireto(codigo, index, evento) {
    evento.stopPropagation(); // Mantido para gerenciar a ordem dos cliques nativos

    let caixaTexto = document.createElement("textarea");
    caixaTexto.value = codigo;
    document.body.appendChild(caixaTexto);
    caixaTexto.select();

    try {
        document.execCommand("copy"); // Executa o comando clássico de cópia

        let b = evento.target;
        b.innerText = "⚡";
        b.style.background = "#2ecc71";
        b.style.color = "white";

        setTimeout(function () {
            b.innerText = "📋";
            b.style.background = "#f1f5f9";
            b.style.color = "#333";
        }, 600);

        // Usa o índice real do item no array todosOsItens (recebido como parâmetro),
        // em vez de tentar redescobrir a posição pela linha da tabela na tela.
        if (typeof index === "number" && todosOsItens[index]) {
            todosOsItens[index].concluido = !todosOsItens[index].concluido;
            atualizarTela();
        }

    } catch (err) {
        alert("Erro ao copiar.");
    }

    document.body.removeChild(caixaTexto);
}

function alternarStatus(index) {
    todosOsItens[index].concluido = !todosOsItens[index].concluido;
    atualizarTela();
}

function limparPainel() {
    todosOsItens = [];
    document.getElementById('input-xml').value = "";
    document.getElementById('nome-arquivo').innerText = "";
    document.getElementById("painel-operacional").style.display = "none";

    clearInterval(intervaloCronometro); // Para o cronômetro ao limpar o painel

    atualizarTela();
}

function escapeHtml(v) {
    return String(v === undefined || v === null ? "" : v)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/* --------------------------------------------------------------------------
   6. Impressão das placas de bloqueio
   -------------------------------------------------------------------------- */

// Gera e imprime uma placa grande de identificação para cada material
// bloqueado presente na tabela atual.
function imprimirMateriaisBloqueados() {
    let itensBloqueados = todosOsItens.filter(function (item) {
        return LISTA_BLOQUEADOS.includes(item.material);
    });

    if (itensBloqueados.length === 0) {
        alert("Não há materiais bloqueados para imprimir.");
        return;
    }

    const now = new Date();
    const dataHora = now.toLocaleString('pt-BR');

    const cards = itensBloqueados.map(function (item, pageIndex) {
        return `
          <section class="plate">
            <div class="plate-top">⚠️ ATENÇÃO</div>
            <div class="plate-title">MATERIAL BLOQUEADO</div>
            <div class="plate-rule"></div>
            <div class="label">CÓDIGO DO MATERIAL</div>
            <div class="code">${escapeHtml(item.material)}</div>
            <div class="label small-label">DESCRIÇÃO</div>
            <div class="desc">${escapeHtml(item.descricao)}</div>
            <div class="info-grid">
              <div><span>QUANTIDADE</span><b>${escapeHtml(item.quantidade)} un.</b></div>
              <div><span>PALETE</span><b>${escapeHtml(item.palete)}</b></div>
            </div>
            <div class="reason-box"><div class="reason-label">MOTIVO DO BLOQUEIO</div><div class="reason">Material identificado na lista de bloqueio de fábrica.</div></div>
            <div class="warning">NÃO UTILIZAR &nbsp; • &nbsp; NÃO LIBERAR &nbsp; • &nbsp; NÃO MOVIMENTAR</div>
            <div class="footer-note">AGUARDAR ANÁLISE / LIBERAÇÃO DO TIME RESPONSÁVEL</div>
            <div class="print-meta">Placa ${pageIndex + 1} de ${itensBloqueados.length} • Gerada em ${escapeHtml(dataHora)}</div>
          </section>`;
    }).join('');

    const printWin = window.open('', '_blank');
    if (!printWin) {
        alert("O navegador bloqueou a janela de impressão. Permita pop-ups para esta página e tente novamente.");
        return;
    }
    printWin.document.open();
    printWin.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>Placas de Bloqueio</title>
        <style>
          @page{size:A4 portrait;margin:8mm;}
          *{box-sizing:border-box;}
          html,body{margin:0;padding:0;background:#fff;font-family:Arial,"Segoe UI",sans-serif;color:#111;}
          .plate{width:100%;min-height:calc(297mm - 16mm);page-break-after:always;border:5px solid #991b1b;border-radius:8px;display:flex;flex-direction:column;align-items:center;padding:15mm 13mm 10mm;text-align:center;position:relative;overflow:hidden;}
          .plate:last-child{page-break-after:auto;}
          .plate-top{font-size:28pt;font-weight:900;color:#991b1b;letter-spacing:2px;margin-bottom:7mm;}
          .plate-title{font-size:34pt;font-weight:1000;color:#991b1b;line-height:1.05;}
          .plate-rule{width:88%;height:3px;background:#111;margin:7mm 0;}
          .label{font-size:15pt;font-weight:800;letter-spacing:1px;margin-top:2mm;}
          .code{font-size:42pt;font-weight:1000;line-height:1.05;margin:4mm 0 5mm;word-break:break-all;}
          .small-label{margin-top:0;}
          .desc{font-size:22pt;font-weight:800;line-height:1.15;max-width:95%;margin:2mm 0 6mm;}
          .info-grid{width:96%;display:grid;grid-template-columns:1fr 1fr;gap:3mm;margin:1mm 0 6mm;}
          .info-grid>div{border:2px solid #333;border-radius:5px;padding:3mm;display:flex;flex-direction:column;gap:1.5mm;min-height:18mm;justify-content:center;}
          .info-grid span{font-size:10pt;font-weight:800;letter-spacing:.8px;}
          .info-grid b{font-size:17pt;line-height:1.1;word-break:break-word;}
          .reason-box{width:96%;border:3px solid #111;border-radius:6px;padding:5mm;margin-top:2mm;}
          .reason-label{font-size:13pt;font-weight:900;margin-bottom:3mm;}
          .reason{font-size:18pt;font-weight:700;line-height:1.2;word-break:break-word;}
          .warning{width:96%;margin-top:7mm;background:#991b1b;color:#fff;font-size:17pt;font-weight:1000;padding:5mm 3mm;border-radius:5px;}
          .footer-note{font-size:13pt;font-weight:900;margin-top:5mm;}
          .print-meta{position:absolute;bottom:3mm;left:0;right:0;font-size:8pt;color:#666;}
          @media print{.plate{break-after:page;}.plate:last-child{break-after:auto;}}
        </style></head><body>${cards}<script>window.onload=function(){setTimeout(function(){window.print();},250);};<\/script></body></html>`);
    printWin.document.close();
}

/* --------------------------------------------------------------------------
   7. Cronômetro de operação
   -------------------------------------------------------------------------- */

function iniciarCronometro() {
    clearInterval(intervaloCronometro);
    segundosOperacao = 0;
    document.getElementById("cronometro").innerText = "00:00:00";

    intervaloCronometro = setInterval(function () {
        segundosOperacao++;
        let horas = Math.floor(segundosOperacao / 3600);
        let minutos = Math.floor((segundosOperacao % 3600) / 60);
        let segundos = segundosOperacao % 60;

        let hFormatada = horas < 10 ? "0" + horas : horas;
        let mFormatada = minutos < 10 ? "0" + minutos : minutos;
        let sFormatada = segundos < 10 ? "0" + segundos : segundos;

        document.getElementById("cronometro").innerText = hFormatada + ":" + mFormatada + ":" + sFormatada;
    }, 1000);
}
