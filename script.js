// --- SISTEMA DE LOGIN ---
const USER_CORRETO = "BibliotecaETEGP";

if (!localStorage.getItem("senha_biblioteca_atual")) {
    localStorage.setItem("senha_biblioteca_atual", "Socorro123");
}

document.getElementById('formLogin').addEventListener('submit', function(e) {
    e.preventDefault();
    const userDigitado = document.getElementById('usuario').value;
    const senhaDigitada = document.getElementById('senha').value;
    const msgErro = document.getElementById('erroLogin');
    const senhaCorreta = localStorage.getItem("senha_biblioteca_atual");

    if (userDigitado === USER_CORRETO && senhaDigitada === senhaCorreta) {
        sessionStorage.setItem("logado_etegpc", "true");
        verificarAutenticacao();
    } else {
        msgErro.innerText = "Usuário ou senha incorretos!";
    }
});

function mostrarTrocarSenha() {
    document.getElementById('caixaLogin').style.display = "none";
    document.getElementById('caixaTrocarSenha').style.display = "block";
}

function mostrarLoginOriginal() {
    document.getElementById('caixaLogin').style.display = "block";
    document.getElementById('caixaTrocarSenha').style.display = "none";
}

document.getElementById('formTrocarSenha').addEventListener('submit', function(e) {
    e.preventDefault();
    const Slider = document.getElementById('senhaAtual').value;
    const nova = document.getElementById('senhaNova').value;
    const senhaSalva = localStorage.getItem("senha_biblioteca_atual");

    if (Slider === senhaSalva) {
        localStorage.setItem("senha_biblioteca_atual", nova);
        alert("Senha alterada com sucesso!");
        mostrarLoginOriginal();
        this.reset();
    } else {
        document.getElementById('msgSenha').innerText = "Senha atual incorreta!";
    }
});

function verificarAutenticacao() {
    const estaLogado = sessionStorage.getItem("logado_etegpc");
    if (estaLogado === "true") {
        document.getElementById('telaLogin').style.display = "none";
        document.getElementById('conteudoSistema').style.display = "block";
        carregarTabelas();
    } else {
        document.getElementById('telaLogin').style.display = "flex";
        document.getElementById('conteudoSistema').style.display = "none";
    }
}

function fazerLogout() {
    sessionStorage.removeItem("logado_etegpc");
    document.getElementById('usuario').value = "";
    document.getElementById('senha').value = "";
    document.getElementById('erroLogin').innerText = "";
    verificarAutenticacao();
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
    document.getElementById('tab-' + tabId).classList.add('active');
}

function alternarLayoutGlobal() {
    const box = document.getElementById('listaEmprestimos');
    if(box.classList.contains('layout-cards')) {
        box.classList.remove('layout-cards');
        box.classList.add('layout-tabela-compacta');
    } else {
        box.classList.remove('layout-tabela-compacta');
        box.classList.add('layout-cards');
    }
    carregarTabelas();
}

// --- INTERFACE DO IMPORTADOR ---
function togglePainelImportar() {
    const painel = document.getElementById('painelImportar');
    painel.style.display = painel.style.display === "none" ? "block" : "none";
}

setTimeout(() => {
    const campoProjeto = document.getElementById('p_nome_importacao');
    if(campoProjeto) campoProjeto.style.display = 'none'; 
}, 500);

// Converte número de série do Excel (ex: 46175) em data padrão AAAA-MM-DD
function converterDataSerialExcel(serial) {
    if (!serial || isNaN(serial)) return serial;
    let numSerial = parseFloat(serial);
    let utc_days  = Math.floor(numSerial - 25569);
    let date_info = new Date(utc_days * 86400 * 1000);
    
    let ano = date_info.getUTCFullYear();
    let mes = String(date_info.getUTCMonth() + 1).padStart(2, '0');
    let dia = String(date_info.getUTCDate()).padStart(2, '0');
    
    return `${ano}-${mes}-${dia}`;
}

// Helper para formatar as datas para o banco (AAAA-MM-DD)
function formatarDataBR(dataStr) {
    if (!dataStr) return "";
    dataStr = String(dataStr).trim();
    
    if (!isNaN(dataStr) && dataStr.length >= 4) {
        return converterDataSerialExcel(dataStr);
    }

    if (dataStr.includes('/')) {
        let partes = dataStr.split('/');
        if (partes.length === 3) {
            return `${partes[2]}-${partes[1].padStart(2, '0')}-${partes[0].padStart(2, '0')}`;
        }
    }
    return dataStr;
}

// Helper para exibir na tela (DD/MM/AAAA)
function exibirDataBR(dataStr) {
    if (!dataStr) return "";
    dataStr = String(dataStr).trim();
    if (dataStr.includes('-')) {
        let partes = dataStr.split('-');
        if (partes.length === 3) {
            return `${partes[2]}/${partes[1]}/${partes[0]}`;
        }
    }
    return dataStr;
}

// --- PROCESSADOR DE IMPORTAÇÃO (LÊ PROJETO DA PLANILHA) ---
function importarTextoColado() {
    const texto = document.getElementById('inputColarExcel').value.trim();
    const arquivoInput = document.getElementById('arquivoExcel');
    const tipoDefinido = document.querySelector('input[name="tipoImportacao"]:checked').value;

    let dadosAtuais = puxarBanco();

    // Cenário A: Importação via Arquivo (.xlsx, .xls, .csv)
    if (arquivoInput && arquivoInput.files.length > 0) {
        const file = arquivoInput.files[0];
        const reader = new FileReader();

        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            const matriz = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            let contagemSucesso = 0;

            matriz.forEach((colunas, index) => {
                if (index === 0) return;

                if (tipoDefinido === 'projeto') {
                    if (colunas.length >= 4) {
                        let projRaw     = colunas[0] ? String(colunas[0]).trim() : "Geral";
                        let turmaRaw    = colunas[1] ? String(colunas[1]).trim() : "";
                        let alunoRaw    = colunas[2] ? String(colunas[2]).trim() : "";
                        let livreRaw    = colunas[3] ? String(colunas[3]).trim() : "";
                        let dataPegou   = colunas[4] ? String(colunas[4]).trim() : "";
                        let dataDevol   = colunas[5] ? String(colunas[5]).trim() : "";

                        if (alunoRaw && livreRaw) {
                            dadosAtuais.projetos.push({
                                projeto: projRaw,
                                turma: turmaRaw.toUpperCase().replace(/\s+/g, ''),
                                aluno: alunoRaw,
                                livro: livreRaw,
                                saida: formatarDataBR(dataPegou),
                                devolucao: formatarDataBR(dataDevol),
                                nota: ""
                            });
                            contagemSucesso++;
                        }
                    }
                } else {
                    if (colunas.length >= 3) {
                        let turmaRaw    = colunas[0] ? String(colunas[0]).trim() : "";
                        let alunoRaw    = colunas[1] ? String(colunas[1]).trim() : "";
                        let livroRaw    = colunas[2] ? String(colunas[2]).trim() : "";
                        let dataPegou   = colunas[3] ? String(colunas[3]).trim() : "";
                        let dataDevol   = colunas[4] ? String(colunas[4]).trim() : "";

                        if (alunoRaw && livroRaw) {
                            dadosAtuais.emprestimos.push({
                                turma: turmaRaw.toUpperCase().replace(/\s+/g, ''),
                                aluno: alunoRaw,
                                livro: livroRaw,
                                saida: formatarDataBR(dataPegou),
                                devolucao: formatarDataBR(dataDevol),
                                nota: ""
                            });
                            contagemSucesso++;
                        }
                    }
                }
            });

            localStorage.setItem(dbName, JSON.stringify(dadosAtuais));
            carregarTabelas();
            limparCamposImportacao();
            alert(`Sucesso! Planilha lida e ${contagemSucesso} registros foram importados.`);
        };

        reader.readAsArrayBuffer(file);

    } else {
        // Cenário B: Importação via Texto Colado
        if (!texto) {
            alert("Cole as linhas do Excel ou selecione um arquivo antes de processar!");
            return;
        }

        const linhas = texto.split('\n');
        let contagemSucesso = 0;

        linhas.forEach((linha) => {
            let colunas = inline_split_excel(linha);
            if (!colunas[0]) return;

            if (colunas[1] && (colunas[1].toLowerCase() === "aluno" || colunas[2] && colunas[2].toLowerCase() === "livro")) return;

            if (tipoDefinido === 'projeto') {
                if (colunas.length >= 4) {
                    let projRaw     = colunas[0].trim();
                    let turmaRaw    = colunas[1].trim();
                    let alunoRaw    = colunas[2].trim();
                    let livroRaw    = colunas[3].trim();
                    let dataPegou   = colunas[4] ? colunas[4].trim() : "";
                    let dataDevol   = colunas[5] ? colunas[5].trim() : "";

                    if (alunoRaw && livroRaw) {
                        dadosAtuais.projetos.push({
                            projeto: projRaw,
                            turma: turmaRaw.toUpperCase().replace(/\s+/g, ''),
                            aluno: alunoRaw,
                            livro: livroRaw,
                            saida: formatarDataBR(dataPegou),
                            devolucao: formatarDataBR(dataDevol),
                            nota: ""
                        });
                        contagemSucesso++;
                    }
                }
            } else {
                if (colunas.length >= 3) {
                    let turmaRaw    = colunas[0].trim();
                    let alunoRaw    = colunas[1].trim();
                    let livroRaw    = colunas[2].trim();
                    let dataPegou   = colunas[3] ? colunas[3].trim() : "";
                    let dataDevol   = colunas[4] ? colunas[4].trim() : "";

                    if (alunoRaw && livroRaw) {
                        dadosAtuais.emprestimos.push({
                            turma: turmaRaw.toUpperCase().replace(/\s+/g, ''),
                            aluno: alunoRaw,
                            livro: livroRaw,
                            saida: formatarDataBR(dataPegou),
                            devolucao: formatarDataBR(dataDevol),
                            nota: ""
                        });
                        contagemSucesso++;
                    }
                }
            }
        });

        localStorage.setItem(dbName, JSON.stringify(dadosAtuais));
        carregarTabelas();
        limparCamposImportacao();
        alert(`Sucesso! Conseguimos importar ${contagemSucesso} registros.`);
    }
}

function inline_split_excel(linha) {
    let colunas = linha.split('\t');
    if (colunas.length < 3) colunas = inline_split_excel_tab(linha);
    return colunas;
}

function inline_split_excel_tab(linha) {
    return linha.split(';');
}

function limparCamposImportacao() {
    document.getElementById('inputColarExcel').value = "";
    const campoTxtProj = document.getElementById('p_nome_importacao');
    if (campoTxtProj) campoTxtProj.value = "";
    const arquivoInput = document.getElementById('arquivoExcel');
    if(arquivoInput) arquivoInput.value = "";
    togglePainelImportar();
}

// --- BANCO DE DADOS LOCAL ---
const dbName = "BD_Biblioteca_ETEGPC";

function puxarBanco() {
    let dados = localStorage.getItem(dbName);
    if (!dados) return { emprestimos: [], projetos: [], mural: [] };
    try {
        let objeto = JSON.parse(dados);
        if (!objeto.emprestimos) objeto.emprestimos = [];
        if (!objeto.projetos) objeto.projetos = [];
        if (!objeto.mural) objeto.mural = [];
        return objeto;
    } catch (e) {
        return { emprestimos: [], projetos: [], mural: [] };
    }
}

function salvarDados(objeto, tipo) {
    let dados = puxarBanco();
    objeto.nota = ""; 
    if(tipo === 'emprestimo') dados.emprestimos.push(objeto);
    else dados.projetos.push(objeto);
    localStorage.setItem(dbName, JSON.stringify(dados));
    carregarTabelas();
}

function concluirEntrega(index, tipo) {
    let dados = puxarBanco();
    if(tipo === 'emprestimo') dados.emprestimos.splice(index, 1);
    else if(tipo === 'projeto') dados.projetos.splice(index, 1);
    else if(tipo === 'mural') dados.mural.splice(index, 1);
    localStorage.setItem(dbName, JSON.stringify(dados));
    fecharZoom();
    carregarTabelas();
}

// Corrigido para suportar o sufixo dinâmico vindo do filtro de salas
function alternarCampoNota(index, tipo, sufixo = '') {
    const idFinal = `nota-input-container-${tipo}-${index}${sufixo}`;
    const container = document.getElementById(idFinal);
    if (container) {
        container.style.display = (container.style.display === "none" || container.style.display === "") ? "block" : "none";
    }
}

// Corrigido para mapear dinamicamente os inputs de notas tanto na principal quanto nas salas
function salvarNota(index, tipo, sufixo = '') {
    let dados = puxarBanco();
    const idInput = `input-nota-${tipo}-${index}${sufixo}`;
    const textoNota = document.getElementById(idInput).value;
    
    if(tipo === 'emprestimo') dados.emprestimos[index].nota = textoNota;
    else dados.projetos[index].nota = textoNota;
    
    localStorage.setItem(dbName, JSON.stringify(dados));
    carregarTabelas();
    if(document.getElementById('modalZoom').style.display === "flex") abrirZoom(index, tipo);
}

// Corrigido para suportar o sufixo dinâmico de salas
function alternarCampoPrazo(index, tipo, sufixo = '') {
    const idFinal = `prazo-input-container-${tipo}-${index}${sufixo}`;
    const container = document.getElementById(idFinal);
    if (container) {
        container.style.display = (container.style.display === "none" || container.style.display === "") ? "block" : "none";
    }
}

// Corrigido para coletar e adiar prazos de qualquer aba sem misturar elementos
function salvarNovoPrazo(index, tipo, sufixo = '') {
    let dados = puxarBanco();
    const idInput = `input-prazo-${tipo}-${index}${sufixo}`;
    const novaData = document.getElementById(idInput).value;
    
    if (!novaData) {
        alert("Por favor, escolha uma data válida!");
        return;
    }
    if(tipo === 'emprestimo') dados.emprestimos[index].devolucao = novaData;
    else dados.projetos[index].devolucao = novaData;
    
    localStorage.setItem(dbName, JSON.stringify(dados));
    carregarTabelas();
    if(document.getElementById('modalZoom').style.display === "flex") abrirZoom(index, tipo);
}

function calcularStatus(dataDevolucaoStr) {
    if(!dataDevolucaoStr) return { classe: 'status-processo', prioridade: 2, texto: 'No prazo' };
    const hoje = new Date(); hoje.setHours(0,0,0,0);
    const devolucao = new Date(dataDevolucaoStr + 'T00:00:00');
    const diferencaTempo = devolucao - hoje;
    const diferencaDias = Math.ceil(diferencaTempo / (1000 * 60 * 60 * 24));

    if (diferencaDias <= 1) return { classe: 'status-urgente', prioridade: 1, texto: 'URGENTE / ATRASADO' };
    else if (diferencaDias > 1 && diferencaDias <= 5) return { classe: 'status-processo', prioridade: 2, texto: 'Prazo quase no fim' };
    else return { classe: 'status-inicio', prioridade: 3, texto: 'Regular' };
}

function abrirZoom(index, tipo) {
    let dados = puxarBanco();
    let e = tipo === 'emprestimo' ? dados.emprestimos[index] : dados.projetos[index];
    if(!e) return;
    const statusInfo = calcularStatus(e.devolucao);
    const container = document.getElementById('zoomConteudo');
    container.innerHTML = `
        <span class="zoom-tag-status ${statusInfo.classe}">${statusInfo.texto.toUpperCase()}</span>
        ${tipo === 'projeto' ? `<h3 style="color:var(--azul); margin:10px 0;">📁 PROJETO: ${e.projeto}</h3>` : '<h3>📖 Detalhes do Empréstimo</h3>'}
        <hr>
        <p><strong>Série/Sala:</strong> ${e.turma}</p>
        <p><strong>Nome do Aluno:</strong> ${e.aluno}</p>
        <p><strong>Livro Retirado:</strong> "${e.livro}"</p>
        <p><strong>Data de Saída:</strong> ${exibirDataBR(e.saida)}</p>
        <p><strong>Data de Devolução:</strong> ${exibirDataBR(e.devolucao)}</p>
        <div class="zoom-nota-box">
            <strong>📌 Observações:</strong>
            <p style="font-size:18px; margin: 5px 0 0 0;">${e.nota ? e.nota : '<i>Nenhuma observação.</i>'}</p>
        </div>
    `;
    document.getElementById('modalZoom').style.display = "flex";
}

function abrirZoomMural(index) {
    let dados = puxarBanco(); let m = dados.mural[index]; if(!m) return;
    const container = document.getElementById('zoomConteudo');
    container.innerHTML = `<h3>📌 Nota Fixada</h3><hr><p><strong>Assunto:</strong> ${m.titulo}</p><p><strong>Data:</strong> ${exibirDataBR(m.data)}</p><div class="zoom-nota-box" style="background:#FFFCE5; max-height:300px; overflow-y:auto; word-break:break-word;"><p style="font-size:18px; white-space:pre-wrap;">${m.comentario}</p></div>`;
    document.getElementById('modalZoom').style.display = "flex";
}

function fecharZoom() { document.getElementById('modalZoom').style.display = "none"; }

// Nova assinatura contendo 'sufixo' para diferenciar IDs duplicados nas abas de salas filtradas
function gerarCardHTML(e, index, tipo, sufixo = '') {
    return `
        <div class="card ${e.classe}">
            <div class="card-info">
                ${tipo === 'projeto' ? `<span class="badge-projeto" style="background:#e6f0fa; color:#003366; padding:2px 6px; border-radius:4px; font-size:11px; font-weight:bold; display:inline-block; margin-bottom:5px;">📁 PROJETO: ${e.projeto || 'Não Definido'}</span><br>` : ''}
                <strong>${e.turma}</strong> - ${e.aluno}<br>
                Livro: <strong>${e.livro}</strong><br>
                <small>Fim/Prazo: <strong>${exibirDataBR(e.devolucao)}</strong></small>
                ${e.nota ? `<div class="nota-exibida">📌 ${e.nota}</div>` : ''}
                
                <div id="nota-input-container-${tipo}-${index}${sufixo}" class="area-comentario">
                    <div class="input-group">
                        <input type="text" id="input-nota-${tipo}-${index}${sufixo}" placeholder="Comentário..." value="${e.nota || ''}">
                        <button onclick="salvarNota(${index}, '${tipo}', '${sufixo}')" class="btn-save-nota">Ok</button>
                    </div>
                </div>

                <div id="prazo-input-container-${tipo}-${index}${sufixo}" class="area-comentario">
                    <div class="input-group" style="align-items: center;">
                        <span style="font-size:12px; font-weight:bold;">Novo prazo:</span>
                        <input type="date" id="input-prazo-${tipo}-${index}${sufixo}" value="${e.devolucao || ''}" style="padding:4px;">
                        <button onclick="salvarNovoPrazo(${index}, '${tipo}', '${sufixo}')" class="btn-save-nota" style="background:#003366;">Salvar</button>
                    </div>
                </div>
            </div>
            <div class="card-actions">
                <button onclick="abrirZoom(${index}, '${tipo}')" class="btn-note-toggle" style="background:#6c757d;" title="Ver Detalhes">🔎</button>
                <button onclick="alternarCampoNota(${index}, '${tipo}', '${sufixo}')" class="btn-note-toggle" title="Adicionar Comentário">💬</button>
                <button onclick="alternarCampoPrazo(${index}, '${tipo}', '${sufixo}')" class="btn-note-toggle" style="background:#ffc107; color:#212529;" title="Estender Prazo">📅</button>
                <button onclick="concluirEntrega(${index}, '${tipo}')" class="btn-done" title="Concluir Devolução">✓</button>
            </div>
        </div>
    `;
}

function carregarTabelas() {
    let dados = puxarBanco();
    const termoBuscaEmp = document.getElementById('buscaEmprestimos').value.toLowerCase();
    const termoBuscaProj = document.getElementById('buscaProjetos').value.toLowerCase();
    const salaFiltrada = document.getElementById('filtroAbasSeries').value;

    let empProcessados = dados.emprestimos.map((e, index) => ({ ...e, originalIndex: index, ...calcularStatus(e.devolucao) }));
    let projProcessados = dados.projetos.map((p, index) => ({ ...p, originalIndex: index, ...calcularStatus(p.devolucao) }));

    if (termoBuscaEmp) empProcessados = empProcessados.filter(e => e.aluno.toLowerCase().includes(termoBuscaEmp) || e.livro.toLowerCase().includes(termoBuscaEmp) || e.turma.toLowerCase().includes(termoBuscaEmp));
    if (termoBuscaProj) projProcessados = projProcessados.filter(p => p.aluno.toLowerCase().includes(termoBuscaProj) || p.livro.toLowerCase().includes(termoBuscaProj) || p.projeto.toLowerCase().includes(termoBuscaProj) || (p.turma && p.turma.toLowerCase().includes(termoBuscaProj)));

    empProcessados.sort((a, b) => {
        if (a.prioridade !== b.prioridade) return a.prioridade - b.prioridade;
        return a.aluno.localeCompare(b.aluno);
    });
    
    projProcessados.sort((a, b) => {
        if (a.prioridade !== b.prioridade) return a.prioridade - b.prioridade;
        return a.aluno.localeCompare(b.aluno);
    });

    const listaEmp = document.getElementById('listaEmprestimos');
    if(listaEmp.classList.contains('layout-tabela-compacta')) {
        listaEmp.innerHTML = `
            <table style="width:100%; border-collapse:collapse; background:white;">
                <thead>
                    <tr style="background:var(--azul); color:white; font-size:14px;">
                        <th style="padding:8px;">Turma</th>
                        <th style="padding:8px;">Aluno</th>
                        <th style="padding:8px;">Livro</th>
                        <th style="padding:8px;">Prazo</th>
                        <th style="padding:8px;">Ações</th>
                    </tr>
                </thead>
                <tbody>
                    ${empProcessados.map(e => `
                        <tr class="${e.classe}">
                            <td style="padding:8px; font-weight:bold;" onclick="abrirZoom(${e.originalIndex}, 'emprestimo')">${e.turma}</td>
                            <td style="padding:8px;" onclick="abrirZoom(${e.originalIndex}, 'emprestimo')">${e.aluno}</td>
                            <td style="padding:8px;" onclick="abrirZoom(${e.originalIndex}, 'emprestimo')">${e.livro}</td>
                            <td style="padding:8px;" onclick="abrirZoom(${e.originalIndex}, 'emprestimo')">${exibirDataBR(e.devolucao)} 🔎</td>
                            <td style="padding:8px; text-align:center;">
                                <button onclick="alternarCampoPrazo(${e.originalIndex}, 'emprestimo', '-tabela')" style="border:none; background:none; cursor:pointer; font-size:16px;">📅</button>
                            </td>
                        </tr>
                        <tr id="prazo-input-container-emprestimo-${e.originalIndex}-tabela" style="display:none; background:#fffde6;">
                            <td colspan="5" style="padding:8px; text-align:right;">
                                <span style="font-size:12px; font-weight:bold;">Adiar Prazo:</span>
                                <input type="date" id="input-prazo-emprestimo-${e.originalIndex}-tabela" value="${e.devolucao}" style="width:auto; display:inline-block; padding:3px;">
                                <button onclick="salvarNovoPrazo(${e.originalIndex}, 'emprestimo', '-tabela')" style="background:var(--azul); color:white; border:none; padding:4px 8px; border-radius:3px; cursor:pointer;">Mudar</button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>`;
    } else {
        listaEmp.innerHTML = empProcessados.map(e => gerarCardHTML(e, e.originalIndex, 'emprestimo', '-principal')).join('');
    }

    document.getElementById('listaProjetos').innerHTML = projProcessados.map(p => gerarCardHTML(p, p.originalIndex, 'projeto', '-principal')).join('');

    const listaSeriesEmp = document.getElementById('listaSeriesEmprestimos');
    const listaSeriesProj = document.getElementById('listaSeriesProjetos');
    
    if (salaFiltrada === "TODAS") {
        listaSeriesEmp.innerHTML = "<p style='text-align:center; color:#666;'>Selecione uma sala para filtrar.</p>";
        listaSeriesProj.innerHTML = "<p style='text-align:center; color:#666;'>Selecione uma sala para filtrar.</p>";
    } else {
        let empFiltradosSala = empProcessados.filter(e => e.turma === salaFiltrada);
        let projFiltradosSala = projProcessados.filter(p => p.turma === salaFiltrada);
        // Enviando o sufixo '-sala' para diferenciar nos cliques dos cartões filtrados!
        listaSeriesEmp.innerHTML = empFiltradosSala.length ? empFiltradosSala.map(e => gerarCardHTML(e, e.originalIndex, 'emprestimo', '-sala')).join('') : "<p style='text-align:center; color:#666;'>Nada ativo.</p>";
        listaSeriesProj.innerHTML = projFiltradosSala.length ? projFiltradosSala.map(p => gerarCardHTML(p, p.originalIndex, 'projeto', '-sala')).join('') : "<p style='text-align:center; color:#666;'>Nada ativo.</p>";
    }

    document.getElementById('quadroMuralAvisos').innerHTML = dados.mural.map((m, index) => `<div class="bloco-nota" onclick="abrirZoomMural(${index})"><div class="bloco-nota-header"><strong>📌 ${m.titulo}</strong><small>${exibirDataBR(m.data)}</small></div><p>${m.comentario}</p><button onclick="event.stopPropagation(); concluirEntrega(${index}, 'mural')" style="background:#ff4d4d; color:white; border:none; padding:5px; border-radius:4px; cursor:pointer; font-size:12px; margin-top:10px;">Remover</button></div>`).join('');
}

document.getElementById('formEmprestimo').addEventListener('submit', function(e) {
    e.preventDefault();
    salvarDados({ turma: document.getElementById('turma').value, aluno: document.getElementById('aluno').value, livro: document.getElementById('livro').value, saida: document.getElementById('dataPegou').value, devolucao: document.getElementById('dataDevolver').value }, 'emprestimo');
    this.reset();
});

document.getElementById('formProjeto').addEventListener('submit', function(e) {
    e.preventDefault();
    salvarDados({ aluno: document.getElementById('p_aluno').value, turma: document.getElementById('p_turma').value, livro: document.getElementById('p_livro').value, projeto: document.getElementById('p_nome').value, saida: document.getElementById('p_dataPegou').value, devolucao: document.getElementById('p_dataDevolver').value }, 'projeto');
    this.reset();
});

document.getElementById('formMural').addEventListener('submit', function(e) {
    e.preventDefault();
    let dados = puxarBanco();
    dados.mural.push({ titulo: document.getElementById('m_titulo').value, data: document.getElementById('m_data').value, comentario: document.getElementById('m_comentario').value });
    localStorage.setItem(dbName, JSON.stringify(dados));
    carregarTabelas();
    this.reset();
    alert("Fixado!");
});

window.onload = verificarAutenticacao;