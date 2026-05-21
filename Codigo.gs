/**
 * Painel Bombeira de Garra - XI ENBOM 2026
 * Backend Google Apps Script
 *
 * Vincular este script à planilha (Extensões → Apps Script).
 * Implantar como Web App (Implantar → Nova implantação):
 *   - Executar como: Eu
 *   - Acesso: Qualquer pessoa
 */

const NOMES_ABAS = {
  EQUIPES: 'Equipes',
  PARTICIPANTES: 'Participantes',
  RESULTADOS: 'Resultados',
  PENALIDADES: 'Penalidades',
  REF_PENALIDADES: 'RefPenalidades',
  REF_PAISES: 'RefPaises',
  REF_DSQ: 'RefDSQ',
  CONFIG: 'Config',
};

const CHAVE_CACHE = 'dados_painel_v1';
const TTL_CACHE_SEG = 21600; // 6h — invalidado manualmente pelo gestor

// ============================================================
// MENU CUSTOMIZADO NA PLANILHA
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Painel ENBOM')
    .addItem('🏆 Gerenciar competição', 'abrirGerenciador')
    .addItem('📢 Atualizar painel público', 'atualizarPainelPublico')
    .addItem('🔗 Mostrar URL do painel', 'mostrarUrlWebApp')
    .addSeparator()
    .addItem('⚙️ Ativar abertura automática da sidebar', 'configurarAberturaAutomatica')
    .addItem('🧪 Testar leitura de dados', 'testarLeitura')
    .addToUi();
}

/**
 * Disparado pelo installable trigger (criado via `configurarAberturaAutomatica`
 * no Setup.gs). Roda com autorização completa, podendo chamar showSidebar().
 */
function aoAbrirPlanilha() {
  try {
    abrirGerenciador();
  } catch (e) {
    // Falha silenciosamente — usuário pode abrir manualmente pelo menu
  }
}

function atualizarPainelPublico() {
  CacheService.getScriptCache().remove(CHAVE_CACHE);
  const agora = new Date();
  gravarConfig('cache_version', agora.toISOString());
  const horaStr = Utilities.formatDate(
    agora,
    Session.getScriptTimeZone() || 'America/Rio_Branco',
    'HH:mm:ss'
  );
  try {
    SpreadsheetApp.getActive().toast(
      'Painel público atualizado em ' + horaStr,
      '✅ Pronto',
      5
    );
  } catch (e) { /* ignora quando chamado via sidebar */ }
  return { sucesso: true, hora: horaStr };
}

function mostrarUrlWebApp() {
  try {
    const cfg = lerConfig(SpreadsheetApp.getActive());
    const urlPublica = String(cfg.url_painel_publico || '').trim();
    const urlAppsScript = ScriptApp.getService().getUrl();
    const url = urlPublica || urlAppsScript;
    const origem = urlPublica
      ? '<p style="color:#059669;font-size:13px;margin:8px 0 0">✅ URL personalizada (GitHub Pages) — preenchida na aba <b>Config</b>, chave <code>url_painel_publico</code>.</p>'
      : '<p style="color:#92400e;font-size:13px;margin:8px 0 0">⚠️ URL do Apps Script (padrão). Para usar o painel rápido no GitHub Pages, preencha <code>url_painel_publico</code> na aba <b>Config</b>.</p>';
    const html = HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:16px">' +
        '<h3>URL do painel público</h3>' +
        '<p>Compartilhe este link ou gere um QR Code:</p>' +
        '<input type="text" value="' + url + '" readonly style="width:100%;padding:8px;font-family:monospace" onclick="this.select()">' +
        origem +
        '<p style="margin-top:16px"><a href="https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' +
          encodeURIComponent(url) + '" target="_blank">📱 Gerar QR Code</a></p>' +
      '</div>'
    ).setWidth(500).setHeight(320);
    SpreadsheetApp.getUi().showModalDialog(html, 'URL do Painel');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Implante o Web App primeiro (Implantar → Nova implantação).');
  }
}

function testarLeitura() {
  const dados = getDadosPainel();
  Logger.log(JSON.stringify(dados, null, 2));
  SpreadsheetApp.getActive().toast(
    'Equipes: ' + dados.equipes.length +
    ' | Concluídas: ' + dados.resultados.concluidas.length +
    ' | Em prova: ' + dados.resultados.emProva.length,
    'Leitura OK',
    5
  );
}

// ============================================================
// SIDEBAR — GERENCIADOR DA COMPETIÇÃO
// ============================================================

function abrirGerenciador() {
  const html = HtmlService.createTemplateFromFile('GerenciarEquipes')
    .evaluate()
    .setTitle('Gerenciador ENBOM')
    .setWidth(1080);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Alias mantido para compatibilidade
function abrirGerenciadorEquipes() { abrirGerenciador(); }

function listarEquipesCompletas() {
  const ss = SpreadsheetApp.getActive();
  const paisesMap = lerPaises(ss);
  const equipes = lerEquipes(ss, paisesMap);
  const participantes = lerParticipantes(ss);
  vincularParticipantes(equipes, participantes);

  // Anexa o resultado de cada equipe (status, tempo, pista) — útil para a
  // listagem da sidebar mostrar pills de status sem precisar de chamadas extras.
  const resultados = lerResultados(ss);
  const resultadoPorEquipe = {};
  resultados.forEach(function (r) { resultadoPorEquipe[r.equipe] = r; });
  equipes.forEach(function (e) {
    e.resultado = resultadoPorEquipe[e.nome] || null;
  });

  return equipes;
}

function listarPaisesParaForm() {
  const mapa = lerPaises(SpreadsheetApp.getActive());
  return Object.keys(mapa).map(function (p) {
    return { pais: p, bandeira: mapa[p] };
  });
}

/**
 * Salva (insere ou atualiza) uma equipe + suas participantes em transação.
 * payload = {
 *   nomeOriginal: string vazia para insert, ou nome atual para update
 *   equipe: { nome, pais, corporacao, unidadeSubnacional, bandeira, capita, observacoes }
 *   participantes: [{ nome, nomeDeGuerra, postoGraduacao, ehCapita }, ...]
 * }
 */
function salvarEquipe(payload) {
  try {
    const ss = SpreadsheetApp.getActive();
    const eq = payload.equipe || {};
    const parts = payload.participantes || [];
    const nomeOriginal = String(payload.nomeOriginal || '').trim();
    const nomeNovo = String(eq.nome || '').trim();
    const isUpdate = nomeOriginal !== '';

    // Validações
    if (!nomeNovo) return { sucesso: false, erro: 'Nome da equipe é obrigatório.' };
    if (!eq.pais) return { sucesso: false, erro: 'Selecione o país.' };
    if (parts.length === 0) return { sucesso: false, erro: 'Adicione ao menos uma participante.' };
    if (parts.length > 3) return { sucesso: false, erro: 'O edital permite no máximo 3 bombeiras por equipe.' };
    const qtdCapitas = parts.filter(function (p) { return p.ehCapita; }).length;
    if (qtdCapitas > 1) return { sucesso: false, erro: 'Marque apenas uma capitã.' };
    for (let i = 0; i < parts.length; i++) {
      if (!String(parts[i].nome || '').trim()) {
        return { sucesso: false, erro: 'Nome da participante #' + (i + 1) + ' é obrigatório.' };
      }
    }

    // Em update, mantém o nome (campo readonly no frontend). Em insert, checa duplicidade.
    const abaEquipes = ss.getSheetByName(NOMES_ABAS.EQUIPES);
    if (!abaEquipes) return { sucesso: false, erro: 'Aba "Equipes" não encontrada.' };
    const valoresEq = abaEquipes.getDataRange().getValues();
    let linhaExistente = -1;
    for (let i = 1; i < valoresEq.length; i++) {
      const nomeLinha = String(valoresEq[i][0] || '').trim();
      if (isUpdate && nomeLinha === nomeOriginal) {
        linhaExistente = i + 1;
        break;
      }
      if (!isUpdate && nomeLinha.toLowerCase() === nomeNovo.toLowerCase()) {
        return { sucesso: false, erro: 'Já existe equipe com o nome "' + nomeNovo + '".' };
      }
    }

    // Monta a linha (ordem das colunas da aba Equipes):
    // nome_equipe | pais | corporacao | unidade_subnacional | bandeira_emoji | capita | observacoes | ordem_prova
    const nomeFinal = isUpdate ? nomeOriginal : nomeNovo;
    const ordemProvaVal = (eq.ordemProva !== '' && eq.ordemProva !== null && eq.ordemProva !== undefined && !isNaN(Number(eq.ordemProva)))
      ? Number(eq.ordemProva) : '';
    const linhaEquipe = [
      nomeFinal,
      String(eq.pais || ''),
      String(eq.corporacao || ''),
      String(eq.unidadeSubnacional || ''),
      String(eq.bandeira || ''),
      String(eq.capita || ''),
      String(eq.observacoes || ''),
      ordemProvaVal,
    ];

    if (linhaExistente > 0) {
      abaEquipes.getRange(linhaExistente, 1, 1, linhaEquipe.length).setValues([linhaEquipe]);
    } else {
      const proximaLinhaEq = encontrarProximaLinhaVazia(abaEquipes, 1);
      abaEquipes.getRange(proximaLinhaEq, 1, 1, linhaEquipe.length).setValues([linhaEquipe]);
    }

    // Garante que a equipe esteja persistida antes de adicionar participantes
    // (validação de "equipe" na coluna F de Participantes referencia Equipes!A2:A1000)
    SpreadsheetApp.flush();

    // Sincroniza Participantes: remove todas as linhas da equipe e re-insere via setValues em lote.
    const abaParts = ss.getSheetByName(NOMES_ABAS.PARTICIPANTES);
    if (!abaParts) return { sucesso: false, erro: 'Aba "Participantes" não encontrada.' };

    // Localiza linhas a deletar olhando APENAS coluna F (equipe), não getLastRow,
    // porque checkboxes ou validações em outras colunas podem inflar getLastRow.
    const ultimaLinhaPart = encontrarUltimaLinhaComConteudo(abaParts, 1);
    if (ultimaLinhaPart >= 2) {
      const valoresF = abaParts.getRange(2, 6, ultimaLinhaPart - 1, 1).getValues();
      for (let i = valoresF.length - 1; i >= 0; i--) {
        const equipeDaLinha = String(valoresF[i][0] || '').trim();
        if (equipeDaLinha === nomeFinal || (isUpdate && equipeDaLinha === nomeOriginal)) {
          abaParts.deleteRow(i + 2);
        }
      }
    }

    SpreadsheetApp.flush();

    // Insere as novas participantes em lote.
    // Ordem: nome | nome_de_guerra | posto_graduacao | pais | corporacao | equipe | eh_capita
    if (parts.length > 0) {
      const matrizParts = parts.map(function (p) {
        return [
          String(p.nome || '').trim(),
          String(p.nomeDeGuerra || '').trim(),
          String(p.postoGraduacao || '').trim(),
          String(eq.pais || ''),
          String(eq.corporacao || ''),
          nomeFinal,
          p.ehCapita === true,
        ];
      });
      const proxLinha = encontrarProximaLinhaVazia(abaParts, 1);
      abaParts.getRange(proxLinha, 1, matrizParts.length, 7).setValues(matrizParts);
      // Aplica checkbox no eh_capita das linhas recém-inseridas (UX no Sheets)
      abaParts.getRange(proxLinha, 7, matrizParts.length, 1).insertCheckboxes();
    }

    // Invalida cache do painel público
    CacheService.getScriptCache().remove(CHAVE_CACHE);
    gravarConfig('cache_version', new Date().toISOString());

    return { sucesso: true, mensagem: isUpdate ? 'Equipe atualizada' : 'Equipe cadastrada' };
  } catch (e) {
    return { sucesso: false, erro: 'Erro inesperado: ' + (e && e.message ? e.message : String(e)) };
  }
}

/**
 * Exclui uma equipe e todas as suas participantes.
 * Bloqueia se houver registros em Resultados ou Penalidades referenciando.
 */
function excluirEquipe(nomeEquipe) {
  try {
    const nome = String(nomeEquipe || '').trim();
    if (!nome) return { sucesso: false, erro: 'Nome da equipe inválido.' };
    const ss = SpreadsheetApp.getActive();

    // Cascade delete: remove resultado, penalidades, participantes e a equipe.
    const abaResultados = ss.getSheetByName(NOMES_ABAS.RESULTADOS);
    if (abaResultados && abaResultados.getLastRow() > 1) {
      const valores = abaResultados.getRange(2, 1, abaResultados.getLastRow() - 1, 1).getValues();
      for (let i = valores.length - 1; i >= 0; i--) {
        if (String(valores[i][0] || '').trim() === nome) abaResultados.deleteRow(i + 2);
      }
    }

    const abaPenalidades = ss.getSheetByName(NOMES_ABAS.PENALIDADES);
    if (abaPenalidades && abaPenalidades.getLastRow() > 1) {
      const valores = abaPenalidades.getRange(2, 1, abaPenalidades.getLastRow() - 1, 1).getValues();
      for (let i = valores.length - 1; i >= 0; i--) {
        if (String(valores[i][0] || '').trim() === nome) abaPenalidades.deleteRow(i + 2);
      }
    }

    const abaParts = ss.getSheetByName(NOMES_ABAS.PARTICIPANTES);
    if (abaParts && abaParts.getLastRow() > 1) {
      const valoresParts = abaParts.getRange(2, 1, abaParts.getLastRow() - 1, 7).getValues();
      for (let i = valoresParts.length - 1; i >= 0; i--) {
        if (String(valoresParts[i][5] || '').trim() === nome) abaParts.deleteRow(i + 2);
      }
    }

    const abaEquipes = ss.getSheetByName(NOMES_ABAS.EQUIPES);
    const valoresEq = abaEquipes.getRange(2, 1, abaEquipes.getLastRow() - 1, 1).getValues();
    for (let i = valoresEq.length - 1; i >= 0; i--) {
      if (String(valoresEq[i][0] || '').trim() === nome) abaEquipes.deleteRow(i + 2);
    }

    // Invalida cache do painel público
    CacheService.getScriptCache().remove(CHAVE_CACHE);
    gravarConfig('cache_version', new Date().toISOString());

    return { sucesso: true, mensagem: 'Equipe excluída' };
  } catch (e) {
    return { sucesso: false, erro: 'Erro inesperado: ' + (e && e.message ? e.message : String(e)) };
  }
}

// ============================================================
// SIDEBAR — GESTÃO DE RESULTADO E PENALIDADES
// ============================================================

/**
 * Carrega tudo que a tela de "pontuação" precisa de uma só vez:
 * - dados da equipe + participantes
 * - resultado atual (se existir)
 * - penalidades da equipe
 * - listas de referência (itens de penalidade 1-10, motivos DSQ 1-7)
 */
function carregarResultadoDaEquipe(nomeEquipe) {
  try {
    const ss = SpreadsheetApp.getActive();
    const paisesMap = lerPaises(ss);
    const equipes = lerEquipes(ss, paisesMap);
    const equipe = equipes.find(function (e) { return e.nome === nomeEquipe; });
    if (!equipe) return { sucesso: false, erro: 'Equipe não encontrada.' };

    const participantes = lerParticipantes(ss);
    vincularParticipantes([equipe], participantes);

    const resultados = lerResultados(ss);
    let resultado = resultados.find(function (r) { return r.equipe === nomeEquipe; });
    if (!resultado) {
      resultado = {
        equipe: nomeEquipe,
        status: 'Aguardando',
        pista: '',
        ordem: '',
        horaInicio: '',
        tempoBrutoStr: '',
        qtdPenalidades: 0,
        tempoTotalStr: '',
        motivoDsq: '',
        observacoes: '',
      };
    }

    const refPen = lerRefSimples(ss, NOMES_ABAS.REF_PENALIDADES);
    const todasPenalidades = lerPenalidades(ss, refPen);
    const penalidades = todasPenalidades.filter(function (p) { return p.equipe === nomeEquipe; });

    const refDsqMap = lerRefSimples(ss, NOMES_ABAS.REF_DSQ);

    return {
      sucesso: true,
      equipe: equipe,
      resultado: resultado,
      penalidades: penalidades,
      refPenalidades: ordenarRefPorItem(refPen),
      refDsq: ordenarRefPorItem(refDsqMap),
    };
  } catch (e) {
    return { sucesso: false, erro: 'Erro: ' + (e && e.message ? e.message : String(e)) };
  }
}

function ordenarRefPorItem(mapa) {
  return Object.keys(mapa)
    .sort(function (a, b) { return Number(a) - Number(b); })
    .map(function (k) { return { item: k, descricao: mapa[k] }; });
}

/**
 * Salva resultado + penalidades de uma equipe em transação atômica.
 * payload = {
 *   nomeEquipe: 'RIT-AC',
 *   resultado: { status, pista, ordem, horaInicio (HH:MM), tempoBrutoStr (mm:ss), motivoDsq, observacoes },
 *   penalidades: [{ item, fase, hora, fiscal, observacoes }, ...]
 * }
 */
function salvarResultadoCompleto(payload) {
  try {
    const ss = SpreadsheetApp.getActive();
    const nomeEquipe = String(payload.nomeEquipe || '').trim();
    if (!nomeEquipe) return { sucesso: false, erro: 'Nome da equipe inválido.' };

    const r = payload.resultado || {};
    const penalidades = payload.penalidades || [];

    // Validações
    const statusValidos = ['Aguardando', 'Em Prova', 'Concluída', 'Desclassificada'];
    if (statusValidos.indexOf(r.status) < 0) return { sucesso: false, erro: 'Status inválido.' };

    if (r.status === 'Desclassificada' && !String(r.motivoDsq || '').trim()) {
      return { sucesso: false, erro: 'Selecione o motivo da desclassificação.' };
    }

    if (r.tempoBrutoStr && !/^\d{1,2}:\d{2}(?:\.\d{1,2})?$/.test(String(r.tempoBrutoStr).trim())) {
      return { sucesso: false, erro: 'Tempo bruto deve estar no formato mm:ss.cc (ex: 12:34.56) ou mm:ss.' };
    }

    if (r.status === 'Concluída' && !String(r.tempoBrutoStr || '').trim()) {
      return { sucesso: false, erro: 'Para status "Concluída", informe o tempo bruto.' };
    }

    // Validação de penalidades
    for (let i = 0; i < penalidades.length; i++) {
      if (!String(penalidades[i].item || '').trim()) {
        return { sucesso: false, erro: 'Penalidade #' + (i + 1) + ': escolha o item do edital.' };
      }
      if (!String(penalidades[i].fase || '').trim()) {
        return { sucesso: false, erro: 'Penalidade #' + (i + 1) + ': escolha a fase.' };
      }
    }

    // Salva linha do resultado (sem mexer em penalidades)
    const resSalvar = _salvarLinhaResultado(ss, nomeEquipe, r);
    if (!resSalvar.sucesso) return resSalvar;

    SpreadsheetApp.flush();

    // Sincronizar Penalidades: deletar todas da equipe e re-inserir
    const abaP = ss.getSheetByName(NOMES_ABAS.PENALIDADES);
    if (!abaP) return { sucesso: false, erro: 'Aba "Penalidades" não encontrada.' };

    const ultimaP = encontrarUltimaLinhaComConteudo(abaP, 1);
    if (ultimaP >= 2) {
      const valoresP = abaP.getRange(2, 1, ultimaP - 1, 1).getValues();
      for (let i = valoresP.length - 1; i >= 0; i--) {
        if (String(valoresP[i][0] || '').trim() === nomeEquipe) {
          abaP.deleteRow(i + 2);
        }
      }
    }

    SpreadsheetApp.flush();

    if (penalidades.length > 0) {
      // Colunas Penalidades (6): equipe(A) item(B) descricao(C) fase(D) fiscal(E) observacoes(F)
      const matriz = penalidades.map(function (p) {
        return [
          nomeEquipe,
          String(p.item || ''),
          '', // descricao — preenchemos via fórmula logo abaixo
          String(p.fase || ''),
          String(p.fiscal || ''),
          String(p.observacoes || ''),
        ];
      });
      const proxLinha = encontrarProximaLinhaVazia(abaP, 1);
      abaP.getRange(proxLinha, 1, matriz.length, 6).setValues(matriz);

      // Fórmula VLOOKUP na coluna C (descricao) — inglês com separador ";"
      const formulas = [];
      for (let i = 0; i < matriz.length; i++) {
        const linha = proxLinha + i;
        formulas.push(['=IF(B' + linha + '="";"";IFERROR(VLOOKUP(B' + linha + ';RefPenalidades!A:B;2;FALSE);""))']);
      }
      abaP.getRange(proxLinha, 3, formulas.length, 1).setFormulas(formulas);
    }

    SpreadsheetApp.flush();

    // Invalida cache do painel público
    CacheService.getScriptCache().remove(CHAVE_CACHE);
    gravarConfig('cache_version', new Date().toISOString());

    return { sucesso: true, mensagem: 'Resultado salvo' };
  } catch (e) {
    return { sucesso: false, erro: 'Erro: ' + (e && e.message ? e.message : String(e)) };
  }
}

/**
 * Atualiza em lote a `ordem_prova` de várias equipes. Usada pelo modal
 * "Definir ordem de execução" da sidebar.
 *
 * @param {Array} ordens — [{ equipe: 'RIT-AC', ordem: 1 }, ...]
 *                         ordem pode ser número, string vazia ou null (sem ordem)
 */
function salvarOrdemExecucao(ordens) {
  try {
    if (!ordens || ordens.length === 0) {
      return { sucesso: false, erro: 'Nenhuma ordem fornecida.' };
    }
    const ss = SpreadsheetApp.getActive();
    const abaEquipes = ss.getSheetByName(NOMES_ABAS.EQUIPES);
    if (!abaEquipes) return { sucesso: false, erro: 'Aba "Equipes" não encontrada.' };

    const ultima = encontrarUltimaLinhaComConteudo(abaEquipes, 1);
    if (ultima < 2) return { sucesso: false, erro: 'Nenhuma equipe cadastrada.' };

    const valoresEq = abaEquipes.getRange(2, 1, ultima - 1, 1).getValues();

    // Indexa nome → linha
    const linhaPorNome = {};
    for (let i = 0; i < valoresEq.length; i++) {
      const nome = String(valoresEq[i][0] || '').trim();
      if (nome) linhaPorNome[nome] = i + 2;
    }

    // Valida unicidade dos números (ignora vazios)
    const numerosUsados = {};
    for (let i = 0; i < ordens.length; i++) {
      const ord = ordens[i].ordem;
      if (ord === '' || ord === null || ord === undefined) continue;
      const n = Number(ord);
      if (isNaN(n) || n < 1) {
        return { sucesso: false, erro: 'Ordem deve ser um número positivo (equipe ' + ordens[i].equipe + ').' };
      }
      if (numerosUsados[n]) {
        return { sucesso: false, erro: 'Número de ordem ' + n + ' está duplicado.' };
      }
      numerosUsados[n] = true;
    }

    // Aplica
    ordens.forEach(function (o) {
      const linha = linhaPorNome[String(o.equipe || '').trim()];
      if (!linha) return;
      const valor = (o.ordem === '' || o.ordem === null || o.ordem === undefined) ? '' : Number(o.ordem);
      abaEquipes.getRange(linha, 8).setValue(valor);
    });

    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(CHAVE_CACHE);
    gravarConfig('cache_version', new Date().toISOString());

    return { sucesso: true, mensagem: 'Ordem de execução salva' };
  } catch (e) {
    return { sucesso: false, erro: 'Erro: ' + (e && e.message ? e.message : String(e)) };
  }
}

/**
 * Inicia até 2 equipes simultaneamente em pistas diferentes.
 * selecoes = [{ equipe: 'RIT-AC', pista: 'A' }, { equipe: 'RIT-SP', pista: 'B' }]
 * horaInicio = string "HH:MM" (opcional — vazio usa hora atual)
 */
function iniciarProvaEmLote(selecoes, horaInicio) {
  try {
    if (!selecoes || selecoes.length === 0) {
      return { sucesso: false, erro: 'Nenhuma equipe selecionada.' };
    }
    if (selecoes.length > 2) {
      return { sucesso: false, erro: 'Máximo de 2 equipes simultâneas (Art. 2º §3º).' };
    }

    // Validar pistas: todas preenchidas e únicas
    const pistas = selecoes.map(function (s) { return String(s.pista || '').trim(); });
    if (pistas.indexOf('') >= 0) {
      return { sucesso: false, erro: 'Selecione a pista de cada equipe.' };
    }
    if (new Set(pistas).size !== pistas.length) {
      return { sucesso: false, erro: 'Equipes precisam estar em pistas diferentes.' };
    }

    const ss = SpreadsheetApp.getActive();

    // Verifica quantas equipes já estão "Em Prova" (excluindo as selecionadas)
    const nomesSelecionados = {};
    selecoes.forEach(function (s) { nomesSelecionados[s.equipe] = true; });
    const resultadosAtuais = lerResultados(ss);
    const emProvaAtual = resultadosAtuais.filter(function (r) {
      return r.status === 'Em Prova' && !nomesSelecionados[r.equipe];
    });
    if (emProvaAtual.length + selecoes.length > 2) {
      return {
        sucesso: false,
        erro: 'Já há ' + emProvaAtual.length + ' equipe(s) em prova. Limite de 2 em pistas simultâneas.',
      };
    }

    // Verifica conflito de pista com equipes que já estão em prova
    for (let i = 0; i < emProvaAtual.length; i++) {
      const pistaUsada = String(emProvaAtual[i].pista || '').trim();
      if (pistas.indexOf(pistaUsada) >= 0) {
        return {
          sucesso: false,
          erro: 'Pista ' + pistaUsada + ' já está em uso pela equipe ' + emProvaAtual[i].equipe + '.',
        };
      }
    }

    // Hora de início: usa fornecida ou "agora"
    const horaUsar = String(horaInicio || '').trim() || _agoraHHMM();

    // Salva linha de resultado para cada equipe (sem mexer em penalidades existentes)
    for (let i = 0; i < selecoes.length; i++) {
      const s = selecoes[i];
      const res = _salvarLinhaResultado(ss, s.equipe, {
        status: 'Em Prova',
        pista: s.pista,
        ordem: '',
        horaInicio: horaUsar,
        tempoBrutoStr: '',
        motivoDsq: '',
        observacoes: '',
      });
      if (!res.sucesso) return res;
    }

    SpreadsheetApp.flush();

    // Invalida cache do painel público
    CacheService.getScriptCache().remove(CHAVE_CACHE);
    gravarConfig('cache_version', new Date().toISOString());

    return { sucesso: true, mensagem: selecoes.length + ' equipe(s) iniciada(s) às ' + horaUsar };
  } catch (e) {
    return { sucesso: false, erro: 'Erro: ' + (e && e.message ? e.message : String(e)) };
  }
}

function _agoraHHMM() {
  const d = new Date();
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  return hh + ':' + mm;
}

/**
 * Salva apenas a linha de resultado (sem mexer em penalidades).
 * Usado por `salvarResultadoCompleto` e `iniciarProvaEmLote`.
 */
function _salvarLinhaResultado(ss, nomeEquipe, r) {
  const horaInicioVal = parseHoraHHMM(r.horaInicio);
  const abaR = ss.getSheetByName(NOMES_ABAS.RESULTADOS);
  if (!abaR) return { sucesso: false, erro: 'Aba "Resultados" não encontrada.' };

  let linhaR = -1;
  const ultimaR = encontrarUltimaLinhaComConteudo(abaR, 1);
  if (ultimaR >= 2) {
    const valoresR = abaR.getRange(2, 1, ultimaR - 1, 1).getValues();
    for (let i = 0; i < valoresR.length; i++) {
      if (String(valoresR[i][0] || '').trim() === nomeEquipe) {
        linhaR = i + 2;
        break;
      }
    }
  }
  if (linhaR < 0) linhaR = encontrarProximaLinhaVazia(abaR, 1);

  abaR.getRange(linhaR, 1).setValue(nomeEquipe);
  abaR.getRange(linhaR, 2).setValue(r.pista || '');
  abaR.getRange(linhaR, 3).setValue(r.ordem !== '' && r.ordem !== null && r.ordem !== undefined ? Number(r.ordem) : '');
  abaR.getRange(linhaR, 4).setValue(horaInicioVal);
  abaR.getRange(linhaR, 5).setValue(r.status || 'Aguardando');
  abaR.getRange(linhaR, 6).setValue(String(r.tempoBrutoStr || '').trim());

  // Fórmulas em inglês com separador ";" (compatível com locale pt-BR)
  abaR.getRange(linhaR, 7).setFormula(
    '=IF(A' + linhaR + '="";"";COUNTIF(Penalidades!A:A;A' + linhaR + '))'
  );
  abaR.getRange(linhaR, 8).setFormula(
    '=IF(G' + linhaR + '="";0;G' + linhaR + '*20)'
  );
  // tempo_total_seg = minutos*60 + segundos + centésimos/100 + segundos_penalidade
  // Aceita "mm:ss" ou "mm:ss.cc" (ou "mm:ss,cc" como variação de locale)
  abaR.getRange(linhaR, 9).setFormula(
    '=IF(F' + linhaR + '="";0;' +
      'IFERROR(VALUE(REGEXEXTRACT(F' + linhaR + ';"^(\\d+):"));0)*60+' +
      'IFERROR(VALUE(REGEXEXTRACT(F' + linhaR + ';":(\\d+)"));0)+' +
      'IFERROR(VALUE(REGEXEXTRACT(F' + linhaR + ';"[.,](\\d+)$"))/100;0)+' +
      'H' + linhaR + ')'
  );

  abaR.getRange(linhaR, 10).setValue(r.status === 'Desclassificada' ? String(r.motivoDsq || '') : '');
  abaR.getRange(linhaR, 11).setValue(String(r.observacoes || ''));

  return { sucesso: true };
}

/**
 * Define a pista (A, B ou '') de uma equipe com status Aguardando,
 * sem alterar nenhum outro campo.
 */
function salvarPistaAguardando(nomeEquipe, pista) {
  try {
    const ss = SpreadsheetApp.getActive();
    const abaR = ss.getSheetByName(NOMES_ABAS.RESULTADOS);
    if (!abaR) return { sucesso: false, erro: 'Aba "Resultados" não encontrada.' };

    const pistaVal = String(pista || '').trim();
    const ultimaR = encontrarUltimaLinhaComConteudo(abaR, 1);

    let linhaR = -1;
    if (ultimaR >= 2) {
      const vals = abaR.getRange(2, 1, ultimaR - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (String(vals[i][0] || '').trim() === nomeEquipe) { linhaR = i + 2; break; }
      }
    }

    if (linhaR >= 0) {
      abaR.getRange(linhaR, 2).setValue(pistaVal);
    } else {
      linhaR = encontrarProximaLinhaVazia(abaR, 1);
      abaR.getRange(linhaR, 1).setValue(nomeEquipe);
      abaR.getRange(linhaR, 2).setValue(pistaVal);
      abaR.getRange(linhaR, 5).setValue('Aguardando');
    }

    CacheService.getScriptCache().remove(CHAVE_CACHE);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// Converte "HH:MM" ou "HH:MM:SS" em Date (para coluna formatada como hora).
// Retorna '' se entrada vazia/inválida.
function parseHoraHHMM(valor) {
  if (!valor) return '';
  const s = String(valor).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return '';
  const d = new Date();
  d.setHours(Number(m[1]), Number(m[2]), Number(m[3] || 0), 0);
  return d;
}

/**
 * Encontra a primeira linha vazia em uma coluna específica.
 * Crítico para abas onde getLastRow() é "envenenado" por checkboxes/fórmulas
 * pré-inseridos em outras colunas. Sempre olhe a coluna A (nome/equipe) que
 * é o identificador real.
 *
 * @param {Sheet} aba
 * @param {number} colIdx 1-based (1 = coluna A)
 * @returns {number} número da próxima linha vazia (mínimo 2)
 */
function encontrarProximaLinhaVazia(aba, colIdx) {
  const ultimaComConteudo = encontrarUltimaLinhaComConteudo(aba, colIdx);
  return ultimaComConteudo + 1;
}

/**
 * Última linha que tem CONTEÚDO REAL na coluna especificada (ignorando
 * células vazias mesmo que outras colunas tenham checkbox/fórmula).
 * Retorna 1 (linha do cabeçalho) se a aba só tem header.
 */
function encontrarUltimaLinhaComConteudo(aba, colIdx) {
  const totalLinhas = aba.getMaxRows();
  if (totalLinhas < 2) return 1;
  const valores = aba.getRange(2, colIdx, totalLinhas - 1, 1).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    const v = valores[i][0];
    if (v !== '' && v !== null && v !== undefined) {
      return i + 2;
    }
  }
  return 1; // só o cabeçalho
}

// ============================================================
// WEB APP — serve a página
// ============================================================

function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.api === 'dados') {
    return ContentService
      .createTextOutput(JSON.stringify(getDadosPainel()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.api === 'versao') {
    return ContentService
      .createTextOutput(JSON.stringify(getVersaoPainel()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Bombeira de Garra — ENBOM 2026')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// API — chamada pelo frontend via google.script.run
// ============================================================

/**
 * Endpoint leve para polling. Lê só a célula cache_version da aba Config.
 * O frontend chama isso a cada 60s; se a versão mudou (gestor clicou em
 * "Atualizar painel público"), aí sim faz reload completo via getDadosPainel.
 */
function getVersaoPainel() {
  const aba = SpreadsheetApp.getActive().getSheetByName(NOMES_ABAS.CONFIG);
  if (!aba) return { cacheVersion: '' };
  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === 'cache_version') {
      return { cacheVersion: String(valores[i][1] || '') };
    }
  }
  return { cacheVersion: '' };
}

function getDadosPainel() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CHAVE_CACHE);
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.getActive();
  const config = lerConfig(ss);
  const paisesMap = lerPaises(ss);
  const refPenalidades = lerRefSimples(ss, NOMES_ABAS.REF_PENALIDADES);
  const refDSQ = lerRefSimples(ss, NOMES_ABAS.REF_DSQ);
  const equipes = lerEquipes(ss, paisesMap);
  const participantes = lerParticipantes(ss);
  vincularParticipantes(equipes, participantes);

  const resultadosRaw = lerResultados(ss);
  const penalidadesRaw = lerPenalidades(ss, refPenalidades);

  const penalidadesPorEquipe = {};
  penalidadesRaw.forEach(function (p) {
    if (!penalidadesPorEquipe[p.equipe]) penalidadesPorEquipe[p.equipe] = [];
    penalidadesPorEquipe[p.equipe].push(p);
  });

  const resultadosEnriquecidos = resultadosRaw.map(function (r) {
    const eq = equipes.find(function (e) { return e.nome === r.equipe; }) || {};
    return Object.assign({}, r, {
      pais: eq.pais || '',
      bandeira: eq.bandeira || '',
      corporacao: eq.corporacao || '',
      unidadeSubnacional: eq.unidadeSubnacional || '',
      capita: eq.capita || '',
      motivoDsqTexto: refDSQ[r.motivoDsq] || '',
    });
  });

  const concluidas = resultadosEnriquecidos
    .filter(function (r) { return r.status === 'Concluída'; })
    .sort(ordenarPorDesempate);
  const emProva = resultadosEnriquecidos.filter(function (r) { return r.status === 'Em Prova'; });
  const desclassificadas = resultadosEnriquecidos.filter(function (r) { return r.status === 'Desclassificada'; });

  const equipesComResultado = new Set(resultadosEnriquecidos.map(function (r) { return r.equipe; }));
  const aguardando = equipes
    .filter(function (e) {
      const r = resultadosEnriquecidos.find(function (x) { return x.equipe === e.nome; });
      return !r || r.status === 'Aguardando';
    })
    .map(function (e) {
      const r = resultadosEnriquecidos.find(function (x) { return x.equipe === e.nome; });
      return {
        equipe: e.nome,
        pais: e.pais,
        bandeira: e.bandeira,
        corporacao: e.corporacao,
        capita: e.capita,
        ordemProva: e.ordemProva,
        pista: (r && r.pista) ? r.pista : '',
      };
    })
    // Ordena por ordem_prova (do sorteio). Equipes sem ordem vão pro fim.
    .sort(function (a, b) {
      if (a.ordemProva == null && b.ordemProva == null) return 0;
      if (a.ordemProva == null) return 1;
      if (b.ordemProva == null) return -1;
      return a.ordemProva - b.ordemProva;
    });

  const dados = {
    titulo: config.titulo_competicao || 'Bombeira de Garra — ENBOM 2026',
    local: config.local || '',
    data: config.data || '',
    ultimaAtualizacao: new Date().toISOString(),
    cacheVersion: config.cache_version || '',
    equipes: equipes,
    resultados: {
      podio: concluidas.slice(0, 3),
      emProva: emProva,
      aguardando: aguardando,
      concluidas: concluidas,
      desclassificadas: desclassificadas,
    },
    penalidadesPorEquipe: penalidadesPorEquipe,
  };

  cache.put(CHAVE_CACHE, JSON.stringify(dados), TTL_CACHE_SEG);
  return dados;
}

// ============================================================
// LEITURA DAS ABAS
// ============================================================

function lerConfig(ss) {
  const aba = ss.getSheetByName(NOMES_ABAS.CONFIG);
  if (!aba) return {};
  const valores = aba.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < valores.length; i++) {
    const chave = String(valores[i][0] || '').trim();
    if (chave) cfg[chave] = valores[i][1];
  }
  return cfg;
}

function gravarConfig(chave, valor) {
  const aba = SpreadsheetApp.getActive().getSheetByName(NOMES_ABAS.CONFIG);
  if (!aba) return;
  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    if (String(valores[i][0]).trim() === chave) {
      aba.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
  aba.appendRow([chave, valor]);
}

function lerPaises(ss) {
  const aba = ss.getSheetByName(NOMES_ABAS.REF_PAISES);
  const mapa = {};
  if (!aba) return mapa;
  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    const pais = String(valores[i][0] || '').trim();
    if (pais) mapa[pais] = String(valores[i][1] || '').trim();
  }
  return mapa;
}

function lerRefSimples(ss, nomeAba) {
  const aba = ss.getSheetByName(nomeAba);
  const mapa = {};
  if (!aba) return mapa;
  const valores = aba.getDataRange().getValues();
  for (let i = 1; i < valores.length; i++) {
    const item = valores[i][0];
    if (item !== '' && item !== null) mapa[String(item)] = String(valores[i][1] || '').trim();
  }
  return mapa;
}

function lerEquipes(ss, paisesMap) {
  const aba = ss.getSheetByName(NOMES_ABAS.EQUIPES);
  if (!aba) return [];
  const valores = aba.getDataRange().getValues();
  const equipes = [];
  for (let i = 1; i < valores.length; i++) {
    const nome = String(valores[i][0] || '').trim();
    if (!nome) continue;
    const pais = String(valores[i][1] || '').trim();
    const bandeiraExplicita = String(valores[i][4] || '').trim();
    const ordemRaw = valores[i][7];
    const ordemProva = (ordemRaw !== '' && ordemRaw !== null && ordemRaw !== undefined && !isNaN(Number(ordemRaw)))
      ? Number(ordemRaw) : null;
    equipes.push({
      nome: nome,
      pais: pais,
      bandeira: bandeiraExplicita || paisesMap[pais] || '🏳️',
      corporacao: String(valores[i][2] || '').trim(),
      unidadeSubnacional: String(valores[i][3] || '').trim(),
      capita: String(valores[i][5] || '').trim(),
      observacoes: String(valores[i][6] || '').trim(),
      ordemProva: ordemProva,
      participantes: [],
    });
  }
  return equipes;
}

function lerParticipantes(ss) {
  const aba = ss.getSheetByName(NOMES_ABAS.PARTICIPANTES);
  if (!aba) return [];
  const valores = aba.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < valores.length; i++) {
    const nome = String(valores[i][0] || '').trim();
    if (!nome) continue;
    const ehCapita = valores[i][6] === true;
    lista.push({
      nome: nome,
      nomeDeGuerra: String(valores[i][1] || '').trim(),
      postoGraduacao: String(valores[i][2] || '').trim(),
      pais: String(valores[i][3] || '').trim(),
      corporacao: String(valores[i][4] || '').trim(),
      equipe: String(valores[i][5] || '').trim(),
      ehCapita: ehCapita,
      funcao: ehCapita ? 'Capitã' : 'Combatente',
    });
  }
  return lista;
}

function vincularParticipantes(equipes, participantes) {
  const porEquipe = {};
  participantes.forEach(function (p) {
    if (!porEquipe[p.equipe]) porEquipe[p.equipe] = [];
    porEquipe[p.equipe].push(p);
  });
  equipes.forEach(function (e) {
    e.participantes = porEquipe[e.nome] || [];
  });
}

function lerResultados(ss) {
  const aba = ss.getSheetByName(NOMES_ABAS.RESULTADOS);
  if (!aba) return [];
  const valores = aba.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < valores.length; i++) {
    const equipe = String(valores[i][0] || '').trim();
    if (!equipe) continue;
    const tempoBrutoStr = String(valores[i][5] || '').trim();
    const qtdPenalidades = Number(valores[i][6]) || 0;
    const tempoBrutoSeg = mmssParaSeg(tempoBrutoStr);
    const tempoTotalSeg = tempoBrutoSeg !== null ? tempoBrutoSeg + qtdPenalidades * 20 : null;
    lista.push({
      equipe: equipe,
      pista: valores[i][1] !== '' ? String(valores[i][1]) : '',
      ordem: Number(valores[i][2]) || null,
      horaInicio: formatarHora(valores[i][3]),
      status: String(valores[i][4] || 'Aguardando').trim(),
      tempoBrutoStr: tempoBrutoStr,
      tempoBrutoSeg: tempoBrutoSeg,
      qtdPenalidades: qtdPenalidades,
      segundosPenalidade: qtdPenalidades * 20,
      tempoTotalSeg: tempoTotalSeg,
      tempoTotalStr: tempoTotalSeg !== null ? segParaMmss(tempoTotalSeg) : '',
      motivoDsq: valores[i][9] !== '' && valores[i][9] !== null ? String(valores[i][9]).trim() : '',
      observacoes: String(valores[i][10] || '').trim(),
    });
  }
  return lista;
}

function lerPenalidades(ss, refPenalidades) {
  const aba = ss.getSheetByName(NOMES_ABAS.PENALIDADES);
  if (!aba) return [];
  const valores = aba.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < valores.length; i++) {
    const equipe = String(valores[i][0] || '').trim();
    if (!equipe) continue;
    const itemStr = String(valores[i][1] !== '' ? valores[i][1] : '').trim();
    // Estrutura nova (6 colunas): equipe | item | descricao | fase | fiscal | observacoes
    lista.push({
      equipe: equipe,
      item: itemStr,
      descricao: String(valores[i][2] || '').trim() || refPenalidades[itemStr] || '',
      fase: String(valores[i][3] || '').trim(),
      fiscal: String(valores[i][4] || '').trim(),
      observacoes: String(valores[i][5] || '').trim(),
    });
  }
  return lista;
}

// ============================================================
// UTILITÁRIOS
// ============================================================

// Converte "mm:ss.cc" (ou "mm:ss") em segundos com casas decimais.
// Aceita 1-2 dígitos nos centésimos (ex.: "10:22.5" = 622.5).
function mmssParaSeg(str) {
  if (!str) return null;
  const m = String(str).match(/^(\d{1,2}):(\d{2})(?:\.(\d{1,2}))?$/);
  if (!m) return null;
  // Padroniza centésimos para 2 dígitos: "5" → "50" (50 centésimos)
  const cent = m[3] ? Number((m[3] + '0').substring(0, 2)) / 100 : 0;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10) + cent;
}

// Converte segundos (com decimais) em string "mm:ss.cc" (sempre com centésimos).
function segParaMmss(seg) {
  if (seg === null || seg === undefined) return '';
  const totalCent = Math.round(seg * 100);
  const mm = Math.floor(totalCent / 6000);
  const ss = Math.floor((totalCent % 6000) / 100);
  const cc = totalCent % 100;
  const pad = function (n) { return (n < 10 ? '0' : '') + n; };
  return pad(mm) + ':' + pad(ss) + '.' + pad(cc);
}

function formatarHora(valor) {
  if (!valor) return '';
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone() || 'America/Rio_Branco', 'HH:mm:ss');
  }
  return String(valor).trim();
}

function ordenarPorDesempate(a, b) {
  // Regra do edital Art. 16:
  // 1º critério: menor tempo total
  // Desempate: menor nº de penalidades → menor tempo bruto
  if (a.tempoTotalSeg !== b.tempoTotalSeg) return a.tempoTotalSeg - b.tempoTotalSeg;
  if (a.qtdPenalidades !== b.qtdPenalidades) return a.qtdPenalidades - b.qtdPenalidades;
  return (a.tempoBrutoSeg || 0) - (b.tempoBrutoSeg || 0);
}
