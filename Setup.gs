/**
 * Setup automatizado da planilha — Painel Bombeira de Garra ENBOM 2026
 *
 * Como usar:
 *   1. Abra a planilha em branco no Google Sheets.
 *   2. Extensões → Apps Script.
 *   3. Cole este arquivo (Setup.gs), Codigo.gs e Index.html no projeto.
 *   4. No editor do Apps Script, selecione a função "inicializarPlanilha"
 *      no menu suspenso e clique em ▶ Executar.
 *   5. Autorize as permissões solicitadas.
 *   6. Aguarde — em ~10s a planilha estará totalmente configurada.
 *   7. Depois disso, este arquivo (Setup.gs) pode ser deletado se quiser.
 *
 * O script é idempotente: pode rodar várias vezes, ele recria as abas do zero
 * (CUIDADO: apaga dados existentes nas abas que ele controla).
 */

function inicializarPlanilha() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    '⚠️ Inicializar planilha',
    'Este script vai (re)criar as 8 abas do painel. Se já houver dados nessas abas, eles serão APAGADOS.\n\n' +
    'Deseja continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resposta !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  const log = [];

  log.push(criarAbaEquipes(ss));
  log.push(criarAbaParticipantes(ss));
  log.push(criarAbaResultados(ss));
  log.push(criarAbaPenalidades(ss));
  log.push(criarAbaRefPenalidades(ss));
  log.push(criarAbaRefDSQ(ss));
  log.push(criarAbaRefPaises(ss));
  log.push(criarAbaConfig(ss));

  // Aplicar validações que dependem de outras abas (precisa fazer depois)
  log.push(aplicarValidacoesEquipes(ss));
  log.push(aplicarValidacoesParticipantes(ss));
  log.push(aplicarValidacoesResultados(ss));
  log.push(aplicarValidacoesPenalidades(ss));

  // Remover aba "Página1"/"Sheet1" padrão se ainda existir e estiver vazia
  removerAbaPadrao(ss);

  // Reordenar abas
  reordenarAbas(ss);

  // Configurar abertura automática da sidebar (installable trigger)
  log.push(configurarAberturaAutomatica());

  ui.alert(
    '✅ Planilha inicializada',
    'Pronto! Configurei:\n\n' + log.join('\n') +
    '\n\n📋 Próximos passos:\n' +
    '1. Cadastre as equipes na aba "Equipes"\n' +
    '2. Cadastre as participantes na aba "Participantes"\n' +
    '3. Recarregue a planilha (F5) para o menu "🏆 Painel ENBOM" aparecer\n' +
    '4. Implante o Web App (Apps Script → Implantar → Nova implantação)',
    ui.ButtonSet.OK
  );
}

// ============================================================
// CRIAÇÃO DAS ABAS
// ============================================================

function criarAbaEquipes(ss) {
  const aba = recriarAba(ss, 'Equipes');
  const headers = ['nome_equipe', 'pais', 'corporacao', 'unidade_subnacional', 'bandeira_emoji', 'capita', 'observacoes', 'ordem_prova'];
  aba.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fde68a');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 130); // nome_equipe
  aba.setColumnWidth(2, 100); // pais
  aba.setColumnWidth(3, 180); // corporacao
  aba.setColumnWidth(4, 140); // unidade_subnacional
  aba.setColumnWidth(5, 100); // bandeira_emoji
  aba.setColumnWidth(6, 220); // capita
  aba.setColumnWidth(7, 220); // observacoes
  aba.setColumnWidth(8, 100); // ordem_prova
  return '✓ Aba "Equipes" criada';
}

function criarAbaParticipantes(ss) {
  const aba = recriarAba(ss, 'Participantes');
  const headers = ['nome', 'nome_de_guerra', 'posto_graduacao', 'pais', 'corporacao', 'equipe', 'eh_capita'];
  aba.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fde68a');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 240); // nome
  aba.setColumnWidth(2, 150); // nome_de_guerra
  aba.setColumnWidth(3, 130); // posto_graduacao
  aba.setColumnWidth(4, 100); // pais
  aba.setColumnWidth(5, 180); // corporacao
  aba.setColumnWidth(6, 130); // equipe
  aba.setColumnWidth(7, 90);  // eh_capita

  // IMPORTANTE: NÃO pré-inserir checkboxes em G2:G inteiro.
  // Isso preenche 999 células com FALSE e quebra getLastRow().
  // Os checkboxes são aplicados linha-a-linha pelo Codigo.gs quando uma
  // participante é cadastrada via interface.
  return '✓ Aba "Participantes" criada';
}

function criarAbaResultados(ss) {
  const aba = recriarAba(ss, 'Resultados');
  const headers = [
    'equipe', 'pista', 'ordem', 'hora_inicio', 'status',
    'tempo_bruto_mmss', 'qtd_penalidades', 'segundos_penalidade',
    'tempo_total_seg', 'motivo_dsq', 'observacoes'
  ];
  aba.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fecaca');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 130); // equipe
  aba.setColumnWidth(2, 70);  // pista
  aba.setColumnWidth(3, 70);  // ordem
  aba.setColumnWidth(4, 110); // hora_inicio
  aba.setColumnWidth(5, 140); // status
  aba.setColumnWidth(6, 130); // tempo_bruto_mmss
  aba.setColumnWidth(7, 130); // qtd_penalidades
  aba.setColumnWidth(8, 140); // segundos_penalidade
  aba.setColumnWidth(9, 130); // tempo_total_seg
  aba.setColumnWidth(10, 110); // motivo_dsq
  aba.setColumnWidth(11, 250); // observacoes

  // Formatos das colunas que serão preenchidas pela interface
  aba.getRange('F2:F').setNumberFormat('@');       // tempo_bruto_mmss como texto
  aba.getRange('D2:D').setNumberFormat('HH:mm:ss'); // hora_inicio

  // Formatação condicional para coluna status
  const regraEmProva = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Em Prova')
    .setBackground('#fee2e2').setFontColor('#991b1b').setBold(true)
    .setRanges([aba.getRange('E2:E')]).build();
  const regraConcluida = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Concluída')
    .setBackground('#dcfce7').setFontColor('#166534').setBold(true)
    .setRanges([aba.getRange('E2:E')]).build();
  const regraDSQ = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Desclassificada')
    .setBackground('#e2e8f0').setFontColor('#475569').setBold(true)
    .setRanges([aba.getRange('E2:E')]).build();
  aba.setConditionalFormatRules([regraEmProva, regraConcluida, regraDSQ]);

  return '✓ Aba "Resultados" criada';
}

function criarAbaPenalidades(ss) {
  const aba = recriarAba(ss, 'Penalidades');
  const headers = ['equipe', 'item', 'descricao', 'fase', 'fiscal', 'observacoes'];
  aba.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fecaca');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 130); // equipe
  aba.setColumnWidth(2, 70);  // item
  aba.setColumnWidth(3, 400); // descricao
  aba.setColumnWidth(4, 70);  // fase
  aba.setColumnWidth(5, 180); // fiscal
  aba.setColumnWidth(6, 250); // observacoes

  return '✓ Aba "Penalidades" criada';
}

function criarAbaRefPenalidades(ss) {
  const aba = recriarAba(ss, 'RefPenalidades');
  const dados = [
    ['item', 'descricao'],
    [1, 'Jogar, arremessar ou deixar equipamentos caírem no chão ou desajustados'],
    [2, 'Deixar equipamento fora da área indicada pelo staff'],
    [3, 'Finalizar a fase, tendo deixado para trás equipamentos ou EPI (exceto mangueiras/esguicho)'],
    [4, 'Executar procedimento técnico que danifique ou tenha potencial de danificar equipamento'],
    [5, 'Iniciar a prova ou deslocar para fase seguinte antes da autorização do staff'],
    [6, 'Abrir o esguicho com apenas uma pessoa na linha de ataque'],
    [7, 'Tentar realizar o combate antes da linha de posicionamento demarcada'],
    [8, 'Não fechar o divisor ou o esguicho após o combate ao incêndio'],
    [9, 'Jogar o esguicho ao chão e não recolocar no local indicado'],
    [10, 'Arrastar a vítima pelo DRD ou outro local não permitido'],
  ];
  aba.getRange(1, 1, dados.length, 2).setValues(dados);
  aba.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#bfdbfe');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 60);
  aba.setColumnWidth(2, 500);
  return '✓ Aba "RefPenalidades" populada (10 itens)';
}

function criarAbaRefDSQ(ss) {
  const aba = recriarAba(ss, 'RefDSQ');
  const dados = [
    ['item', 'descricao'],
    [1, 'Qualquer membro da equipe retirar a máscara fora da zona fria'],
    [2, 'Não comparecer (guarnição de 3 pessoas) no local de início quando chamado pelo staff'],
    [3, 'Desacato por qualquer membro ao staff'],
    [4, 'Repetir, insistindo em procedimento inseguro, após apontamento de correção'],
    [5, 'Causar lesões ou deixar partes do corpo da vítima chocar-se contra solo ou equipamentos'],
    [6, 'Não completar a prova em até 15 minutos'],
    [7, 'Encontrar-se com menos de 2 pessoas na finalização da prova'],
  ];
  aba.getRange(1, 1, dados.length, 2).setValues(dados);
  aba.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#bfdbfe');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 60);
  aba.setColumnWidth(2, 500);
  return '✓ Aba "RefDSQ" populada (7 motivos)';
}

function criarAbaRefPaises(ss) {
  const aba = recriarAba(ss, 'RefPaises');
  const dados = [
    ['pais', 'bandeira_emoji'],
    ['Brasil', '🇧🇷'],
    ['Argentina', '🇦🇷'],
    ['Bolívia', '🇧🇴'],
    ['Chile', '🇨🇱'],
    ['Colômbia', '🇨🇴'],
    ['Equador', '🇪🇨'],
    ['Paraguai', '🇵🇾'],
    ['Peru', '🇵🇪'],
    ['Uruguai', '🇺🇾'],
    ['Venezuela', '🇻🇪'],
  ];
  aba.getRange(1, 1, dados.length, 2).setValues(dados);
  aba.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#bfdbfe');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 140);
  aba.setColumnWidth(2, 140);
  return '✓ Aba "RefPaises" populada (10 países)';
}

function criarAbaConfig(ss) {
  const aba = recriarAba(ss, 'Config');
  const dados = [
    ['chave', 'valor'],
    ['segundos_por_penalidade', 20],
    ['tempo_maximo_min', 15],
    ['titulo_competicao', 'Bombeira de Garra — XI ENBOM 2026'],
    ['local', 'CBMAC — Rio Branco/AC'],
    ['data', '22/05/2026'],
    ['cache_version', ''],
    ['url_painel_publico', ''],
  ];
  aba.getRange(1, 1, dados.length, 2).setValues(dados);
  aba.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#bfdbfe');
  aba.setFrozenRows(1);
  aba.setColumnWidth(1, 200);
  aba.setColumnWidth(2, 320);
  return '✓ Aba "Config" populada';
}

// ============================================================
// VALIDAÇÕES (DROPDOWNS)
// ============================================================

function aplicarValidacoesEquipes(ss) {
  const aba = ss.getSheetByName('Equipes');
  const abaPaises = ss.getSheetByName('RefPaises');
  const ruleVal = SpreadsheetApp.newDataValidation()
    .requireValueInRange(abaPaises.getRange('A2:A1000'), true)
    .setAllowInvalid(false)
    .build();
  aba.getRange('B2:B').setDataValidation(ruleVal);
  return '✓ Validações da aba "Equipes" aplicadas';
}

function aplicarValidacoesParticipantes(ss) {
  const aba = ss.getSheetByName('Participantes');
  const abaPaises = ss.getSheetByName('RefPaises');
  const abaEquipes = ss.getSheetByName('Equipes');

  // pais agora está na coluna D (após nome_de_guerra)
  aba.getRange('D2:D').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(abaPaises.getRange('A2:A1000'), true)
      .setAllowInvalid(false).build()
  );
  // equipe agora está na coluna F
  aba.getRange('F2:F').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(abaEquipes.getRange('A2:A1000'), true)
      .setAllowInvalid(false).build()
  );
  return '✓ Validações da aba "Participantes" aplicadas';
}

function aplicarValidacoesResultados(ss) {
  const aba = ss.getSheetByName('Resultados');
  const abaEquipes = ss.getSheetByName('Equipes');
  const abaDSQ = ss.getSheetByName('RefDSQ');

  // equipe (A)
  aba.getRange('A2:A').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(abaEquipes.getRange('A2:A1000'), true)
      .setAllowInvalid(false).build()
  );
  // pista (B)
  aba.getRange('B2:B').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['1', '2'], true)
      .setAllowInvalid(false).build()
  );
  // status (E)
  aba.getRange('E2:E').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Aguardando', 'Em Prova', 'Concluída', 'Desclassificada'], true)
      .setAllowInvalid(false).build()
  );
  // motivo_dsq (J)
  aba.getRange('J2:J').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(abaDSQ.getRange('A2:A1000'), true)
      .setAllowInvalid(true).build()
  );
  return '✓ Validações da aba "Resultados" aplicadas';
}

function aplicarValidacoesPenalidades(ss) {
  const aba = ss.getSheetByName('Penalidades');
  const abaEquipes = ss.getSheetByName('Equipes');
  const abaPen = ss.getSheetByName('RefPenalidades');

  aba.getRange('A2:A').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(abaEquipes.getRange('A2:A1000'), true)
      .setAllowInvalid(false).build()
  );
  aba.getRange('B2:B').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(abaPen.getRange('A2:A1000'), true)
      .setAllowInvalid(false).build()
  );
  aba.getRange('D2:D').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['I', 'II', 'III', 'IV', 'V'], true)
      .setAllowInvalid(false).build()
  );
  return '✓ Validações da aba "Penalidades" aplicadas';
}

// ============================================================
// TRIGGER INSTALÁVEL — abre sidebar automaticamente
// ============================================================

function configurarAberturaAutomatica() {
  try {
    // Remove triggers antigos pra evitar duplicação
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (t.getHandlerFunction() === 'aoAbrirPlanilha') {
        ScriptApp.deleteTrigger(t);
      }
    });
    // Cria trigger instalável (rodando com auth completa, ao contrário do simple onOpen)
    ScriptApp.newTrigger('aoAbrirPlanilha')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
    return '✓ Abertura automática da sidebar configurada';
  } catch (e) {
    return '⚠ Não foi possível configurar abertura automática: ' + (e && e.message ? e.message : String(e));
  }
}

// ============================================================
// HELPERS
// ============================================================

function recriarAba(ss, nome) {
  const existente = ss.getSheetByName(nome);
  if (existente) {
    existente.clear();
    existente.clearConditionalFormatRules();
    // Remove validações antigas
    const range = existente.getRange(1, 1, existente.getMaxRows(), existente.getMaxColumns());
    range.setDataValidation(null);
    return existente;
  }
  return ss.insertSheet(nome);
}

function removerAbaPadrao(ss) {
  const candidatos = ['Página1', 'Sheet1', 'Folha1', 'Hoja 1'];
  candidatos.forEach(function (nome) {
    const aba = ss.getSheetByName(nome);
    if (aba && aba.getLastRow() === 0) {
      ss.deleteSheet(aba);
    }
  });
}

function reordenarAbas(ss) {
  const ordemDesejada = [
    'Resultados', 'Penalidades', 'Equipes', 'Participantes',
    'Config', 'RefPenalidades', 'RefDSQ', 'RefPaises'
  ];
  ordemDesejada.forEach(function (nome, idx) {
    const aba = ss.getSheetByName(nome);
    if (aba) {
      ss.setActiveSheet(aba);
      ss.moveActiveSheet(idx + 1);
    }
  });
  ss.setActiveSheet(ss.getSheetByName('Resultados'));
}

// ============================================================
// FUNÇÃO BÔNUS: popular dados de exemplo para teste
// ============================================================

function popularDadosExemplo() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.alert(
    '🧪 Popular dados de exemplo',
    'Vou adicionar 4 equipes de exemplo, 8 participantes e 4 resultados (1 aguardando, 1 em prova, 1 concluída com 2 penalidades, 1 desclassificada).\n\n' +
    'Útil apenas para testar o painel. Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resposta !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();

  const equipes = [
    ['RIT-AC', 'Brasil', 'CBMAC', 'Acre', '🇧🇷', '', 'Equipe anfitriã', 1],
    ['RIT-SP', 'Brasil', 'CBPMESP', 'São Paulo', '🇧🇷', '', '', 2],
    ['RIT-RJ', 'Brasil', 'CBMERJ', 'Rio de Janeiro', '🇧🇷', '', '', 3],
    ['RIT-MG', 'Brasil', 'CBMMG', 'Minas Gerais', '🇧🇷', '', '', 4],
  ];
  ss.getSheetByName('Equipes').getRange(2, 1, equipes.length, 8).setValues(equipes);

  const participantes = [
    ['CAP Akauany Ferraz Pereira', 'AKAUANY', 'CAP', 'Brasil', 'CBMAC', 'RIT-AC', true],
    ['SGT Mariana Souza', 'SOUZA', 'SGT', 'Brasil', 'CBMAC', 'RIT-AC', false],
    ['CB Patrícia Oliveira', 'PATRÍCIA', 'CB', 'Brasil', 'CBMAC', 'RIT-AC', false],
    ['TEN Maria Silva', 'SILVA', 'TEN', 'Brasil', 'CBPMESP', 'RIT-SP', true],
    ['SGT Camila Rocha', 'ROCHA', 'SGT', 'Brasil', 'CBPMESP', 'RIT-SP', false],
    ['CAP Ana Costa', 'COSTA', 'CAP', 'Brasil', 'CBMERJ', 'RIT-RJ', true],
    ['SGT Beatriz Almeida', 'ALMEIDA', 'SGT', 'Brasil', 'CBMERJ', 'RIT-RJ', false],
    ['TEN Joana Lima', 'LIMA', 'TEN', 'Brasil', 'CBMMG', 'RIT-MG', true],
  ];
  ss.getSheetByName('Participantes').getRange(2, 1, participantes.length, 7).setValues(participantes);

  const resultados = [
    ['RIT-AC', '1', 1, new Date(2026, 4, 22, 8, 0, 0), 'Concluída', '10:22', 1, '', '', '', 'Excelente atuação'],
    ['RIT-SP', '2', 2, new Date(2026, 4, 22, 8, 0, 0), 'Concluída', '11:08', 0, '', '', '', ''],
    ['RIT-RJ', '1', 3, new Date(2026, 4, 22, 8, 30, 0), 'Em Prova', '', 0, '', '', '', ''],
    ['RIT-MG', '', 4, '', 'Aguardando', '', 0, '', '', '', ''],
  ];
  ss.getSheetByName('Resultados').getRange(2, 1, resultados.length, 11).setValues(resultados);

  const penalidades = [
    ['RIT-AC', 5, '', 'I', 'CAP Marques', 'Avançou antes da liberação'],
  ];
  ss.getSheetByName('Penalidades').getRange(2, 1, penalidades.length, 6).setValues(penalidades);

  ui.alert('✅ Dados de exemplo inseridos! Agora rode "atualizarPainelPublico" no menu da planilha e abra o painel para conferir.');
}
