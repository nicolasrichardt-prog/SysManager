/**
 * ====================================================================
 * SysManager - Sistema de Gestão Ágil de Tarefas e Esteira de Demandas
 * Backend Google Apps Script com Autenticação e Persistência Completa (Code.gs)
 * ====================================================================
 */

// 🔒 ADMINISTRADOR: Sempre autorizado por padrão
const ADMIN_EMAIL = 'nicolasrichardt@gmail.com';
const ADMIN_MESTRE_EMAIL = ADMIN_EMAIL; // Compatibilidade retroativa

// ⚙️ CONFIGURAÇÃO DA PLANILHA (OPCIONAL):
// - Se o script foi criado via "Extensões > Apps Script" dentro de uma Planilha Google, deixe vazio ('').
// - Se o script for do tipo avulso (criado em script.google.com), cole o ID da sua planilha abaixo entre aspas:
const SPREADSHEET_ID = '';

const SYSMANAGER_EMAIL = 'sysmanager.sm@gmail.com';
const TIMEZONE_PADRAO = 'America/Sao_Paulo';

const SHEETS = {
  USUARIOS: 'Usuarios',
  TAREFAS: 'Tarefas',
  SOLICITACOES: 'SolicitacoesPrazo',
  DEMANDAS_FUTURAS: 'DemandasFuturas',
  ANOTACOES: 'Anotacoes',
  ANOTACOES_ABAS: 'AnotacoesAbas'
};

function obterTimeZone_() {
  try {
    const tz = Session.getScriptTimeZone();
    if (tz && (tz.indexOf('Sao_Paulo') > -1 || tz.indexOf('Brasilia') > -1 || tz.indexOf('Fortaleza') > -1 || tz.indexOf('Recife') > -1 || tz.indexOf('GMT-03') > -1)) {
      return tz;
    }
  } catch (e) {}
  return TIMEZONE_PADRAO;
}

/**
 * Converte string YYYY-MM-DD, DD/MM/YYYY ou Date para objeto Date local (meio-dia) sem recuo de fuso UTC
 */
function parseDataSegura_(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return isNaN(val.getTime()) ? null : new Date(val.getFullYear(), val.getMonth(), val.getDate(), 12, 0, 0);
  }
  const str = val.toString().trim();
  if (!str || str === '-' || str.toLowerCase() === 'a definir') return null;

  const mIso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (mIso) {
    return new Date(parseInt(mIso[1], 10), parseInt(mIso[2], 10) - 1, parseInt(mIso[3], 10), 12, 0, 0);
  }

  const mBr = str.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (mBr) {
    return new Date(parseInt(mBr[3], 10), parseInt(mBr[2], 10) - 1, parseInt(mBr[1], 10), 12, 0, 0);
  }

  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate(), 12, 0, 0);
}

function formatarDataIsoSegura_(val) {
  const d = parseDataSegura_(val);
  if (!d) return '';
  return Utilities.formatDate(d, obterTimeZone_(), 'yyyy-MM-dd');
}

function getDatabaseSpreadsheet() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {
      Logger.log('Aviso ao abrir SPREADSHEET_ID: ' + e.message);
    }
  }

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SYSMANAGER_SHEET_ID');
  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (e) {}
  }

  ss = SpreadsheetApp.create('SysManager_Database');
  props.setProperty('SYSMANAGER_SHEET_ID', ss.getId());
  return ss;
}

/**
 * Obtém a lista dinâmica de e-mails autorizados diretamente da aba 'Usuarios' da Planilha
 */
function getEmailsAutorizadosDinamicos() {
  const emails = [ADMIN_EMAIL.toLowerCase().trim()];
  try {
    const ss = getDatabaseSpreadsheet();
    const shUsuarios = ss.getSheetByName(SHEETS.USUARIOS);

    if (shUsuarios && shUsuarios.getLastRow() > 1) {
      const rawUsers = shUsuarios.getRange(2, 1, shUsuarios.getLastRow() - 1, Math.max(5, shUsuarios.getLastColumn())).getValues();

      rawUsers.forEach(r => {
        for (let c = 1; c < r.length; c++) {
          const val = (r[c] || '').toString().trim().toLowerCase();
          if (val.includes('@') && !emails.includes(val)) {
            emails.push(val);
          }
        }
      });
    }
  } catch (err) {
    Logger.log('Aviso ao ler e-mails autorizados: ' + err.message);
  }

  return emails;
}

/**
 * Validação de e-mail para acesso
 */
function verificarEmailAutorizado(email) {
  const emailLimpo = (email || '').toLowerCase().trim();
  if (!emailLimpo) {
    return {
      autorizado: false,
      email: '',
      message: 'Por favor, informe o seu e-mail @gmail.com.'
    };
  }

  // Fast-track para o Administrador: NUNCA bloqueia, resposta em 0ms
  if (emailLimpo === ADMIN_EMAIL.toLowerCase()) {
    return {
      autorizado: true,
      email: ADMIN_EMAIL,
      usuario: {
        id: 'USR-1',
        nome: 'Ten Nicolas',
        email: ADMIN_EMAIL,
        cargo: 'Gerente de Projetos'
      }
    };
  }

  try {
    const autorizados = getEmailsAutorizadosDinamicos();
    const isAuth = autorizados.includes(emailLimpo);

    if (!isAuth) {
      return {
        autorizado: false,
        email: emailLimpo,
        message: 'E-mail "' + emailLimpo + '" não está cadastrado na equipe. Solicite ao Administrador Nicolas Richardt o cadastro na aba Equipe.'
      };
    }

    const ss = getDatabaseSpreadsheet();
    const shUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
    let userEncontrado = null;

    if (shUsuarios && shUsuarios.getLastRow() > 1) {
      const rows = shUsuarios.getRange(2, 1, shUsuarios.getLastRow() - 1, 5).getValues();

      for (let r of rows) {
        if (!r[0]) continue;
        const uEmail = (r[2] || '').toString().trim().toLowerCase();
        if (uEmail === emailLimpo) {
          let cargoVal = (r[3] || 'Desenvolvedor').toString().trim();
          if (cargoVal === 'Ambos' || cargoVal === 'Gerente' || cargoVal.toLowerCase().includes('gerente') || cargoVal.toLowerCase().includes('ambos')) {
            cargoVal = 'Gerente de Projetos';
          } else {
            cargoVal = 'Desenvolvedor';
          }
          userEncontrado = {
            id: (r[0] || '').toString().trim(),
            nome: (r[1] || '').toString().trim(),
            email: uEmail,
            cargo: cargoVal
          };
          break;
        }
      }
    }

    return {
      autorizado: true,
      email: emailLimpo,
      usuario: userEncontrado || { id: 'USR-DEV', nome: emailLimpo.split('@')[0], email: emailLimpo, cargo: 'Desenvolvedor' }
    };
  } catch (err) {
    Logger.log('Erro ao validar email: ' + err.message);
    return {
      autorizado: false,
      email: emailLimpo,
      message: 'Erro no servidor ao validar acesso: ' + err.message
    };
  }
}

/**
 * Endpoint de verificação chamado pela tela de login do cliente
 */
function verificarEmailLogin(email) {
  try {
    return verificarEmailAutorizado(email);
  } catch (e) {
    return {
      autorizado: false,
      email: email || '',
      message: 'Erro na autenticação: ' + e.message
    };
  }
}

/**
 * Validação de autorização para ações no backend
 */
function verificarAutorizacao(emailParam) {
  return {
    autorizado: true,
    email: emailParam || ADMIN_EMAIL
  };
}

/**
 * Ponto de entrada do Web App
 */
function doGet(e) {
  initDatabaseIfNeeded();

  if (e && e.parameter && e.parameter.action) {
    return processarRequisicaoApi(e.parameter.action, e.parameter);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SysManager - Gestão Ágil')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  initDatabaseIfNeeded();
  var action = '';
  var payload = {};

  if (e && e.postData && e.postData.contents) {
    try {
      var parsed = JSON.parse(e.postData.contents);
      action = parsed.action || '';
      payload = parsed.data || {};
    } catch (err) {
      action = (e.parameter && e.parameter.action) || '';
      payload = e.parameter || {};
    }
  } else if (e && e.parameter) {
    action = e.parameter.action || '';
    payload = e.parameter || {};
  }

  return processarRequisicaoApi(action, payload);
}

function processarRequisicaoApi(action, data) {
  var result = { success: false, message: 'Ação não reconhecida: ' + action };

  try {
    if (action === 'getAppData') {
      result = getAppData(data.email || '');
    } else if (action === 'verificarEmailLogin') {
      result = verificarEmailLogin(data.email || '');
    } else if (action === 'cadastrarTarefa') {
      result = cadastrarTarefa(data);
    } else if (action === 'selecionarTarefaAtiva') {
      result = selecionarTarefaAtiva(data.tarefaId, data.devId, data);
    } else if (action === 'concluirTarefa' || action === 'finalizarTarefa') {
      result = concluirTarefa(data.tarefaId, data.devId);
    } else if (action === 'solicitarNovoPrazo') {
      result = solicitarNovoPrazo(data);
    } else if (action === 'darCientePrazo') {
      result = darCientePrazo(data.solicitacaoId, data.gerenteNome || '');
    } else if (action === 'excluirTarefa') {
      result = excluirTarefa(data.tarefaId || data.id);
    } else if (action === 'trocarOrdemTarefas' || action === 'reordenarTarefas') {
      result = trocarOrdemTarefas(data.tarefaId1 || data.id1, data.tarefaId2 || data.id2);
    } else if (action === 'atualizarTarefa' || action === 'editarTarefa' || action === 'salvarEdicaoTarefa') {
      result = atualizarTarefa(data);
    } else if (action === 'cadastrarUsuario') {
      result = cadastrarUsuario(data);
    } else if (action === 'atualizarUsuario') {
      result = atualizarUsuario(data);
    } else if (action === 'excluirUsuario') {
      result = excluirUsuario(data.usuarioId || data.id || data);
    } else if (action === 'cadastrarDemandaFutura') {
      result = cadastrarDemandaFutura(data);
    } else if (action === 'atualizarDemandaFutura') {
      result = atualizarDemandaFutura(data);
    } else if (action === 'excluirDemandaFutura') {
      result = excluirDemandaFutura(data.demandaFuturaId || data.id || data);
    } else if (action === 'processarDemandasFuturas') {
      result = processarDemandasFuturasAgendadas();
    } else if (action === 'dispararEmailDemandaFuturaManual') {
      result = dispararEmailDemandaFuturaManual(data.demandaFuturaId || data.id);
    } else if (action === 'salvarAnotacao') {
      result = salvarAnotacaoNoBanco(data);
    } else if (action === 'excluirAnotacao') {
      result = excluirAnotacaoDoBanco(data);
    } else if (action === 'salvarAbasAnotacoes') {
      result = salvarAbasAnotacoesNoBanco(data);
    } else if (action === 'sincronizarAnotacoesUsuario') {
      result = sincronizarAnotacoesUsuarioNoBanco(data);
    }
  } catch (err) {
    result = { success: false, error: err.message || err.toString(), message: err.message || err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function initDatabaseIfNeeded() {
  const ss = getDatabaseSpreadsheet();
  let shUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
  if (!shUsuarios) {
    initDatabase();
  } else {
    try { ensureDemandasFuturasSheet_(ss); } catch (e) {}
    try { ensureAnotacoesSheets_(ss); } catch (e) {}
  }
  try { garantirGatilhoDemandasFuturas_(); } catch (e) {}
}

function initDatabase() {
  const ss = getDatabaseSpreadsheet();

  // 1. Aba Usuários padronizada com ID, Nome, Email, Cargo, DataCriacao
  let shUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
  if (!shUsuarios) {
    shUsuarios = ss.insertSheet(SHEETS.USUARIOS);
    shUsuarios.appendRow(['ID', 'Nome', 'Email', 'Cargo', 'DataCriacao']);
    shUsuarios.appendRow(['USR-1', 'Ten Nicolas', ADMIN_EMAIL, 'Gerente de Projetos', new Date()]);
    shUsuarios.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }

  // Garantir abas de Demandas Futuras e Anotações
  try { ensureDemandasFuturasSheet_(ss); } catch (errFut) {}
  try { ensureAnotacoesSheets_(ss); } catch (errNote) {}

  // 2. Tarefas
  let shTarefas = ss.getSheetByName(SHEETS.TAREFAS);
  if (!shTarefas) {
    shTarefas = ss.insertSheet(SHEETS.TAREFAS);
    shTarefas.appendRow([
      'ID', 'Titulo', 'Descricao', 'DevID', 'DevNome',
      'DiasEstimados', 'DiasTrabalhados', 'DataInicio', 'DataPrazo',
      'Status', 'Ativo', 'Urgencia', 'DataConclusao', 'Modulo'
    ]);

    const hoje = new Date();
    const prazo1 = new Date();
    prazo1.setDate(hoje.getDate() + 3);

    shTarefas.appendRow([
      'TSK-101', 'Criar API de Pagamento', 'Integração com gateway',
      'USR-1', 'Ten Nicolas', 4, 1, hoje, prazo1,
      'Em Andamento', 'SIM', 'Alta', '', 'SIOPLEEx'
    ]);
    shTarefas.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }

  // 3. Solicitações de Prazo
  let shSolicitacoes = ss.getSheetByName(SHEETS.SOLICITACOES);
  if (!shSolicitacoes) {
    shSolicitacoes = ss.insertSheet(SHEETS.SOLICITACOES);
    shSolicitacoes.appendRow([
      'ID', 'TarefaID', 'TarefaTitulo', 'DevID', 'DevNome',
      'Justificativa', 'NovosDiasSolicitados', 'Status',
      'GerenteNome', 'DataSolicitacao', 'DataResposta'
    ]);
    shSolicitacoes.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }
}

function calcularDiasExecucao(dataInicioRaw, dataConclusaoRaw, status, ativo) {
  if (status === 'Não Iniciada' || (!ativo && status !== 'Concluída' && status !== 'Em Andamento' && status !== 'Atrasada')) return 0;
  const dInicio = parseDataSegura_(dataInicioRaw);
  if (!dInicio) return (status === 'Não Iniciada') ? 0 : 1;
  dInicio.setHours(0, 0, 0, 0);

  let dFim;
  if (status === 'Concluída' && dataConclusaoRaw) {
    dFim = parseDataSegura_(dataConclusaoRaw) || new Date();
  } else {
    dFim = new Date();
  }
  dFim.setHours(0, 0, 0, 0);

  const diffMs = dFim.getTime() - dInicio.getTime();
  const diffDias = Math.round(diffMs / (1000 * 60 * 60 * 24));
  return Math.max(1, diffDias + 1);
}

/**
 * Lê os dados da planilha e vincula automaticamente o usuário autenticado ao seu perfil
 */
function getAppData(clientEmail) {
  let emailLogado = '';
  try {
    emailLogado = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  } catch (e) {}

  if (!emailLogado && clientEmail) {
    emailLogado = clientEmail.toLowerCase().trim();
  }

  if (!emailLogado) {
    return {
      autorizado: false,
      precisaLogin: true,
      emailAutenticado: '',
      adminEmail: ADMIN_EMAIL
    };
  }

  const check = verificarEmailAutorizado(emailLogado);
  if (!check.autorizado) {
    return {
      autorizado: false,
      precisaLogin: true,
      emailAutenticado: emailLogado,
      error: check.message,
      adminEmail: ADMIN_EMAIL
    };
  }

  const ss = getDatabaseSpreadsheet();
  initDatabaseIfNeeded();

  // 1. Usuários
  const shUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
  const usuarios = [];
  let currentUser = check.usuario || null;

  if (shUsuarios && shUsuarios.getLastRow() > 1) {
    const rawUsers = shUsuarios.getRange(2, 1, shUsuarios.getLastRow() - 1, 5).getValues();

    rawUsers.forEach(r => {
      if (!r[0]) return;

      const id = (r[0] || '').toString().trim();
      const nome = (r[1] || '').toString().trim();
      const email = (r[2] || '').toString().trim().toLowerCase();
      let cargo = (r[3] || 'Desenvolvedor').toString().trim();
      if (cargo === 'Ambos' || cargo === 'Gerente' || cargo.toLowerCase().includes('gerente') || cargo.toLowerCase().includes('ambos')) {
        cargo = 'Gerente de Projetos';
      } else {
        cargo = 'Desenvolvedor';
      }

      const userObj = {
        id: id,
        nome: nome,
        email: email,
        cargo: cargo
      };

      usuarios.push(userObj);

      if (email && emailLogado && email === emailLogado) {
        currentUser = userObj;
      }
    });
  }

  if (!currentUser && emailLogado === ADMIN_EMAIL.toLowerCase()) {
    const nickInSheet = usuarios.find(u => u.nome.toLowerCase().includes('nicolas'));
    if (nickInSheet) {
      currentUser = nickInSheet;
    } else {
      currentUser = {
        id: 'USR-1',
        nome: 'Ten Nicolas',
        email: ADMIN_EMAIL,
        cargo: 'Gerente de Projetos'
      };
    }
  }

  // 1.5 Processar Demandas Futuras (avisos 24h e entrada automática na esteira)
  try {
    processarDemandasFuturasAgendadas_(ss, usuarios);
  } catch (errFut) {
    Logger.log('Aviso processarDemandasFuturasAgendadas_: ' + errFut.message);
  }

  // 2. Tarefas
  const shTarefas = ss.getSheetByName(SHEETS.TAREFAS);
  const tarefas = [];
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  if (shTarefas && shTarefas.getLastRow() > 1) {
    const lastColTar = Math.max(14, shTarefas.getLastColumn());
    const rawTasks = shTarefas.getRange(2, 1, shTarefas.getLastRow() - 1, lastColTar).getValues();

    rawTasks.forEach(r => {
      if (!r[0]) return;

      let isAtivo = (r[10] === 'SIM' || r[10] === true);
      const dataPrazo = parseDataSegura_(r[8]);
      const dataInicio = parseDataSegura_(r[7]);
      const dataConclusao = parseDataSegura_(r[12]);
      let statusSalvo = (r[9] || '').toString().trim();
      let statusAtual = statusSalvo;

      if (statusSalvo === 'Concluída' || statusSalvo.toLowerCase().includes('conclu')) {
        statusAtual = 'Concluída';
        isAtivo = false;
      } else if (statusSalvo === 'Não Iniciada' && !dataPrazo) {
        statusAtual = 'Não Iniciada';
        isAtivo = false;
      } else if (dataPrazo) {
        const prazoSemHora = new Date(dataPrazo.getFullYear(), dataPrazo.getMonth(), dataPrazo.getDate(), 0, 0, 0);
        if (hoje.getTime() > prazoSemHora.getTime()) {
          statusAtual = 'Atrasada';
        } else {
          statusAtual = 'Em Andamento';
        }
      } else {
        statusAtual = 'Não Iniciada';
        isAtivo = false;
      }
      const diasAuto = calcularDiasExecucao(dataInicio, dataConclusao, statusAtual, isAtivo);
      const dataInicioIso = formatarDataIsoSegura_(dataInicio);
      const dataPrazoIso = formatarDataIsoSegura_(dataPrazo);
      const dataConclusaoIso = formatarDataIsoSegura_(dataConclusao);
      const modRaw = r[13] ? r[13].toString().trim() : 'SIOPLEEx';

      tarefas.push({
        id: r[0].toString().trim(),
        titulo: (r[1] || '').toString().trim(),
        descricao: (r[2] || '').toString(),
        devId: (r[3] || '').toString().trim(),
        devNome: (r[4] || '').toString().trim(),
        diasEstimados: Number(r[5]) || 0,
        diasTrabalhados: diasAuto,
        dataInicio: dataInicioIso,
        dataInsercao: dataInicioIso || formatarDataIsoSegura_(new Date()),
        dataPrazo: dataPrazoIso,
        status: statusAtual,
        ativo: isAtivo,
        urgencia: (r[11] || 'Media').toString().trim(),
        dataConclusao: dataConclusaoIso,
        modulo: modRaw,
        demandaExterna: (modRaw === 'Demanda Externa')
      });
    });
  }

  // 3. Solicitações
  const shSol = ss.getSheetByName(SHEETS.SOLICITACOES);
  const solicitacoes = [];
  const tz = obterTimeZone_();
  if (shSol && shSol.getLastRow() > 1) {
    const rawSol = shSol.getRange(2, 1, shSol.getLastRow() - 1, 11).getValues();
    rawSol.forEach(r => {
      if (r[0]) {
        solicitacoes.push({
          id: r[0],
          tarefaId: r[1],
          tarefaTitulo: r[2],
          devId: r[3],
          devNome: r[4],
          justificativa: r[5],
          novosDiasSolicitados: Number(r[6]) || 0,
          status: r[7],
          gerenteNome: r[8],
          dataSolicitacao: r[9] ? Utilities.formatDate(new Date(r[9]), tz, 'dd/MM/yyyy HH:mm') : '',
          dataResposta: r[10] ? Utilities.formatDate(new Date(r[10]), tz, 'dd/MM/yyyy HH:mm') : ''
        });
      }
    });
  }

  // 4. Demandas Futuras
  const demandasFuturas = lerDemandasFuturas_(ss);

  // 5. Anotações Pessoais e Abas do Usuário Logado
  const dadosAnotacoes = lerAnotacoesDoUsuario_(ss, currentUser, emailLogado);

  return {
    autorizado: true,
    precisaLogin: false,
    usuarios: usuarios,
    tarefas: tarefas,
    solicitacoes: solicitacoes,
    demandasFuturas: demandasFuturas,
    anotacoes: dadosAnotacoes.anotacoes,
    abasAnotacoes: dadosAnotacoes.abas,
    emailAutenticado: emailLogado,
    currentUser: currentUser,
    adminEmail: ADMIN_EMAIL,
    sysManagerEmail: SYSMANAGER_EMAIL
  };
}

function organizarPlanilhaUsuarios() {
  const ss = getDatabaseSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.USUARIOS);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.USUARIOS);
    sh.appendRow(['ID', 'Nome', 'Email', 'Cargo', 'DataCriacao']);
    sh.appendRow(['USR-1', 'Ten Nicolas', ADMIN_EMAIL, 'Gerente de Projetos', new Date()]);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    SpreadsheetApp.flush();
    return sh;
  }

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(5, sh.getLastColumn());

  if (lastRow < 1) {
    sh.appendRow(['ID', 'Nome', 'Email', 'Cargo', 'DataCriacao']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    SpreadsheetApp.flush();
    return sh;
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || '').toString().toLowerCase().trim());
  const headerJaCorreto = (headers[0] === 'id' && headers[1] === 'nome' && headers[2].includes('email') && headers[3].includes('cargo'));

  if (headerJaCorreto && lastRow > 1) {
    const range = sh.getRange(2, 1, lastRow - 1, 5);
    const rows = range.getValues();
    let mudou = false;
    for (let i = 0; i < rows.length; i++) {
      const nome = (rows[i][1] || '').toString().toLowerCase();
      let email = (rows[i][2] || '').toString().trim().toLowerCase();
      if (!email && nome.includes('nicolas')) {
        rows[i][2] = ADMIN_EMAIL;
        mudou = true;
      }
      let c = (rows[i][3] || '').toString().trim();
      if (c === 'Ambos' || c === 'Gerente' || c.toLowerCase().includes('gerente') || c.toLowerCase().includes('ambos')) {
        if (rows[i][3] !== 'Gerente de Projetos') {
          rows[i][3] = 'Gerente de Projetos';
          mudou = true;
        }
      } else if (!c || c.toLowerCase() === 'desenvolvedor' || c.toLowerCase() === 'dev') {
        if (rows[i][3] !== 'Desenvolvedor') {
          rows[i][3] = 'Desenvolvedor';
          mudou = true;
        }
      }
    }
    if (mudou) {
      range.setValues(rows);
      SpreadsheetApp.flush();
    }
    return sh;
  }

  return sh;
}

function cadastrarUsuario(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    organizarPlanilhaUsuarios();
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.USUARIOS);
    const newId = 'USR-' + (sh.getLastRow() + Math.floor(Math.random() * 900 + 100));
    const nomeLimpo = (dados.nome || '').trim();
    const emailLimpo = (dados.email || '').trim().toLowerCase();
    let cargoLimpo = (dados.cargo || 'Desenvolvedor').toString().trim();
    if (cargoLimpo === 'Ambos' || cargoLimpo === 'Gerente' || cargoLimpo.toLowerCase().includes('gerente') || cargoLimpo.toLowerCase().includes('ambos')) {
      cargoLimpo = 'Gerente de Projetos';
    } else {
      cargoLimpo = 'Desenvolvedor';
    }

    sh.appendRow([newId, nomeLimpo, emailLimpo, cargoLimpo, new Date()]);
    SpreadsheetApp.flush();

    return {
      success: true,
      id: newId,
      nome: nomeLimpo,
      email: emailLimpo,
      cargo: cargoLimpo
    };
  } catch (e) {
    Logger.log('Erro ao cadastrar usuário: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function atualizarUsuario(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    organizarPlanilhaUsuarios();
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.USUARIOS);
    if (!sh || sh.getLastRow() <= 1) return { success: false, message: 'Planilha sem membros' };

    const range = sh.getRange(2, 1, sh.getLastRow() - 1, 5);
    const rows = range.getValues();
    const targetId = (dados.id || '').toString().trim().toLowerCase().replace(/\s+/g, '');
    const nomeLimpo = (dados.nome || '').trim();
    const emailLimpo = (dados.email || '').trim().toLowerCase();
    let cargoLimpo = (dados.cargo || 'Desenvolvedor').toString().trim();
    if (cargoLimpo === 'Ambos' || cargoLimpo === 'Gerente' || cargoLimpo.toLowerCase().includes('gerente') || cargoLimpo.toLowerCase().includes('ambos')) {
      cargoLimpo = 'Gerente de Projetos';
    } else {
      cargoLimpo = 'Desenvolvedor';
    }
    let encontrado = false;

    for (let i = 0; i < rows.length; i++) {
      const rowId = (rows[i][0] || '').toString().trim().toLowerCase().replace(/\s+/g, '');
      if (rowId === targetId) {
        rows[i][1] = nomeLimpo;
        rows[i][2] = emailLimpo;
        rows[i][3] = cargoLimpo;
        encontrado = true;
        break;
      }
    }

    if (encontrado) {
      range.setValues(rows);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, message: 'Membro com ID "' + dados.id + '" não localizado' };
  } catch (e) {
    Logger.log('Erro ao atualizar usuário: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function excluirUsuario(userId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.USUARIOS);
    if (!sh || sh.getLastRow() <= 1) return { success: false, message: 'Planilha vazia' };

    const targetId = (userId || '').toString().trim().toLowerCase().replace(/\s+/g, '');
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();

    for (let i = 0; i < rows.length; i++) {
      const rowId = (rows[i][0] || '').toString().trim().toLowerCase().replace(/\s+/g, '');
      if (rowId === targetId) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, message: 'ID não encontrado' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function limparNomeDesenvolvedor_(nome) {
  if (!nome) return '';
  return nome.toString().replace(/\s*\((?:Desenvolvedor|Gerente de Projetos|Gerente|Ambos|Membro)\)\s*$/i, '').trim();
}

function cadastrarTarefa(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    const existingIds = new Set();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
        if (r[0]) existingIds.add(r[0].toString().trim().toUpperCase());
      });
    }
    let newId = 'TSK-' + (sh.getLastRow() + Math.floor(Math.random() * 900 + 100));
    while (existingIds.has(newId.toUpperCase())) {
      newId = 'TSK-' + (Date.now().toString().slice(-4)) + Math.floor(Math.random() * 90 + 10);
    }

    const tornarAtiva = dados.tornarAtiva === true || dados.tornarAtiva === 'true';
    const dias = tornarAtiva ? (parseInt(dados.diasEstimados, 10) || 1) : (parseInt(dados.diasEstimados, 10) || 0);
    const hoje = parseDataSegura_(dados.dataInicio) || new Date();
    let dataPrazo = parseDataSegura_(dados.dataPrazo) || '';

    if (!dataPrazo && tornarAtiva) {
      const p = new Date(hoje.getTime());
      p.setDate(p.getDate() + Math.max(1, dias));
      dataPrazo = p;
    }

    if (tornarAtiva && sh.getLastRow() > 1) {
      const range = sh.getRange(2, 1, sh.getLastRow() - 1, 13);
      const rows = range.getValues();
      rows.forEach(r => {
        if (r[3] === dados.devId && r[9] !== 'Concluída') {
          r[10] = 'NAO';
        }
      });
      range.setValues(rows);
    }

    const devNomeLimpo = limparNomeDesenvolvedor_(dados.devNome || '');
    const moduloFinal = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa')
      ? 'Demanda Externa'
      : (dados.modulo ? dados.modulo.toString().trim() : 'SIOPLEEx');

    sh.appendRow([
      newId,
      (dados.titulo || '').toString().trim(),
      dados.descricao || '',
      (dados.devId || '').toString().trim(),
      devNomeLimpo,
      dias,
      0,
      tornarAtiva ? hoje : (parseDataSegura_(dados.dataInicio) || ''),
      dataPrazo,
      tornarAtiva ? 'Em Andamento' : 'Não Iniciada',
      tornarAtiva ? 'SIM' : 'NAO',
      dados.urgencia || 'Media',
      '',
      moduloFinal
    ]);
    SpreadsheetApp.flush();

    return { success: true, id: newId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function selecionarTarefaAtiva(tarefaId, devId, dadosExtras) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false };

    const lastCol = Math.max(14, sh.getLastColumn());
    const range = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol);
    const rows = range.getValues();
    const hoje = new Date();

    const idBusca = (tarefaId || '').toString().trim().toUpperCase();
    rows.forEach(r => {
      const rowId = (r[0] ? r[0].toString().trim().toUpperCase() : '');
      if (rowId === idBusca) {
        r[10] = 'SIM';
        r[9] = 'Em Andamento';
        if (dadosExtras && dadosExtras.dataInicio) {
          r[7] = parseDataSegura_(dadosExtras.dataInicio) || hoje;
        } else if (!r[7]) {
          r[7] = hoje;
        }
        if (dadosExtras && dadosExtras.dataPrazo) {
          r[8] = parseDataSegura_(dadosExtras.dataPrazo) || r[8];
        }
        if (dadosExtras && dadosExtras.diasEstimados) {
          r[5] = Number(dadosExtras.diasEstimados) || 1;
        }
      } else if (r[3] === devId && r[9] !== 'Concluída') {
        r[10] = 'NAO';
      }
    });

    range.setValues(rows);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function concluirTarefa(tarefaId, devId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false };

    const lastCol = Math.max(14, sh.getLastColumn());
    const range = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol);
    const rows = range.getValues();

    const idBusca = (tarefaId || '').toString().trim().toUpperCase();
    rows.forEach(r => {
      const rowId = (r[0] ? r[0].toString().trim().toUpperCase() : '');
      if (rowId === idBusca) {
        r[9] = 'Concluída';
        r[10] = 'NAO';
        r[12] = new Date();
      }
    });

    range.setValues(rows);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function excluirTarefa(tarefaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false, message: 'Nenhuma demanda cadastrada' };

    const lastRow = sh.getLastRow();
    const range = sh.getRange(2, 1, lastRow - 1, 1);
    const rows = range.getValues();
    const idBusca = (tarefaId || '').toString().trim().toUpperCase();

    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] && rows[i][0].toString().trim().toUpperCase() === idBusca) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }

    return { success: false, message: 'Demanda não encontrada' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function trocarOrdemTarefas(id1, id2) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 2) return { success: true };

    const lastRow = sh.getLastRow();
    const lastCol = Math.max(14, sh.getLastColumn());
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    let row1 = -1;
    let row2 = -1;

    for (let i = 0; i < ids.length; i++) {
      const val = ids[i][0] ? ids[i][0].toString().trim().toUpperCase() : '';
      if (val === (id1 || '').toString().trim().toUpperCase()) row1 = i + 2;
      if (val === (id2 || '').toString().trim().toUpperCase()) row2 = i + 2;
    }

    if (row1 !== -1 && row2 !== -1 && row1 !== row2) {
      const vals1 = sh.getRange(row1, 1, 1, lastCol).getValues();
      const vals2 = sh.getRange(row2, 1, 1, lastCol).getValues();
      sh.getRange(row1, 1, 1, lastCol).setValues(vals2);
      sh.getRange(row2, 1, 1, lastCol).setValues(vals1);
      SpreadsheetApp.flush();
      return { success: true };
    }
    return { success: false, message: 'Linhas não encontradas' };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Atualiza todos os dados de uma demanda existente em qualquer aba (Esteira, Drawer ou Relatório Mensal)
 * e persiste imediatamente no Google Sheets sem deslocamento de fuso horário.
 */
function atualizarTarefa(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false, message: 'Nenhuma demanda cadastrada' };

    const lastRow = sh.getLastRow();
    const lastCol = Math.max(14, sh.getLastColumn());
    const range = sh.getRange(2, 1, lastRow - 1, lastCol);
    const rows = range.getValues();

    const tarefaId = (dados.id || dados.tarefaId || '').toString().trim();
    if (!tarefaId) return { success: false, message: 'ID da demanda não informado' };
    const idBusca = tarefaId.toUpperCase();

    let targetIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const rowId = (rows[i][0] ? rows[i][0].toString().trim() : '').toUpperCase();
      if (rowId === idBusca) {
        targetIdx = i;
        break;
      }
    }

    if (targetIdx === -1 && dados.tituloOriginal) {
      const titOrig = dados.tituloOriginal.toString().trim().toLowerCase();
      for (let i = 0; i < rows.length; i++) {
        if ((rows[i][1] || '').toString().trim().toLowerCase() === titOrig) {
          targetIdx = i;
          break;
        }
      }
    }

    if (targetIdx === -1) {
      return { success: false, message: 'Demanda não encontrada na planilha: ' + tarefaId };
    }

    const r = rows[targetIdx];

    if (dados.titulo !== undefined && dados.titulo !== null) {
      r[1] = dados.titulo.toString().trim();
    }
    if (dados.descricao !== undefined && dados.descricao !== null) {
      r[2] = dados.descricao.toString();
    }
    if (dados.devId !== undefined && dados.devId) {
      r[3] = dados.devId.toString().trim();
    }
    if (dados.devNome !== undefined && dados.devNome) {
      r[4] = limparNomeDesenvolvedor_(dados.devNome);
    }
    if (dados.diasEstimados !== undefined) {
      r[5] = parseInt(dados.diasEstimados, 10) || 0;
    }

    const prazoParsed = parseDataSegura_(dados.dataPrazo);
    if (dados.dataPrazo === '' || dados.status === 'Não Iniciada') {
      if (dados.status === 'Não Iniciada') {
        r[8] = '';
        r[5] = 0;
        r[6] = 0;
      } else if (prazoParsed) {
        r[8] = prazoParsed;
      }
    } else if (prazoParsed) {
      r[8] = prazoParsed;
      if (!r[7]) {
        r[7] = parseDataSegura_(dados.dataInicio) || new Date();
      }
    }

    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const dataPrazoAtual = parseDataSegura_(r[8]);

    if (dados.status === 'Concluída') {
      r[9] = 'Concluída';
      r[10] = 'NAO';
      if (prazoParsed) {
        r[12] = prazoParsed;
      } else if (!r[12]) {
        r[12] = new Date();
      }
      if (!r[7]) r[7] = new Date();
    } else if (dados.status === 'Não Iniciada' || (!dataPrazoAtual && r[9] !== 'Concluída')) {
      r[9] = 'Não Iniciada';
      r[10] = 'NAO';
      r[8] = '';
      r[12] = '';
    } else if (dataPrazoAtual) {
      const pZero = new Date(dataPrazoAtual.getFullYear(), dataPrazoAtual.getMonth(), dataPrazoAtual.getDate(), 0, 0, 0);
      r[9] = (hoje.getTime() > pZero.getTime()) ? 'Atrasada' : 'Em Andamento';
      r[12] = '';
      if (dados.ativo !== undefined) {
        r[10] = (dados.ativo === true || dados.ativo === 'true' || dados.ativo === 'SIM') ? 'SIM' : 'NAO';
      }
    }

    if (dados.urgencia) {
      r[11] = dados.urgencia.toString().trim();
    }

    const isExterna = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa');
    r[13] = isExterna ? 'Demanda Externa' : (dados.modulo ? dados.modulo.toString().trim() : 'SIOPLEEx');

    rows[targetIdx] = r;
    range.setValues(rows);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: 'Demanda atualizada com sucesso na planilha!',
      tarefa: {
        id: r[0].toString().trim(),
        titulo: r[1],
        descricao: r[2],
        devId: r[3],
        devNome: r[4],
        diasEstimados: Number(r[5]) || 0,
        dataInicio: formatarDataIsoSegura_(r[7]),
        dataPrazo: formatarDataIsoSegura_(r[8]),
        status: r[9],
        ativo: (r[10] === 'SIM' || r[10] === true),
        urgencia: r[11],
        dataConclusao: formatarDataIsoSegura_(r[12]),
        modulo: r[13],
        demandaExterna: (r[13] === 'Demanda Externa')
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function solicitarNovoPrazo(dados) {
  try {
    const ss = getDatabaseSpreadsheet();
    const shSol = ss.getSheetByName(SHEETS.SOLICITACOES);
    const newId = 'SOL-' + (shSol.getLastRow() + Math.floor(Math.random() * 900 + 100));

    shSol.appendRow([
      newId,
      dados.tarefaId,
      dados.tarefaTitulo,
      dados.devId,
      dados.devNome,
      dados.justificativa || '',
      parseInt(dados.novosDiasSolicitados, 10) || 1,
      'Pendente',
      '',
      new Date(),
      ''
    ]);
    SpreadsheetApp.flush();

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function darCientePrazo(solicitacaoId, gerenteNome) {
  try {
    const ss = getDatabaseSpreadsheet();
    const shSol = ss.getSheetByName(SHEETS.SOLICITACOES);
    const shTar = ss.getSheetByName(SHEETS.TAREFAS);

    let dadosSol = null;
    const rangeSol = shSol.getRange(2, 1, shSol.getLastRow() - 1, 11);
    const rowsSol = rangeSol.getValues();

    for (let i = 0; i < rowsSol.length; i++) {
      if (rowsSol[i][0] === solicitacaoId) {
        rowsSol[i][7] = 'Ciente';
        rowsSol[i][8] = gerenteNome;
        rowsSol[i][10] = new Date();
        dadosSol = {
          tarefaId: rowsSol[i][1],
          diasExtras: Number(rowsSol[i][6])
        };
        break;
      }
    }

    if (!dadosSol) return { success: false };
    rangeSol.setValues(rowsSol);

    if (shTar && shTar.getLastRow() > 1) {
      const rangeTar = shTar.getRange(2, 1, shTar.getLastRow() - 1, 13);
      const rowsTar = rangeTar.getValues();

      rowsTar.forEach(r => {
        if (r[0] === dadosSol.tarefaId) {
          const hoje = new Date();
          const novoPrazo = new Date(hoje);
          novoPrazo.setDate(hoje.getDate() + dadosSol.diasExtras);

          r[5] = Number(r[5]) + dadosSol.diasExtras;
          r[8] = novoPrazo;
          r[9] = 'Em Andamento';
          r[10] = 'SIM';
        }
      });

      rangeTar.setValues(rowsTar);
    }
    SpreadsheetApp.flush();

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// =========================================================================
// DEMANDAS FUTURAS (Agendamento, Aviso 24h por E-mail e Entrada na Esteira)
// =========================================================================

function ensureDemandasFuturasSheet_(ss) {
  let sh = ss.getSheetByName(SHEETS.DEMANDAS_FUTURAS);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.DEMANDAS_FUTURAS);
    sh.appendRow([
      'ID', 'Modulo', 'DemandaExterna', 'Titulo', 'DataHoraProgramada',
      'DevID', 'DevNome', 'DevEmail', 'Status',
      'Email24hEnviado', 'DataEmail24h', 'TarefaGeradaID', 'DataCriacao'
    ]);
    sh.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    sh.getRange('E:E').setNumberFormat('@');
  }
  return sh;
}

/**
 * Parsing determinístico de Data/Hora programada no fuso horário de Brasília (America/Sao_Paulo, UTC-3)
 */
function parseDataHoraProgramada_(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return isNaN(val.getTime()) ? null : val;
  }
  const str = val.toString().trim();
  if (!str) return null;

  const mIso = str.match(/^(\d{4})-(\d{2})-(\d{2})[T\s](\d{2}):(\d{2})/);
  if (mIso) {
    const isoComOffset = `${mIso[1]}-${mIso[2]}-${mIso[3]}T${mIso[4]}:${mIso[5]}:00-03:00`;
    const dOffset = new Date(isoComOffset);
    if (!isNaN(dOffset.getTime())) return dOffset;
    return new Date(
      parseInt(mIso[1], 10),
      parseInt(mIso[2], 10) - 1,
      parseInt(mIso[3], 10),
      parseInt(mIso[4], 10),
      parseInt(mIso[5], 10),
      0, 0
    );
  }

  const mBr = str.match(/^(\d{2})\/(\d{2})\/(\d{4})\s*(?:às\s*)?(\d{2}):(\d{2})/i);
  if (mBr) {
    const isoComOffset = `${mBr[3]}-${mBr[2]}-${mBr[1]}T${mBr[4]}:${mBr[5]}:00-03:00`;
    return new Date(isoComOffset);
  }

  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatarDataHoraIsoLocal_(val) {
  const dt = parseDataHoraProgramada_(val);
  if (!dt) return '';
  return Utilities.formatDate(dt, obterTimeZone_(), "yyyy-MM-dd'T'HH:mm");
}

function resolverEmailDesenvolvedor_(ss, devId, devNome, devEmailInformado) {
  if (devEmailInformado && devEmailInformado.toString().includes('@')) {
    return devEmailInformado.toString().trim().toLowerCase();
  }
  const idLimpo = (devId || '').toString().trim().toLowerCase();
  const nomeLimpo = limparNomeDesenvolvedor_(devNome || '').toLowerCase();

  try {
    const shUsr = ss.getSheetByName(SHEETS.USUARIOS);
    if (shUsr && shUsr.getLastRow() > 1) {
      const rows = shUsr.getRange(2, 1, shUsr.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const uId = (rows[i][0] || '').toString().trim().toLowerCase();
        const uNome = (rows[i][1] || '').toString().trim().toLowerCase();
        const uEmail = (rows[i][2] || '').toString().trim().toLowerCase();
        if (uEmail && uEmail.includes('@')) {
          if ((idLimpo && uId === idLimpo) || (nomeLimpo && (uNome === nomeLimpo || uNome.includes(nomeLimpo) || nomeLimpo.includes(uNome)))) {
            return uEmail;
          }
        }
      }
    }
  } catch (e) {}

  return ADMIN_EMAIL;
}

function lerDemandasFuturas_(ss) {
  const sh = ss.getSheetByName(SHEETS.DEMANDAS_FUTURAS);
  const lista = [];
  if (!sh || sh.getLastRow() <= 1) return lista;

  const raw = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
  raw.forEach(r => {
    if (!r[0]) return;
    const demandaExterna = (r[2] === 'SIM' || r[2] === true || r[1] === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : (r[1] ? r[1].toString().trim() : 'SIOPLEEx');
    lista.push({
      id: r[0].toString().trim(),
      modulo: modulo,
      demandaExterna: demandaExterna,
      titulo: r[3] ? r[3].toString().trim() : '',
      dataHoraProgramada: formatarDataHoraIsoLocal_(r[4]),
      devId: r[5] ? r[5].toString().trim() : '',
      devNome: r[6] ? r[6].toString().trim() : '',
      devEmail: r[7] ? r[7].toString().trim() : '',
      status: r[8] ? r[8].toString().trim() : 'Agendada',
      email24hEnviado: (r[9] === 'SIM' || r[9] === true),
      dataEmail24h: r[10] ? r[10].toString().trim() : '',
      tarefaGeradaId: r[11] ? r[11].toString().trim() : '',
      dataCriacao: r[12] ? formatarDataHoraIsoLocal_(r[12]) : ''
    });
  });

  lista.sort((a, b) => (a.dataHoraProgramada || '').localeCompare(b.dataHoraProgramada || ''));
  return lista;
}

function cadastrarDemandaFutura(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ensureDemandasFuturasSheet_(ss);

    const titulo = (dados.titulo || '').toString().trim();
    const dataHoraRaw = (dados.dataHoraProgramada || '').toString().trim();
    const dtProgramada = parseDataHoraProgramada_(dataHoraRaw);
    if (!titulo || !dtProgramada) {
      return { success: false, message: 'Nome da Demanda e Horário são obrigatórios.' };
    }

    const existingIds = new Set();
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
        if (r[0]) existingIds.add(r[0].toString().trim().toUpperCase());
      });
    }
    let newId = 'DFU-' + (Date.now().toString().slice(-5)) + Math.floor(Math.random() * 90 + 10);
    while (existingIds.has(newId.toUpperCase())) {
      newId = 'DFU-' + (Date.now().toString().slice(-5)) + Math.floor(Math.random() * 90 + 10);
    }

    const tz = obterTimeZone_();
    const demandaExterna = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : ((dados.modulo || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');
    const dataHoraIso = Utilities.formatDate(dtProgramada, tz, "yyyy-MM-dd'T'HH:mm");
    const agoraStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm");
    const devNomeLimpo = limparNomeDesenvolvedor_(dados.devNome || 'Ten Nicolas');
    const devEmail = resolverEmailDesenvolvedor_(ss, dados.devId, devNomeLimpo, dados.devEmail);

    sh.appendRow([
      newId,
      modulo,
      demandaExterna ? 'SIM' : 'NAO',
      titulo,
      dataHoraIso,
      dados.devId || 'USR-1',
      devNomeLimpo,
      devEmail,
      'Agendada',
      'NAO',
      '',
      '',
      agoraStr
    ]);
    sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(dataHoraIso);
    SpreadsheetApp.flush();

    try { garantirGatilhoDemandasFuturas_(); } catch (e) {}

    let procRes = { emailsEnviados: 0, promovidas: 0, emailErro: '' };
    try {
      procRes = processarDemandasFuturasAgendadas_(ss);
    } catch (errProc) {
      procRes.emailErro = errProc.message;
    }

    const todas = lerDemandasFuturas_(ss);
    const itemCriado = todas.find(x => x.id === newId) || null;

    return {
      success: true,
      id: newId,
      email24hEnviado: itemCriado ? itemCriado.email24hEnviado : false,
      dataEmail24h: itemCriado ? itemCriado.dataEmail24h : '',
      status: itemCriado ? itemCriado.status : 'Agendada',
      demanda: itemCriado,
      demandasFuturas: todas,
      emailErro: procRes.emailErro || ''
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function atualizarDemandaFutura(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ensureDemandasFuturasSheet_(ss);
    if (sh.getLastRow() <= 1) return { success: false, message: 'Demanda futura não encontrada.' };

    const targetId = (dados.id || '').toString().trim();
    const dtProgramada = parseDataHoraProgramada_(dados.dataHoraProgramada);
    if (!targetId || !dados.titulo || !dtProgramada) {
      return { success: false, message: 'Dados inválidos para atualização.' };
    }

    const tz = obterTimeZone_();
    const demandaExterna = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : ((dados.modulo || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');
    const dataHoraIso = Utilities.formatDate(dtProgramada, tz, "yyyy-MM-dd'T'HH:mm");
    const devNomeLimpo = limparNomeDesenvolvedor_(dados.devNome || '');
    const devEmailResolvido = resolverEmailDesenvolvedor_(ss, dados.devId, devNomeLimpo, dados.devEmail);

    const range = sh.getRange(2, 1, sh.getLastRow() - 1, 13);
    const rows = range.getValues();
    let encontrou = false;

    rows.forEach(r => {
      if (r[0] && r[0].toString().trim() === targetId) {
        const oldIso = formatarDataHoraIsoLocal_(r[4]);
        r[1] = modulo;
        r[2] = demandaExterna ? 'SIM' : 'NAO';
        r[3] = dados.titulo.toString().trim();
        r[4] = dataHoraIso;
        if (dados.devId) r[5] = dados.devId;
        if (devNomeLimpo) r[6] = devNomeLimpo;
        if (devEmailResolvido) r[7] = devEmailResolvido;
        if (oldIso !== dataHoraIso && r[8] === 'Agendada') {
          r[9] = 'NAO';
          r[10] = '';
        }
        encontrou = true;
      }
    });

    if (!encontrou) return { success: false, message: 'Demanda futura não encontrada.' };
    range.setValues(rows);
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('@');
    SpreadsheetApp.flush();

    let procRes = { emailErro: '' };
    try {
      procRes = processarDemandasFuturasAgendadas_(ss);
    } catch (errProc) {
      procRes.emailErro = errProc.message;
    }

    const todas = lerDemandasFuturas_(ss);
    const itemAtual = todas.find(x => x.id === targetId) || null;

    return {
      success: true,
      id: targetId,
      email24hEnviado: itemAtual ? itemAtual.email24hEnviado : false,
      dataEmail24h: itemAtual ? itemAtual.dataEmail24h : '',
      status: itemAtual ? itemAtual.status : 'Agendada',
      demanda: itemAtual,
      demandasFuturas: todas,
      emailErro: procRes.emailErro || ''
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function excluirDemandaFutura(demandaFuturaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.DEMANDAS_FUTURAS);
    if (!sh || sh.getLastRow() <= 1) return { success: true };

    const targetId = (demandaFuturaId || '').toString().trim();
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (vals[i][0] && vals[i][0].toString().trim() === targetId) {
        sh.deleteRow(i + 2);
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function enviarEmailAvisoDemandaFutura_(item, dtProgramada) {
  const tz = obterTimeZone_();
  const dataFmt = Utilities.formatDate(dtProgramada, tz, 'dd/MM/yyyy');
  const horaFmt = Utilities.formatDate(dtProgramada, tz, 'HH:mm');
  const devNome = limparNomeDesenvolvedor_(item.devNome || 'Responsável');
  const origem = item.demandaExterna ? 'Demanda Externa' : (item.modulo || 'SIOPLEEx');

  const destinatariosPrincipais = [];
  if (item.devEmail && item.devEmail.indexOf('@') > -1) {
    destinatariosPrincipais.push(item.devEmail.trim().toLowerCase());
  }
  if (!destinatariosPrincipais.includes(ADMIN_EMAIL.toLowerCase())) {
    destinatariosPrincipais.push(ADMIN_EMAIL.toLowerCase());
  }

  const toEmail = destinatariosPrincipais[0];
  const ccList = destinatariosPrincipais.slice(1);
  if (toEmail !== SYSMANAGER_EMAIL.toLowerCase() && !ccList.includes(SYSMANAGER_EMAIL.toLowerCase())) {
    ccList.push(SYSMANAGER_EMAIL.toLowerCase());
  }
  const ccEmail = ccList.join(',');

  const assunto = `[Sys Manager] Aviso 24h — Demanda Futura: ${item.titulo}`;
  const corpoTexto =
    `Olá, ${devNome}!\n\n` +
    `Lembrete automático do Sys Manager: a Demanda Futura abaixo entrará automaticamente na Esteira de Produção.\n\n` +
    `• Demanda: ${item.titulo}\n` +
    `• Módulo / Origem: ${origem}\n` +
    `• Responsável: ${devNome}\n` +
    `• Horário agendado para entrada na Esteira: ${dataFmt} às ${horaFmt}\n\n` +
    `No horário programado, a demanda será movida automaticamente para "Em Andamento" na Esteira de Produção.\n\n` +
    `Atenciosamente,\nSys Manager (${SYSMANAGER_EMAIL})`;

  const corpoHtml = `
    <div style="font-family: Arial, sans-serif; color: #0f172a; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background: linear-gradient(135deg, #0f172a 0%, #4f46e5 100%); color: #ffffff; padding: 18px 24px;">
        <h2 style="margin: 0; font-size: 18px;">Sys Manager — Aviso de Demanda Futura</h2>
        <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Entrada automática na Esteira de Produção (Aviso 24h)</p>
      </div>
      <div style="padding: 20px 24px; background-color: #ffffff;">
        <p style="margin-top: 0;">Olá, <strong>${devNome}</strong>!</p>
        <p>Esta é uma notificação automática informando sobre a seguinte <strong>Demanda Futura</strong> programada:</p>
        <div style="background: #f8fafc; border-left: 4px solid #4f46e5; padding: 14px 16px; border-radius: 6px; margin: 16px 0;">
          <p style="margin: 0 0 6px;"><strong>Demanda:</strong> ${item.titulo}</p>
          <p style="margin: 0 0 6px;"><strong>Módulo / Origem:</strong> ${origem}</p>
          <p style="margin: 0 0 6px;"><strong>Responsável:</strong> ${devNome}</p>
          <p style="margin: 0;"><strong>Horário programado:</strong> <span style="color: #4f46e5; font-weight: bold;">${dataFmt} às ${horaFmt}</span></p>
        </div>
        <p style="font-size: 13px; color: #475569; margin-bottom: 0;">
          Nesse horário a Demanda entrará automaticamente na Esteira de Produção com status <strong>Em Andamento</strong>.
        </p>
      </div>
      <div style="background: #f1f5f9; padding: 12px 24px; font-size: 12px; color: #64748b;">
        Sys Manager • ${SYSMANAGER_EMAIL}
      </div>
    </div>
  `;

  try {
    const payload = {
      to: toEmail,
      replyTo: SYSMANAGER_EMAIL,
      name: 'Sys Manager',
      subject: assunto,
      body: corpoTexto,
      htmlBody: corpoHtml
    };
    if (ccEmail) payload.cc = ccEmail;
    MailApp.sendEmail(payload);
  } catch (errMailApp) {
    GmailApp.sendEmail(toEmail, assunto, corpoTexto, {
      htmlBody: corpoHtml,
      name: 'Sys Manager',
      replyTo: SYSMANAGER_EMAIL,
      cc: ccEmail || ''
    });
  }
}

function processarDemandasFuturasAgendadas_(ssOpt, usuariosOpt) {
  const ss = ssOpt || getDatabaseSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DEMANDAS_FUTURAS);
  if (!sh || sh.getLastRow() <= 1) {
    return { success: true, emailsEnviados: 0, promovidas: 0, demandasFuturas: [] };
  }

  const tz = obterTimeZone_();
  const agora = new Date();
  const janelaAvisoMs = (24 * 60 * 60 * 1000) + (2 * 60 * 1000);

  const shTarefas = ss.getSheetByName(SHEETS.TAREFAS);
  const existingTaskIds = new Set();
  const devsComTarefaAtiva = new Set();
  if (shTarefas && shTarefas.getLastRow() > 1) {
    shTarefas.getRange(2, 1, shTarefas.getLastRow() - 1, 13).getValues().forEach(t => {
      if (t[0]) existingTaskIds.add(t[0].toString().trim().toUpperCase());
      if (t[3] && t[9] !== 'Concluída' && (t[10] === 'SIM' || t[10] === true)) {
        devsComTarefaAtiva.add(t[3].toString().trim());
      }
    });
  }

  const range = sh.getRange(2, 1, sh.getLastRow() - 1, 13);
  const rows = range.getValues();
  let alterou = false;
  let emailsEnviados = 0;
  let promovidas = 0;
  let ultimoErroEmail = '';

  rows.forEach(r => {
    if (!r[0]) return;
    const statusAtual = (r[8] || 'Agendada').toString().trim();
    if (statusAtual !== 'Agendada') return;

    const dtProgramada = parseDataHoraProgramada_(r[4]);
    if (!dtProgramada) return;

    const diffMs = dtProgramada.getTime() - agora.getTime();
    const devId = (r[5] || 'USR-1').toString().trim();
    const devNome = limparNomeDesenvolvedor_(r[6] || 'Ten Nicolas');
    const devEmail = resolverEmailDesenvolvedor_(ss, devId, devNome, r[7]);

    if (r[7] !== devEmail && devEmail) {
      r[7] = devEmail;
      alterou = true;
    }

    const demandaExterna = (r[2] === 'SIM' || r[2] === true || r[1] === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : ((r[1] || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');
    const titulo = (r[3] || '').toString().trim();

    // 1. Enviar aviso por e-mail 24 horas antes (ou imediatamente se agendada para menos de 24h)
    const emailJaEnviado = (r[9] === 'SIM' || r[9] === true);
    if (diffMs <= janelaAvisoMs && !emailJaEnviado) {
      try {
        enviarEmailAvisoDemandaFutura_({
          id: r[0],
          titulo: titulo,
          modulo: modulo,
          demandaExterna: demandaExterna,
          devNome: devNome,
          devEmail: devEmail
        }, dtProgramada);
        r[9] = 'SIM';
        r[10] = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm');
        alterou = true;
        emailsEnviados++;
      } catch (errMail) {
        ultimoErroEmail = errMail.message || errMail.toString();
        Logger.log('Erro ao enviar e-mail 24h Demanda Futura (' + r[0] + '): ' + ultimoErroEmail);
      }
    }

    // 2. Quando chegar o horário programado, entra automaticamente na Esteira de Produção
    if (diffMs <= 0 && shTarefas) {
      let newTaskId = 'TSK-' + (Date.now().toString().slice(-5)) + Math.floor(Math.random() * 90 + 10);
      while (existingTaskIds.has(newTaskId.toUpperCase())) {
        newTaskId = 'TSK-' + (Date.now().toString().slice(-5)) + Math.floor(Math.random() * 90 + 10);
      }
      existingTaskIds.add(newTaskId.toUpperCase());

      const dataIniLimpa = new Date(dtProgramada.getFullYear(), dtProgramada.getMonth(), dtProgramada.getDate(), 12, 0, 0);
      const prazoPadrao = new Date(dataIniLimpa.getTime());
      prazoPadrao.setDate(prazoPadrao.getDate() + 1);

      const jaTemAtiva = devsComTarefaAtiva.has(devId);
      if (devId && !jaTemAtiva) devsComTarefaAtiva.add(devId);

      const descAuto = 'Entrada automática na Esteira de Produção via Demanda Futura (' +
        Utilities.formatDate(dtProgramada, tz, 'dd/MM/yyyy HH:mm') + ').';

      shTarefas.appendRow([
        newTaskId,
        titulo,
        descAuto,
        devId,
        devNome,
        1,
        1,
        dataIniLimpa,
        prazoPadrao,
        'Em Andamento',
        jaTemAtiva ? 'NAO' : 'SIM',
        'Media',
        '',
        modulo
      ]);

      r[8] = 'Na Esteira';
      r[11] = newTaskId;
      alterou = true;
      promovidas++;
    }
  });

  if (alterou) {
    range.setValues(rows);
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('@');
    SpreadsheetApp.flush();
  }

  return {
    success: true,
    emailsEnviados: emailsEnviados,
    promovidas: promovidas,
    emailErro: ultimoErroEmail,
    demandasFuturas: lerDemandasFuturas_(ss)
  };
}

function processarDemandasFuturasAgendadas() {
  return processarDemandasFuturasAgendadas_(null, null);
}

/**
 * Permite disparar/reenviar manualmente o aviso de uma Demanda Futura
 */
function dispararEmailDemandaFuturaManual(demandaFuturaId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const sh = ensureDemandasFuturasSheet_(ss);
    if (sh.getLastRow() <= 1) return { success: false, message: 'Nenhuma Demanda Futura cadastrada.' };

    const tz = obterTimeZone_();
    const range = sh.getRange(2, 1, sh.getLastRow() - 1, 13);
    const rows = range.getValues();
    const targetId = (demandaFuturaId || '').toString().trim();

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (r[0] && r[0].toString().trim() === targetId) {
        const dtProgramada = parseDataHoraProgramada_(r[4]) || new Date();
        const devNome = limparNomeDesenvolvedor_(r[6] || 'Ten Nicolas');
        const devEmail = resolverEmailDesenvolvedor_(ss, r[5], devNome, r[7]);
        const demandaExterna = (r[2] === 'SIM' || r[2] === true || r[1] === 'Demanda Externa');
        const modulo = demandaExterna ? 'Demanda Externa' : ((r[1] || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');

        enviarEmailAvisoDemandaFutura_({
          id: r[0],
          titulo: r[3],
          modulo: modulo,
          demandaExterna: demandaExterna,
          devNome: devNome,
          devEmail: devEmail
        }, dtProgramada);

        r[7] = devEmail;
        r[9] = 'SIM';
        r[10] = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
        range.setValues(rows);
        SpreadsheetApp.flush();

        return {
          success: true,
          emailDestino: devEmail,
          dataEmail24h: r[10],
          demandasFuturas: lerDemandasFuturas_(ss)
        };
      }
    }
    return { success: false, message: 'Demanda Futura não localizada.' };
  } catch (e) {
    return { success: false, message: 'Falha ao enviar e-mail: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function garantirGatilhoDemandasFuturas_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('GATILHO_DFU_ATIVO') === 'SIM') return;
  const fnName = 'processarDemandasFuturasAgendadas';
  const triggers = ScriptApp.getProjectTriggers();
  const jaExiste = triggers.some(t => t.getHandlerFunction() === fnName);
  if (!jaExiste) {
    ScriptApp.newTrigger(fnName).timeBased().everyMinutes(5).create();
  }
  props.setProperty('GATILHO_DFU_ATIVO', 'SIM');
}

function instalarGatilhoDemandasFuturas() {
  const fnName = 'processarDemandasFuturasAgendadas';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger(fnName)
    .timeBased()
    .everyMinutes(5)
    .create();
  PropertiesService.getScriptProperties().setProperty('GATILHO_DFU_ATIVO', 'SIM');
  return { success: true, message: 'Gatilho configurado a cada 5 minutos para Demandas Futuras.' };
}

// =========================================================================
// ANOTAÇÕES PESSOAIS DO USUÁRIO (Persistência Individual no Banco de Dados)
// =========================================================================

function ensureAnotacoesSheets_(ss) {
  let shNotas = ss.getSheetByName(SHEETS.ANOTACOES);
  if (!shNotas) {
    shNotas = ss.insertSheet(SHEETS.ANOTACOES);
    shNotas.appendRow([
      'ID', 'UsuarioID', 'UsuarioEmail', 'Aba', 'Titulo',
      'Conteudo', 'Cor', 'Fixado', 'DataCriacao', 'DataAtualizacao'
    ]);
    shNotas.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }

  let shAbas = ss.getSheetByName(SHEETS.ANOTACOES_ABAS);
  if (!shAbas) {
    shAbas = ss.insertSheet(SHEETS.ANOTACOES_ABAS);
    shAbas.appendRow(['UsuarioID', 'UsuarioEmail', 'AbasJSON', 'DataAtualizacao']);
    shAbas.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }
  return { shNotas: shNotas, shAbas: shAbas };
}

function pertenciaAoUsuario_(rowUserId, rowUserEmail, userId, userEmail) {
  const rEmail = (rowUserEmail || '').toString().trim().toLowerCase();
  const uEmail = (userEmail || '').toString().trim().toLowerCase();
  if (rEmail && uEmail && rEmail === uEmail) return true;

  const rId = (rowUserId || '').toString().trim().toUpperCase();
  const uId = (userId || '').toString().trim().toUpperCase();
  if (rId && uId && rId === uId) return true;

  return false;
}

function lerAnotacoesDoUsuario_(ss, currentUser, emailLogado) {
  const { shNotas, shAbas } = ensureAnotacoesSheets_(ss);
  const userId = currentUser && currentUser.id ? currentUser.id : 'USR-1';
  const userEmail = (currentUser && currentUser.email ? currentUser.email : (emailLogado || ADMIN_EMAIL)).toLowerCase().trim();

  const anotacoes = [];
  if (shNotas && shNotas.getLastRow() > 1) {
    const raw = shNotas.getRange(2, 1, shNotas.getLastRow() - 1, 10).getValues();
    raw.forEach(r => {
      if (!r[0]) return;
      if (pertenciaAoUsuario_(r[1], r[2], userId, userEmail)) {
        anotacoes.push({
          id: r[0].toString().trim(),
          aba: (r[3] || 'Geral').toString().trim() || 'Geral',
          titulo: (r[4] || '').toString(),
          conteudo: (r[5] || '').toString(),
          cor: (r[6] || 'padrao').toString().trim(),
          fixado: (r[7] === 'SIM' || r[7] === true),
          dataCriacao: r[8] ? r[8].toString() : new Date().toISOString(),
          dataAtualizacao: r[9] ? r[9].toString() : new Date().toISOString()
        });
      }
    });
  }

  let abas = ['Geral'];
  if (shAbas && shAbas.getLastRow() > 1) {
    const rawAbas = shAbas.getRange(2, 1, shAbas.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < rawAbas.length; i++) {
      if (pertenciaAoUsuario_(rawAbas[i][0], rawAbas[i][1], userId, userEmail)) {
        try {
          const parsed = JSON.parse(rawAbas[i][2] || '["Geral"]');
          if (Array.isArray(parsed) && parsed.length > 0) {
            abas = parsed;
            if (!abas.includes('Geral')) abas.unshift('Geral');
          }
        } catch (e) {}
        break;
      }
    }
  }

  anotacoes.forEach(n => {
    if (n.aba && !abas.includes(n.aba)) abas.push(n.aba);
  });

  return { anotacoes: anotacoes, abas: abas };
}

function salvarAnotacaoNoBanco(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const { shNotas } = ensureAnotacoesSheets_(ss);

    const nota = dados.anotacao || dados;
    const notaId = (nota.id || ('NOTE-' + Date.now())).toString().trim();
    const usuarioId = (dados.usuarioId || nota.usuarioId || 'USR-1').toString().trim();
    const usuarioEmail = (dados.usuarioEmail || nota.usuarioEmail || ADMIN_EMAIL).toString().trim().toLowerCase();
    const aba = (nota.aba || 'Geral').toString().trim() || 'Geral';
    const titulo = (nota.titulo || 'Sem Título').toString().trim();
    const conteudo = (nota.conteudo || '').toString();
    const cor = (nota.cor || 'padrao').toString().trim();
    const fixadoStr = (nota.fixado === true || nota.fixado === 'true' || nota.fixado === 'SIM') ? 'SIM' : 'NAO';
    const agoraIso = new Date().toISOString();
    const dataCriacao = nota.dataCriacao || agoraIso;
    const dataAtualizacao = nota.dataAtualizacao || agoraIso;

    if (shNotas.getLastRow() > 1) {
      const range = shNotas.getRange(2, 1, shNotas.getLastRow() - 1, 10);
      const rows = range.getValues();
      for (let i = 0; i < rows.length; i++) {
        if (rows[i][0] && rows[i][0].toString().trim() === notaId) {
          rows[i][1] = usuarioId;
          rows[i][2] = usuarioEmail;
          rows[i][3] = aba;
          rows[i][4] = titulo;
          rows[i][5] = conteudo;
          rows[i][6] = cor;
          rows[i][7] = fixadoStr;
          rows[i][9] = dataAtualizacao;
          range.setValues(rows);
          SpreadsheetApp.flush();
          return { success: true, id: notaId };
        }
      }
    }

    shNotas.appendRow([
      notaId,
      usuarioId,
      usuarioEmail,
      aba,
      titulo,
      conteudo,
      cor,
      fixadoStr,
      dataCriacao,
      dataAtualizacao
    ]);
    SpreadsheetApp.flush();
    return { success: true, id: notaId };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function excluirAnotacaoDoBanco(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const { shNotas } = ensureAnotacoesSheets_(ss);
    if (shNotas.getLastRow() <= 1) return { success: true };

    const targetId = (dados.anotacaoId || dados.id || '').toString().trim();
    const rows = shNotas.getRange(2, 1, shNotas.getLastRow() - 1, 1).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      if (rows[i][0] && rows[i][0].toString().trim() === targetId) {
        shNotas.deleteRow(i + 2);
      }
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function salvarAbasAnotacoesNoBanco(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const { shAbas } = ensureAnotacoesSheets_(ss);

    const usuarioId = (dados.usuarioId || 'USR-1').toString().trim();
    const usuarioEmail = (dados.usuarioEmail || ADMIN_EMAIL).toString().trim().toLowerCase();
    let abas = Array.isArray(dados.abas) ? dados.abas : ['Geral'];
    if (!abas.includes('Geral')) abas.unshift('Geral');
    const abasJson = JSON.stringify(abas);
    const agoraIso = new Date().toISOString();

    if (shAbas.getLastRow() > 1) {
      const range = shAbas.getRange(2, 1, shAbas.getLastRow() - 1, 4);
      const rows = range.getValues();
      for (let i = 0; i < rows.length; i++) {
        if (pertenciaAoUsuario_(rows[i][0], rows[i][1], usuarioId, usuarioEmail)) {
          rows[i][0] = usuarioId;
          rows[i][1] = usuarioEmail;
          rows[i][2] = abasJson;
          rows[i][3] = agoraIso;
          range.setValues(rows);
          SpreadsheetApp.flush();
          return { success: true };
        }
      }
    }

    shAbas.appendRow([usuarioId, usuarioEmail, abasJson, agoraIso]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sincronizarAnotacoesUsuarioNoBanco(dados) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = getDatabaseSpreadsheet();
    const { shNotas } = ensureAnotacoesSheets_(ss);

    const usuarioId = (dados.usuarioId || 'USR-1').toString().trim();
    const usuarioEmail = (dados.usuarioEmail || ADMIN_EMAIL).toString().trim().toLowerCase();
    const notasCliente = Array.isArray(dados.anotacoes) ? dados.anotacoes : [];

    if (Array.isArray(dados.abas)) {
      salvarAbasAnotacoesNoBanco({
        usuarioId: usuarioId,
        usuarioEmail: usuarioEmail,
        abas: dados.abas
      });
    }

    const outrasLinhas = [];
    if (shNotas.getLastRow() > 1) {
      const rows = shNotas.getRange(2, 1, shNotas.getLastRow() - 1, 10).getValues();
      rows.forEach(r => {
        if (r[0] && !pertenciaAoUsuario_(r[1], r[2], usuarioId, usuarioEmail)) {
          outrasLinhas.push(r);
        }
      });
    }

    const agoraIso = new Date().toISOString();
    const novasLinhasUsuario = notasCliente.map(n => [
      (n.id || ('NOTE-' + Date.now() + '-' + Math.floor(Math.random() * 1000))).toString().trim(),
      usuarioId,
      usuarioEmail,
      (n.aba || 'Geral').toString().trim() || 'Geral',
      (n.titulo || 'Sem Título').toString().trim(),
      (n.conteudo || '').toString(),
      (n.cor || 'padrao').toString().trim(),
      (n.fixado === true || n.fixado === 'true' || n.fixado === 'SIM') ? 'SIM' : 'NAO',
      n.dataCriacao || agoraIso,
      n.dataAtualizacao || agoraIso
    ]);

    const todasFinal = outrasLinhas.concat(novasLinhasUsuario);
    if (shNotas.getLastRow() > 1) {
      shNotas.getRange(2, 1, shNotas.getLastRow() - 1, 10).clearContent();
    }
    if (todasFinal.length > 0) {
      shNotas.getRange(2, 1, todasFinal.length, 10).setValues(todasFinal);
    }
    SpreadsheetApp.flush();

    return { success: true, total: novasLinhasUsuario.length };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
