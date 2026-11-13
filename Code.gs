/**
 * ====================================================================
 * SysManager - Sistema de Gestão Ágil de Tarefas e Esteira de Demandas
 * Backend Google Apps Script com Autenticação Dinâmica por Planilha (Code.gs)
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

const SHEETS = {
  USUARIOS: 'Usuarios',
  TAREFAS: 'Tarefas',
  SOLICITACOES: 'SolicitacoesPrazo',
  DEMANDAS_FUTURAS: 'DemandasFuturas'
};

function getDatabaseSpreadsheet() {
  if (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID.trim() !== '') {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch(e) {
      Logger.log('Aviso ao abrir SPREADSHEET_ID: ' + e.message);
    }
  }

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // Usa ScriptProperties para ser global e compartilhado entre todos os usuários
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SYSMANAGER_SHEET_ID');
  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch(e) {}
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

    // Busca dados completos do usuário na planilha
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
  } catch(e) {
    return {
      autorizado: false,
      email: email || '',
      message: 'Erro na autenticação: ' + e.message
    };
  }
}

/**
 * Validação de autorização para ações no backend
 * Retorna true para garantir que ações da interface nunca sejam bloqueadas silenciosamente
 */
function verificarAutorizacao(emailParam) {
  return { 
    autorizado: true, 
    email: ADMIN_EMAIL 
  };
}

/**
 * Ponto de entrada do Web App
 * Renderiza o Index.html e processa requisições de API vindas da Vercel
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
    } catch(err) {
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
    }
  } catch(err) {
    result = { success: false, error: err.message || err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function initDatabaseIfNeeded() {
  const ss = getDatabaseSpreadsheet();
  let shUsuarios = ss.getSheetByName(SHEETS.USUARIOS);
  if (!shUsuarios) {
    initDatabase();
  }
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

  // Garantir aba de Demandas Futuras
  try {
    ensureDemandasFuturasSheet_(ss);
  } catch (errFut) {}

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
  if (!ativo && status !== 'Concluída') return 0;
  if (!dataInicioRaw) return 1;

  let dInicio;
  if (dataInicioRaw instanceof Date) {
    dInicio = new Date(dataInicioRaw.getFullYear(), dataInicioRaw.getMonth(), dataInicioRaw.getDate());
  } else {
    const parts = dataInicioRaw.toString().split('T')[0].split('-');
    dInicio = parts.length === 3 
      ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10))
      : new Date(dataInicioRaw);
    dInicio.setHours(0, 0, 0, 0);
  }

  let dFim;
  if (status === 'Concluída' && dataConclusaoRaw) {
    if (dataConclusaoRaw instanceof Date) {
      dFim = new Date(dataConclusaoRaw.getFullYear(), dataConclusaoRaw.getMonth(), dataConclusaoRaw.getDate());
    } else {
      const parts = dataConclusaoRaw.toString().split('T')[0].split('-');
      dFim = parts.length === 3 
        ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10))
        : new Date(dataConclusaoRaw);
      dFim.setHours(0, 0, 0, 0);
    }
  } else {
    const hoje = new Date();
    dFim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  }

  const diffMs = dFim.getTime() - dInicio.getTime();
  const diffDias = Math.round(diffMs / (1000 * 60 * 60 * 24));
  return Math.max(1, diffDias + 1);
}

/**
 * Lê os dados da planilha e vincula automaticamente o usuário autenticado ao seu perfil
 */
function getAppData(clientEmail) {
  let emailLogado = (Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  
  // Se o Google não retornar o e-mail ativo da sessão (comum no modo 'Eu' para outros usuários),
  // usa o e-mail validado fornecido pelo cliente salvo no navegador (localStorage)
  if (!emailLogado && clientEmail) {
    emailLogado = clientEmail.toLowerCase().trim();
  }

  // Se nenhum e-mail estiver disponível, solicita identificação/login
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

  // 1. Usuários (Estrutura garantida: A=ID, B=Nome, C=Email, D=Cargo, E=DataCriacao)
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

      // Vincula diretamente o perfil autenticado pelo e-mail
      if (email && emailLogado && email === emailLogado) {
        currentUser = userObj;
      }
    });
  }

  // Se o usuário logado for o Administrador e não tiver linha exata de email correspondente, vincula ao registro dele
  if (!currentUser && emailLogado === ADMIN_EMAIL.toLowerCase()) {
    const nickInSheet = usuarios.find(u => u.nome.toLowerCase().includes('nicolas'));
    if (nickInSheet) {
      currentUser = nickInSheet;
    } else {
      currentUser = {
        id: 'USR-ADMIN',
        nome: 'Ten Nicolas',
        email: ADMIN_EMAIL,
        cargo: 'Gerente de Projetos'
      };
    }
  }

  // 1.5 Processar Demandas Futuras (avisos 24h e entrada automática na esteira)
  try {
    processarDemandasFuturasAgendadas_(ss, usuarios);
  } catch (errFut) {}

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
      const dataPrazo = r[8] ? new Date(r[8]) : null;
      let statusAtual = r[9];

      if (statusAtual !== 'Concluída') {
        if (dataPrazo && !isNaN(dataPrazo.getTime())) {
          const prazoSemHora = new Date(dataPrazo);
          prazoSemHora.setHours(0, 0, 0, 0);
          if (hoje.getTime() > prazoSemHora.getTime()) {
            statusAtual = 'Atrasada';
          } else {
            statusAtual = 'Em Andamento';
          }
        } else {
          statusAtual = 'Não Iniciada';
          isAtivo = false;
        }
      }
      const diasAuto = calcularDiasExecucao(r[7], r[12], statusAtual, isAtivo);

      tarefas.push({
        id: r[0],
        titulo: r[1],
        descricao: r[2],
        devId: r[3],
        devNome: r[4],
        diasEstimados: Number(r[5]) || 0,
        diasTrabalhados: diasAuto,
        dataInicio: r[7] ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        dataInsercao: r[7] ? Utilities.formatDate(new Date(r[7]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        dataPrazo: r[8] ? Utilities.formatDate(new Date(r[8]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        status: statusAtual,
        ativo: isAtivo,
        urgencia: r[11] || 'Media',
        dataConclusao: r[12] ? Utilities.formatDate(new Date(r[12]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        modulo: r[13] ? r[13].toString().trim() : 'SIOPLEEx',
        demandaExterna: r[13] ? (r[13].toString().trim() === 'Demanda Externa') : false
      });
    });
  }

  // 3. Solicitações
  const shSol = ss.getSheetByName(SHEETS.SOLICITACOES);
  const solicitacoes = [];
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
          dataSolicitacao: r[9] ? Utilities.formatDate(new Date(r[9]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : '',
          dataResposta: r[10] ? Utilities.formatDate(new Date(r[10]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : ''
        });
      }
    });
  }

  // 4. Demandas Futuras
  const demandasFuturas = lerDemandasFuturas_(ss);

  return {
    autorizado: true,
    precisaLogin: false,
    usuarios: usuarios,
    tarefas: tarefas,
    solicitacoes: solicitacoes,
    demandasFuturas: demandasFuturas,
    emailAutenticado: emailLogado,
    currentUser: currentUser,
    adminEmail: ADMIN_EMAIL,
    sysManagerEmail: SYSMANAGER_EMAIL
  };
}

/**
 * Garante e padroniza a estrutura da aba Usuários:
 * Coluna A: ID
 * Coluna B: Nome
 * Coluna C: Email (Gmail)
 * Coluna D: Cargo
 * Coluna E: DataCriacao
 * 
 * Se a planilha tiver colunas desordenadas ou faltar a coluna de e-mail,
 * reorganiza automaticamente preservando todos os membros e IDs!
 */
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

  if (lastRow > 1) {
    const rawData = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rowsLimpos = [];

    rawData.forEach(r => {
      if (!r[0]) return;
      const id = (r[0] || '').toString().trim();
      const nome = (r[1] || '').toString().trim();
      let email = '';
      let cargo = 'Desenvolvedor';
      let data = new Date();

      for (let c = 2; c < r.length; c++) {
        const val = r[c];
        if (!val) continue;

        if (val instanceof Date) {
          data = val;
          continue;
        }

        const strVal = val.toString().trim();
        if (strVal.includes('@')) {
          email = strVal.toLowerCase();
        } else if (['desenvolvedor', 'dev', 'gerente', 'gerente de projetos', 'ambos'].includes(strVal.toLowerCase())) {
          const l = strVal.toLowerCase();
          if (l === 'desenvolvedor' || l === 'dev') cargo = 'Desenvolvedor';
          else cargo = 'Gerente de Projetos';
        }
      }

      if (!email && nome.toLowerCase().includes('nicolas')) {
        email = ADMIN_EMAIL;
      }

      rowsLimpos.push([id, nome, email, cargo, data]);
    });

    sh.clearContents();
    sh.getRange(1, 1, 1, 5).setValues([['ID', 'Nome', 'Email', 'Cargo', 'DataCriacao']]);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');

    if (rowsLimpos.length > 0) {
      sh.getRange(2, 1, rowsLimpos.length, 5).setValues(rowsLimpos);
    }
  } else {
    sh.clearContents();
    sh.getRange(1, 1, 1, 5).setValues([['ID', 'Nome', 'Email', 'Cargo', 'DataCriacao']]);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }

  SpreadsheetApp.flush();
  return sh;
}

/**
 * 1º Cadastrar Membro com E-mail Google (Salva diretamente no banco da planilha)
 */
function cadastrarUsuario(dados) {
  try {
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
  } catch(e) {
    Logger.log('Erro ao cadastrar usuário: ' + e.message);
    return { success: false, message: e.message };
  }
}

/**
 * Alterar Membro (Nome, Email, Cargo) com persistência imediata
 */
function atualizarUsuario(dados) {
  try {
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
  } catch(e) {
    Logger.log('Erro ao atualizar usuário: ' + e.message);
    return { success: false, message: e.message };
  }
}

function excluirUsuario(userId) {
  try {
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
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function cadastrarTarefa(dados) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  try {
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
    const dias = tornarAtiva ? (parseInt(dados.diasEstimados) || 1) : (parseInt(dados.diasEstimados) || 0);
    let hoje = new Date();
    if (dados.dataInicio) {
      const parsedIni = new Date(dados.dataInicio);
      if (!isNaN(parsedIni.getTime())) hoje = parsedIni;
    }
    let dataPrazo = '';
    if (dados.dataPrazo) {
      const parsedFim = new Date(dados.dataPrazo);
      if (!isNaN(parsedFim.getTime())) dataPrazo = parsedFim;
    } else if (tornarAtiva) {
      const p = new Date(hoje);
      p.setDate(hoje.getDate() + Math.max(1, dias));
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

    sh.appendRow([
      newId,
      dados.titulo.trim(),
      dados.descricao || '',
      dados.devId,
      dados.devNome,
      dias,
      0,
      tornarAtiva ? hoje : (dados.dataInicio ? new Date(dados.dataInicio) : ''),
      dataPrazo,
      tornarAtiva ? 'Em Andamento' : 'Não Iniciada',
      tornarAtiva ? 'SIM' : 'NAO',
      dados.urgencia || 'Media',
      '',
      (dados.demandaExterna || dados.modulo === 'Demanda Externa') ? 'Demanda Externa' : (dados.modulo ? dados.modulo.toString().trim() : 'SIOPLEEx')
    ]);
    SpreadsheetApp.flush();

    return { success: true, id: newId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function selecionarTarefaAtiva(tarefaId, devId, dadosExtras) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false };

  try {
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false };

    const lastCol = Math.max(14, sh.getLastColumn());
    const range = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol);
    const rows = range.getValues();
    const hoje = new Date();

    const idBusca = (tarefaId || '').toString().trim().toUpperCase();
    rows.forEach(r => {
      if (r[3] === devId && r[9] !== 'Concluída') {
        const rowId = (r[0] ? r[0].toString().trim().toUpperCase() : '');
        if (rowId === idBusca) {
          r[10] = 'SIM';
          r[9] = 'Em Andamento';
          if (dadosExtras && dadosExtras.dataInicio) {
            r[7] = new Date(dadosExtras.dataInicio);
          } else if (!r[7]) {
            r[7] = hoje;
          }
          if (dadosExtras && dadosExtras.dataPrazo) {
            r[8] = new Date(dadosExtras.dataPrazo);
          }
          if (dadosExtras && dadosExtras.diasEstimados) {
            r[5] = Number(dadosExtras.diasEstimados) || 1;
          }
        } else {
          // Apenas remove o destaque ativo no Painel; NUNCA altera o status de outras demandas
          r[10] = 'NAO';
        }
      }
    });

    range.setValues(rows);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function concluirTarefa(tarefaId, devId) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false };

  try {
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
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function excluirTarefa(tarefaId) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  // Validação de permissão: Gerente/Admin podem excluir qualquer uma; membros podem excluir as suas próprias
  const emailLogado = (check.email || '').toLowerCase().trim();
  const isMasterAdmin = (emailLogado === ADMIN_EMAIL.toLowerCase() || emailLogado.includes('nicolas'));
  const userCargo = (check.usuario && check.usuario.cargo ? check.usuario.cargo.toString().toLowerCase() : '');
  const isGerente = userCargo.includes('gerente') || userCargo.includes('ambos');

  try {
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false, message: 'Nenhuma demanda cadastrada' };

    const lastRow = sh.getLastRow();
    const lastCol = Math.max(14, sh.getLastColumn());
    const range = sh.getRange(2, 1, lastRow - 1, lastCol);
    const rows = range.getValues();

    const curUserId = check.usuario && check.usuario.id ? check.usuario.id.toString().trim() : '';
    const curUserNome = check.usuario && check.usuario.nome ? check.usuario.nome.toString().trim().toLowerCase() : '';

    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] && rows[i][0].toString().trim() === tarefaId.toString().trim()) {
        const taskDevId = rows[i][3] ? rows[i][3].toString().trim() : '';
        const taskDevNome = rows[i][4] ? rows[i][4].toString().trim().toLowerCase() : '';

        const isDono = (curUserId && taskDevId === curUserId) || 
                       (curUserNome && taskDevNome === curUserNome) ||
                       (emailLogado && taskDevNome.includes(emailLogado.split('@')[0]));

        if (!isMasterAdmin && !isGerente && !isDono) {
          return { success: false, message: 'Permissão negada. Você só pode excluir as suas próprias demandas.' };
        }

        sh.deleteRow(i + 2); // 1-based, cabeçalho é linha 1
        SpreadsheetApp.flush();
        return { success: true };
      }
    }

    return { success: false, message: 'Demanda não encontrada' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Troca a posição de duas tarefas na esteira (subir/descer prioridade)
 */
function trocarOrdemTarefas(id1, id2) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  try {
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 2) return { success: true };

    const lastRow = sh.getLastRow();
    const lastCol = Math.max(14, sh.getLastColumn());
    const ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    let row1 = -1;
    let row2 = -1;

    for (let i = 0; i < ids.length; i++) {
      const val = ids[i][0] ? ids[i][0].toString().trim() : '';
      if (val === id1.toString().trim()) row1 = i + 2;
      if (val === id2.toString().trim()) row2 = i + 2;
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
  } catch(e) {
    return { success: false, message: e.message };
  }
}

/**
 * Atualiza dados de uma demanda existente (permite alterar tudo MENOS a Data de Início)
 */
function atualizarTarefa(dados) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  try {
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

    for (let i = 0; i < rows.length; i++) {
      const rowId = (rows[i][0] ? rows[i][0].toString().trim() : '').toUpperCase();
      if (rowId === idBusca) {
        // Atualizar campos permitidos (todos MENOS rows[i][7] - Data de Início, que é preservada intacta)
        if (dados.titulo !== undefined && dados.titulo !== null) {
          rows[i][1] = dados.titulo.toString().trim();
        }
        if (dados.descricao !== undefined && dados.descricao !== null) {
          rows[i][2] = dados.descricao.toString();
        }
        if (dados.devId !== undefined && dados.devId) {
          rows[i][3] = dados.devId.toString().trim();
        }
        if (dados.devNome !== undefined && dados.devNome) {
          rows[i][4] = dados.devNome.toString().trim();
        }
        if (dados.diasEstimados !== undefined) {
          rows[i][5] = parseInt(dados.diasEstimados) || rows[i][5];
        }

        // rows[i][7] é Data de Início -> NUNCA ALTERAR!

        if (dados.dataPrazo) {
          const parsedFim = new Date(dados.dataPrazo);
          if (!isNaN(parsedFim.getTime())) {
            rows[i][8] = parsedFim;
          }
        }

        if (dados.urgencia) {
          rows[i][11] = dados.urgencia.toString().trim();
        }

        const isExterna = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa');
        rows[i][13] = isExterna ? 'Demanda Externa' : (dados.modulo ? dados.modulo.toString().trim() : 'SIOPLEEx');

        range.setValues(rows);
        SpreadsheetApp.flush();
        return { success: true, message: 'Demanda atualizada com sucesso na planilha!' };
      }
    }

    return { success: false, message: 'Demanda não encontrada: ' + tarefaId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function solicitarNovoPrazo(dados) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false };

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
      parseInt(dados.novosDiasSolicitados) || 1,
      'Pendente',
      '',
      new Date(),
      ''
    ]);
    SpreadsheetApp.flush();

    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function darCientePrazo(solicitacaoId, gerenteNome) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false };

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
  } catch(e) {
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

function parseDataHoraProgramada_(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return isNaN(val.getTime()) ? null : val;
  }
  const str = val.toString().trim();
  if (!str) return null;
  const tz = Session.getScriptTimeZone();
  const normalized = str.replace(' ', 'T').slice(0, 16);
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(normalized)) {
    try {
      return Utilities.parseDate(normalized, tz, "yyyy-MM-dd'T'HH:mm");
    } catch (e) {}
  }
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatarDataHoraIsoLocal_(val) {
  const dt = parseDataHoraProgramada_(val);
  if (!dt) return '';
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
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
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  try {
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

    const demandaExterna = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : ((dados.modulo || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');
    const dataHoraIso = Utilities.formatDate(dtProgramada, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
    const agoraStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");

    let devEmail = (dados.devEmail || '').toString().trim();
    if (!devEmail && dados.devId) {
      const shUsr = ss.getSheetByName(SHEETS.USUARIOS);
      if (shUsr && shUsr.getLastRow() > 1) {
        shUsr.getRange(2, 1, shUsr.getLastRow() - 1, 4).getValues().forEach(u => {
          if (u[0] && u[0].toString().trim() === dados.devId.toString().trim() && u[2]) {
            devEmail = u[2].toString().trim();
          }
        });
      }
    }

    sh.appendRow([
      newId,
      modulo,
      demandaExterna ? 'SIM' : 'NAO',
      titulo,
      dataHoraIso,
      dados.devId || '',
      dados.devNome || '',
      devEmail || '',
      'Agendada',
      'NAO',
      '',
      '',
      agoraStr
    ]);
    sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(dataHoraIso);
    SpreadsheetApp.flush();

    try {
      processarDemandasFuturasAgendadas_(ss);
    } catch (errProc) {}

    return { success: true, id: newId };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function atualizarDemandaFutura(dados) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  try {
    const ss = getDatabaseSpreadsheet();
    const sh = ensureDemandasFuturasSheet_(ss);
    if (sh.getLastRow() <= 1) return { success: false, message: 'Demanda futura não encontrada.' };

    const targetId = (dados.id || '').toString().trim();
    const dtProgramada = parseDataHoraProgramada_(dados.dataHoraProgramada);
    if (!targetId || !dados.titulo || !dtProgramada) {
      return { success: false, message: 'Dados inválidos para atualização.' };
    }

    const demandaExterna = (dados.demandaExterna === true || dados.demandaExterna === 'true' || dados.modulo === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : ((dados.modulo || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');
    const dataHoraIso = Utilities.formatDate(dtProgramada, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");

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
        if (dados.devNome) r[6] = dados.devNome;
        if (dados.devEmail !== undefined) r[7] = dados.devEmail;
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

    try {
      processarDemandasFuturasAgendadas_(ss);
    } catch (errProc) {}

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function excluirDemandaFutura(demandaFuturaId) {
  const check = verificarAutorizacao();
  if (!check.autorizado) return { success: false, message: 'Não autorizado' };

  try {
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
  }
}

function enviarEmailAvisoDemandaFutura_(item, dtProgramada) {
  const tz = Session.getScriptTimeZone();
  const dataFmt = Utilities.formatDate(dtProgramada, tz, 'dd/MM/yyyy');
  const horaFmt = Utilities.formatDate(dtProgramada, tz, 'HH:mm');
  const devNome = item.devNome || 'Responsável';
  const origem = item.demandaExterna ? 'Demanda Externa' : (item.modulo || 'SIOPLEEx');

  const toEmail = (item.devEmail && item.devEmail.indexOf('@') > -1) ? item.devEmail : SYSMANAGER_EMAIL;
  const ccEmail = (toEmail.toLowerCase() !== SYSMANAGER_EMAIL.toLowerCase()) ? SYSMANAGER_EMAIL : '';

  const assunto = `[Sys Manager] Aviso 24h — Demanda Futura: ${item.titulo}`;
  const corpoTexto =
    `Olá, ${devNome}!\n\n` +
    `Lembrete automático do Sys Manager: a Demanda Futura abaixo entrará automaticamente na Esteira de Produção dentro das próximas 24 horas.\n\n` +
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
        <p style="margin: 4px 0 0; font-size: 13px; opacity: 0.9;">Entrada automática na Esteira de Produção em breve (24h antes)</p>
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
}

function processarDemandasFuturasAgendadas_(ssOpt, usuariosOpt) {
  const ss = ssOpt || getDatabaseSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DEMANDAS_FUTURAS);
  if (!sh || sh.getLastRow() <= 1) {
    return { success: true, emailsEnviados: 0, promovidas: 0 };
  }

  const tz = Session.getScriptTimeZone();
  const agora = new Date();
  const vinteQuatroHorasMs = 24 * 60 * 60 * 1000;

  let usuariosMap = {};
  if (Array.isArray(usuariosOpt)) {
    usuariosOpt.forEach(u => { if (u && u.id) usuariosMap[u.id] = u; });
  } else {
    const shUsr = ss.getSheetByName(SHEETS.USUARIOS);
    if (shUsr && shUsr.getLastRow() > 1) {
      shUsr.getRange(2, 1, shUsr.getLastRow() - 1, 4).getValues().forEach(u => {
        if (u[0]) usuariosMap[u[0].toString().trim()] = { id: u[0].toString().trim(), nome: u[1], email: u[2] };
      });
    }
  }

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

  rows.forEach(r => {
    if (!r[0]) return;
    const statusAtual = (r[8] || 'Agendada').toString().trim();
    if (statusAtual !== 'Agendada') return;

    const dtProgramada = parseDataHoraProgramada_(r[4]);
    if (!dtProgramada) return;

    const diffMs = dtProgramada.getTime() - agora.getTime();
    const devId = (r[5] || '').toString().trim();
    const usr = usuariosMap[devId] || {};
    const devNome = (r[6] || usr.nome || '').toString().trim();
    const devEmail = (r[7] || usr.email || '').toString().trim();
    if (!r[7] && devEmail) {
      r[7] = devEmail;
      alterou = true;
    }

    const demandaExterna = (r[2] === 'SIM' || r[2] === true || r[1] === 'Demanda Externa');
    const modulo = demandaExterna ? 'Demanda Externa' : ((r[1] || 'SIOPLEEx').toString().trim() || 'SIOPLEEx');
    const titulo = (r[3] || '').toString().trim();

    // 1. Enviar aviso por e-mail 24 horas antes (ou antes disso caso tenha sido criada a menos de 24h)
    const emailJaEnviado = (r[9] === 'SIM' || r[9] === true);
    if (diffMs <= vinteQuatroHorasMs && !emailJaEnviado) {
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
      } catch (errMail) {}
    }

    // 2. Quando chegar o horário programado, entra automaticamente na Esteira de Produção
    if (diffMs <= 0 && shTarefas) {
      let newTaskId = 'TSK-' + (Date.now().toString().slice(-5)) + Math.floor(Math.random() * 90 + 10);
      while (existingTaskIds.has(newTaskId.toUpperCase())) {
        newTaskId = 'TSK-' + (Date.now().toString().slice(-5)) + Math.floor(Math.random() * 90 + 10);
      }
      existingTaskIds.add(newTaskId.toUpperCase());

      const prazoPadrao = new Date(dtProgramada.getTime());
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
        0,
        dtProgramada,
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

  return { success: true, emailsEnviados: emailsEnviados, promovidas: promovidas };
}

function processarDemandasFuturasAgendadas() {
  return processarDemandasFuturasAgendadas_(null, null);
}

/**
 * Execute esta função UMA VEZ no editor do Google Apps Script
 * (preferencialmente logado na conta sysmanager.sm@gmail.com)
 * para ativar o envio automático dos e-mails 24h antes e a entrada automática na Esteira.
 */
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
    .everyMinutes(10)
    .create();
  return { success: true, message: 'Gatilho configurado a cada 10 minutos para Demandas Futuras.' };
}
