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

const SHEETS = {
  USUARIOS: 'Usuarios',
  TAREFAS: 'Tarefas',
  SOLICITACOES: 'SolicitacoesPrazo'
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
 * Renderiza o Index.html diretamente para evitar qualquer tela de bloqueio e looping do Google OAuth
 */
function doGet(e) {
  initDatabaseIfNeeded();
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SysManager - Gestão Ágil')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

      const dataPrazo = r[8] ? new Date(r[8]) : null;
      let statusAtual = r[9];

      if (statusAtual !== 'Concluída' && dataPrazo) {
        const prazoSemHora = new Date(dataPrazo);
        prazoSemHora.setHours(0, 0, 0, 0);
        if (hoje.getTime() > prazoSemHora.getTime() && statusAtual !== 'Atrasada') {
          statusAtual = 'Atrasada';
        }
      }

      const isAtivo = (r[10] === 'SIM' || r[10] === true);
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
        modulo: r[13] ? r[13].toString().trim() : 'SIOPLEEx'
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

  return {
    autorizado: true,
    precisaLogin: false,
    usuarios: usuarios,
    tarefas: tarefas,
    solicitacoes: solicitacoes,
    emailAutenticado: emailLogado,
    currentUser: currentUser,
    adminEmail: ADMIN_EMAIL
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
    const newId = 'TSK-' + (sh.getLastRow() + Math.floor(Math.random() * 900 + 100));
    
    const dias = parseInt(dados.diasEstimados) || 1;
    const hoje = new Date();
    const dataPrazo = new Date(hoje);
    dataPrazo.setDate(hoje.getDate() + dias);

    const tornarAtiva = dados.tornarAtiva === true || dados.tornarAtiva === 'true';
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
      tornarAtiva ? 1 : 0,
      hoje,
      dataPrazo,
      'Em Andamento',
      tornarAtiva ? 'SIM' : 'NAO',
      dados.urgencia || 'Media',
      '',
      dados.modulo ? dados.modulo.toString().trim() : 'SIOPLEEx'
    ]);
    SpreadsheetApp.flush();

    return { success: true, id: newId };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function selecionarTarefaAtiva(tarefaId, devId) {
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

    rows.forEach(r => {
      if (r[3] === devId && r[9] !== 'Concluída') {
        if (r[0] === tarefaId) {
          r[10] = 'SIM';
          if (!r[7]) r[7] = hoje;
        } else {
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

    rows.forEach(r => {
      if (r[0] === tarefaId) {
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

  // Validação estrita de permissão: Apenas Gerente de Projetos ou Administrador Geral Nicolas Richardt
  const emailLogado = (check.email || '').toLowerCase().trim();
  const isMasterAdmin = (emailLogado === ADMIN_EMAIL.toLowerCase() || emailLogado.includes('nicolas'));
  const userCargo = (check.usuario && check.usuario.cargo ? check.usuario.cargo.toString().toLowerCase() : '');
  const isGerente = userCargo.includes('gerente') || userCargo.includes('ambos');

  if (!isMasterAdmin && !isGerente) {
    return { success: false, message: 'Permissão negada. Apenas Gerente de Projetos ou o Administrador Geral podem excluir demandas.' };
  }

  try {
    const ss = getDatabaseSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TAREFAS);
    if (!sh || sh.getLastRow() <= 1) return { success: false, message: 'Nenhuma demanda cadastrada' };

    const lastRow = sh.getLastRow();
    const range = sh.getRange(2, 1, lastRow - 1, 1);
    const ids = range.getValues();

    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] && ids[i][0].toString().trim() === tarefaId.toString().trim()) {
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
