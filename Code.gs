/**
 * Plataforma Municipal do Associativismo — Município de Gouveia
 * Backend único para os formulários nativos do portal.
 *
 * 1) Preencha INTERNAL_EMAILS com Vereadora, Vice-Presidente e técnico responsável.
 * 2) Implemente como Aplicação Web: Executar como "Eu" e acesso "Qualquer pessoa".
 * 3) Copie o URL /exec para config.js no site.
 */

const INTERNAL_EMAILS = [
  'COLOCAR_EMAIL_VEREADORA',
  'COLOCAR_EMAIL_VICE_PRESIDENTE',
  'COLOCAR_O_SEU_EMAIL'
];

// Formulários que podem receber submissões neste momento.
// Cultura e Desporto ficam preparados na plataforma, mas bloqueados até abertura por aviso.
const ACTIVE_FORMS = {
  rma: true,
  apoio: true,
  transporte: true,
  cultura: false,
  desporto: false,
  impulso: 'auto'
};

const FORM_NAMES = {
  rma: 'RMA',
  apoio: 'Apoio Técnico e Material',
  transporte: 'Transporte Municipal',
  cultura: 'Apoio Anual - Cultura',
  desporto: 'Apoio Anual - Desporto',
  impulso: 'Programa de Impulso Associativo'
};

function doGet() {
  return HtmlService.createHtmlOutput('Plataforma Municipal do Associativismo — serviço ativo.');
}

function doPost(e) {
  try {
    if (!e || !e.parameter || !e.parameter.payload) {
      return resposta_('Pedido sem dados.', false);
    }

    const payload = JSON.parse(e.parameter.payload);
    const tipo = String(payload.formType || '').trim();
    if (!FORM_NAMES[tipo]) return resposta_('Tipo de formulário inválido.', false);
    if (!formularioAberto_(tipo)) return resposta_('Este procedimento não se encontra aberto.', false);

    const dados = payload.data || {};
    const ficheiros = Array.isArray(payload.files) ? payload.files : [];
    const associacao = valor_(dados.associacao || dados.nome || dados.entidade || 'Entidade não identificada');
    const email = valor_(dados.email_associacao || dados.email || '');
    const agora = new Date();

    const store = obterArmazenamento_();
    const pastaPedido = criarPastaPedido_(store.folder, FORM_NAMES[tipo], associacao, agora);
    const links = guardarFicheiros_(pastaPedido, ficheiros);

    registarNaFolha_(store.sheet, tipo, payload.formTitle || FORM_NAMES[tipo], associacao, email, dados, links, agora);
    enviarResumoInterno_(tipo, associacao, email, dados, links, agora);
    enviarConfirmacao_(tipo, associacao, email, dados);

    return resposta_('Pedido registado com sucesso.', true);
  } catch (err) {
    console.error(err);
    return resposta_('Não foi possível registar o pedido: ' + err.message, false);
  }
}

function formularioAberto_(tipo) {
  const estado = ACTIVE_FORMS[tipo];
  if (estado === true) return true;
  if (estado === false) return false;
  if (tipo === 'impulso' && estado === 'auto') {
    const d = new Date();
    const y = d.getFullYear();
    const a1 = new Date(y, 0, 15, 0, 0, 0);
    const b1 = new Date(y, 2, 15, 23, 59, 59);
    const a2 = new Date(y, 5, 1, 0, 0, 0);
    const b2 = new Date(y, 5, 30, 23, 59, 59);
    return (d >= a1 && d <= b1) || (d >= a2 && d <= b2);
  }
  return false;
}

function obterArmazenamento_() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('MASTER_SHEET_ID');
  let folderId = props.getProperty('ROOT_FOLDER_ID');

  let sheet;
  if (!sheetId) {
    sheet = SpreadsheetApp.create('Plataforma Municipal do Associativismo — Pedidos');
    sheetId = sheet.getId();
    props.setProperty('MASTER_SHEET_ID', sheetId);
  } else {
    sheet = SpreadsheetApp.openById(sheetId);
  }

  let folder;
  if (!folderId) {
    folder = DriveApp.createFolder('Plataforma Municipal do Associativismo — Anexos');
    folderId = folder.getId();
    props.setProperty('ROOT_FOLDER_ID', folderId);
  } else {
    folder = DriveApp.getFolderById(folderId);
  }
  return { sheet, folder };
}

function criarPastaPedido_(root, formName, associacao, data) {
  const f1 = subpasta_(root, limparNome_(formName));
  const ano = subpasta_(f1, String(data.getFullYear()));
  const nome = Utilities.formatDate(data, Session.getScriptTimeZone() || 'Europe/Lisbon', 'yyyy-MM-dd_HHmmss') + ' - ' + limparNome_(associacao);
  return ano.createFolder(nome.substring(0, 120));
}

function subpasta_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function guardarFicheiros_(folder, ficheiros) {
  const links = [];
  ficheiros.forEach(f => {
    if (!f || !f.data || !f.name) return;
    const bytes = Utilities.base64Decode(f.data);
    const blob = Utilities.newBlob(bytes, f.mime || 'application/octet-stream', limparNome_(f.name));
    const file = folder.createFile(blob);
    links.push({ campo: f.field || 'anexo', nome: file.getName(), url: file.getUrl() });
  });
  return links;
}

function registarNaFolha_(book, tipo, titulo, associacao, email, dados, links, agora) {
  const nomeAba = FORM_NAMES[tipo].substring(0, 90);
  let sh = book.getSheetByName(nomeAba);
  if (!sh) {
    sh = book.insertSheet(nomeAba);
    sh.appendRow(['Data/Hora', 'Tipo', 'Associação', 'Email', 'Resumo', 'Anexos', 'Dados completos (JSON)']);
    sh.setFrozenRows(1);
  }

  const resumo = resumoTexto_(tipo, dados);
  const anexos = links.map(x => x.nome + ': ' + x.url).join('\n');
  sh.appendRow([
    agora,
    titulo,
    associacao,
    email,
    resumo,
    anexos,
    JSON.stringify(dados)
  ]);
  sh.autoResizeColumn(1);
}

function enviarResumoInterno_(tipo, associacao, email, dados, links, agora) {
  const configurados = INTERNAL_EMAILS
    .map(x => String(x || '').trim())
    .filter(x => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x));

  const conta = Session.getEffectiveUser().getEmail();
  if (conta && !configurados.includes(conta)) configurados.push(conta);
  if (!configurados.length) return;

  const nomeForm = FORM_NAMES[tipo];
  const assunto = `[Associativismo] ${nomeForm} — ${associacao}`;
  const linhas = resumoTexto_(tipo, dados);
  const anexos = links.length ? '\n\nANEXOS\n' + links.map(x => `• ${x.nome}: ${x.url}`).join('\n') : '';
  const corpo = `Foi submetido um novo pedido na Plataforma Municipal do Associativismo.\n\nPROCEDIMENTO\n${nomeForm}\n\nENTIDADE\n${associacao}\n${email ? 'Email: ' + email + '\n' : ''}\nRESUMO\n${linhas}${anexos}\n\nData: ${Utilities.formatDate(agora, Session.getScriptTimeZone() || 'Europe/Lisbon', 'dd/MM/yyyy HH:mm')}\n\nA submissão não equivale a aprovação.`;

  MailApp.sendEmail({
    to: configurados.join(','),
    subject: assunto,
    body: corpo,
    replyTo: email || undefined,
    name: 'Plataforma Municipal do Associativismo'
  });
}

function enviarConfirmacao_(tipo, associacao, email, dados) {
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
  const nomeForm = FORM_NAMES[tipo];
  const assunto = `Confirmação de receção — ${nomeForm}`;
  const corpo = `Exmos. Senhores,\n\nConfirmamos a receção da submissão efetuada através da Plataforma Municipal do Associativismo do Município de Gouveia.\n\nEntidade: ${associacao}\nProcedimento: ${nomeForm}\n\n${resumoCurto_(tipo, dados)}\n\nA receção do pedido não confere aprovação automática. O processo será analisado pelos serviços municipais competentes.\n\nCom os melhores cumprimentos,\nMunicípio de Gouveia`;
  MailApp.sendEmail({ to: email, subject: assunto, body: corpo, name: 'Município de Gouveia' });
}

function resumoTexto_(tipo, d) {
  const linhas = [];
  add_(linhas, 'RMA', d.numero_rma);
  if (tipo === 'apoio') {
    add_(linhas, 'Atividade', d.atividade); add_(linhas, 'Data', d.data); add_(linhas, 'Local', d.local); add_(linhas, 'Tipo de apoio', d.tipo_apoio); add_(linhas, 'Descrição', d.descricao);
  } else if (tipo === 'transporte') {
    for (let i=1;i<=10;i++) if (d[`d${i}_atividade`]) {
      linhas.push(`Deslocação ${i}: ${d[`d${i}_atividade`]} | ${d[`d${i}_data`] || ''} | ${d[`d${i}_destino`] || ''} | ${d[`d${i}_passageiros`] || ''} passageiros | ${d[`d${i}_partida`] || ''}–${d[`d${i}_regresso`] || ''}`);
    }
  } else if (tipo === 'impulso') {
    for (let i=1;i<=2;i++) if (d[`p${i}_titulo`]) linhas.push(`Projeto ${i}: ${d[`p${i}_titulo`]} | Custo: ${d[`p${i}_custo`] || '—'} € | Solicitado: ${d[`p${i}_pedido`] || '—'} € | ${d[`p${i}_periodo`] || ''}`);
  } else if (tipo === 'cultura' || tipo === 'desporto') {
    add_(linhas, 'Enquadramento', d.objetivos || d.modalidade);
    const prefixo = tipo === 'desporto' ? 's' : 'a';
    for (let i=1;i<=15;i++) if (d[`${prefixo}${i}_designacao`]) linhas.push(`Atividade ${i}: ${d[`${prefixo}${i}_designacao`]} | ${d[`${prefixo}${i}_periodo`] || ''} | ${d[`${prefixo}${i}_local`] || ''}`);
  } else if (tipo === 'rma') {
    add_(linhas, 'Procedimento', d.procedimento); add_(linhas, 'NIF/NIPC', d.nif); add_(linhas, 'Representante', d.representante); add_(linhas, 'Telefone', d.telefone);
  }
  return linhas.join('\n') || 'Consulte os dados completos na folha de registo.';
}

function resumoCurto_(tipo, d) {
  if (tipo === 'apoio') return `Atividade: ${valor_(d.atividade)}\nData: ${valor_(d.data)}\nLocal: ${valor_(d.local)}`;
  if (tipo === 'transporte') return `Primeira deslocação: ${valor_(d.d1_atividade)}\nData: ${valor_(d.d1_data)}\nDestino: ${valor_(d.d1_destino)}`;
  if (tipo === 'impulso') return `Projeto: ${valor_(d.p1_titulo)}\nValor solicitado: ${valor_(d.p1_pedido)} €`;
  return 'A submissão ficou registada para análise.';
}

function add_(arr, rotulo, valor) { if (valor !== undefined && valor !== null && String(valor).trim() !== '') arr.push(rotulo + ': ' + valor_(valor)); }
function valor_(v) { return Array.isArray(v) ? v.join(', ') : String(v == null ? '' : v); }
function limparNome_(s) { return String(s || 'Sem nome').replace(/[\\/:*?"<>|#%{}~&]/g, '-').replace(/\s+/g, ' ').trim(); }
function resposta_(mensagem, ok) { return HtmlService.createHtmlOutput(`<meta charset="utf-8"><body style="font-family:Arial;padding:20px"><strong>${ok ? 'OK' : 'ERRO'}</strong><p>${mensagem}</p></body>`); }
