/**
 * FriPremium — Sistema de Chamados
 * Backend em Google Apps Script, vinculado à planilha "FriPremium — Chamados".
 *
 * Implantação:
 *   1. Colar este arquivo no editor (Extensões > Apps Script) da planilha.
 *   2. Executar a função setup() uma vez (cria abas, listas e usuários).
 *   3. Implantar > Nova implantação > App da Web
 *        Executar como: Eu  |  Quem pode acessar: Qualquer pessoa
 *   4. Copiar a URL /exec para a constante API do index.html.
 *
 * Leituras públicas: GET  ?acao=opcoes | ?acao=consultar&protocolo=...
 * Demais ações:      POST (Content-Type text/plain) com JSON {acao, ...}
 */

var ABA = {
  CHAMADOS: 'Chamados',
  HIST: 'Historico',
  USU: 'Usuarios',
  LOCAIS: 'Setores',
  DEST: 'AreasDestino',
  CAT: 'Categorias',
  CFG: 'Config'
};

var COLS = ['Protocolo', 'DataAbertura', 'Solicitante', 'SetorSolicitante', 'Local', 'DetalheLocal',
  'AreaDestino', 'Categoria', 'Urgencia', 'Descricao', 'FotoId', 'Status', 'DataInicio',
  'DataResolucao', 'Resolucao', 'UltimaAtualizacao', 'AtualizadoPor'];

var COLS_HIST = ['Protocolo', 'DataHora', 'Usuario', 'StatusAnterior', 'StatusNovo', 'Comentario'];
var COLS_USU = ['CPF', 'Nome', 'Perfil', 'AreaDestino', 'SenhaHash', 'Salt', 'Ativo'];

var STATUS = ['Aberto', 'Em andamento', 'Aguardando material/terceiro', 'Resolvido', 'Cancelado'];
var STATUS_ABERTOS = ['Aberto', 'Em andamento', 'Aguardando material/terceiro'];
var URGENCIAS = ['Crítica', 'Alta', 'Normal'];

var PASTA_FOTOS = 'FriPremium Chamados - Fotos';
var SESSAO_SEG = 6 * 3600;          // sessão de 6 horas
var MAX_TENTATIVAS = 5;             // erros de senha antes do bloqueio
var BLOQUEIO_SEG = 15 * 60;         // bloqueio de 15 minutos
var MAX_FOTO_BYTES = 3 * 1024 * 1024;

/* ===================== SETUP ===================== */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('America/Sao_Paulo');

  criarAba_(ss, ABA.CHAMADOS, COLS, []);
  criarAba_(ss, ABA.HIST, COLS_HIST, []);

  criarAba_(ss, ABA.USU, COLS_USU, [
    ['14917223717', 'Jeferson', 'gestor', 'Manutenção', '', '', 'Sim'],
    ['06612663111', 'Lucas', 'diretor', '*', '', '', 'Sim']
  ]);
  // CPF como texto, para preservar o zero à esquerda
  ss.getSheetByName(ABA.USU).getRange('A:A').setNumberFormat('@');

  criarAba_(ss, ABA.DEST, ['Area', 'Prefixo', 'Ativo'], [
    ['Manutenção', 'MAN', 'Sim']
  ]);

  criarAba_(ss, ABA.CAT, ['AreaDestino', 'Categoria', 'Ativo'], [
    ['Manutenção', 'Elétrica', 'Sim'],
    ['Manutenção', 'Hidráulica', 'Sim'],
    ['Manutenção', 'Civil / Predial', 'Sim'],
    ['Manutenção', 'Refrigeração', 'Sim'],
    ['Manutenção', 'Equipamentos de produção', 'Sim'],
    ['Manutenção', 'Outros', 'Sim']
  ]);

  criarAba_(ss, ABA.LOCAIS, ['Setor', 'Ativo'], [
    ['Administrativo', 'Sim'], ['Recepção / Portaria', 'Sim'], ['Recebimento', 'Sim'],
    ['Desossa', 'Sim'], ['Câmaras frias', 'Sim'], ['Embalagem', 'Sim'],
    ['Expedição / Carregamento', 'Sim'], ['Almoxarifado', 'Sim'], ['Refeitório', 'Sim'],
    ['Vestiários / Banheiros', 'Sim'], ['Área externa / Pátio', 'Sim'], ['Outro', 'Sim']
  ]);

  criarAba_(ss, ABA.CFG, ['Chave', 'Valor', 'Descricao'], [
    ['SLA_Critica_h', 24, 'Prazo-alvo em horas para urgência Crítica'],
    ['SLA_Alta_h', 72, 'Prazo-alvo em horas para urgência Alta'],
    ['SLA_Normal_h', 168, 'Prazo-alvo em horas para urgência Normal (7 dias)']
  ]);

  var padrao = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);

  pastaFotos_();
  return 'Setup concluído.';
}

function criarAba_(ss, nome, cabecalho, linhas) {
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
      .setFontWeight('bold').setBackground('#c8102e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (linhas.length) sh.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
  }
  return sh;
}

function pastaFotos_() {
  var it = DriveApp.getFoldersByName(PASTA_FOTOS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA_FOTOS);
}

/**
 * Apaga a senha de um usuário para que ele defina outra no próximo acesso.
 * Uso: editar o CPF abaixo e executar no editor.
 */
function resetarSenha() {
  var cpf = 'INFORME_O_CPF';
  var sh = sh_(ABA.USU);
  var dados = sh.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (soDigitos_(dados[i][0]) === soDigitos_(cpf)) {
      sh.getRange(i + 1, 5, 1, 2).setValues([['', '']]);
      return 'Senha apagada para ' + dados[i][1];
    }
  }
  return 'CPF não encontrado.';
}

/* ===================== ROTEAMENTO ===================== */

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.acao === 'opcoes') return json_(opcoes_());
    if (p.acao === 'consultar') return json_(consultar_(p.protocolo));
    return json_({ ok: true, sistema: 'FriPremium Chamados' });
  } catch (err) {
    return json_({ ok: false, erro: 'Erro interno: ' + err.message });
  }
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, erro: 'Requisição inválida.' });
  }
  try {
    switch (req.acao) {
      case 'abrir': return json_(abrir_(req));
      case 'verificarCpf': return json_(verificarCpf_(req.cpf));
      case 'definirSenha': return json_(definirSenha_(req.cpf, req.senha));
      case 'login': return json_(login_(req.cpf, req.senha));
      case 'listar': return json_(listar_(sessao_(req.token)));
      case 'historico': return json_(historico_(sessao_(req.token), req.protocolo));
      case 'atualizar': return json_(atualizar_(sessao_(req.token), req));
      case 'foto': return json_(foto_(sessao_(req.token), req.protocolo));
      case 'sair': CacheService.getScriptCache().remove('s_' + req.token); return json_({ ok: true });
      default: return json_({ ok: false, erro: 'Ação desconhecida.' });
    }
  } catch (err) {
    if (err.sessao) return json_({ ok: false, sessaoExpirada: true, erro: 'Sessão expirada. Entre novamente.' });
    return json_({ ok: false, erro: err.message });
  }
}

/* ===================== PÚBLICO ===================== */

function opcoes_() {
  var destinos = linhasAtivas_(ABA.DEST).map(function (r) { return r[0]; });
  var cats = {};
  linhasAtivas_(ABA.CAT).forEach(function (r) {
    (cats[r[0]] = cats[r[0]] || []).push(r[1]);
  });
  return {
    ok: true,
    setores: linhasAtivas_(ABA.LOCAIS).map(function (r) { return r[0]; }),
    destinos: destinos,
    categorias: cats,
    urgencias: URGENCIAS,
    sla: sla_()
  };
}

function abrir_(r) {
  var nome = limpa_(r.nome, 80);
  var setor = limpa_(r.setor, 60);
  var local = limpa_(r.local, 60);
  var detalhe = limpa_(r.detalheLocal, 120);
  var destino = limpa_(r.destino, 60);
  var categoria = limpa_(r.categoria, 60);
  var urgencia = limpa_(r.urgencia, 20);
  var descricao = limpa_(r.descricao, 2000);

  if (!nome || nome.length < 3) throw new Error('Informe seu nome.');
  if (!local) throw new Error('Informe o local do problema.');
  if (!descricao || descricao.length < 10) throw new Error('Descreva o problema (mínimo 10 caracteres).');

  var op = opcoes_();
  if (op.setores.indexOf(local) < 0) throw new Error('Local inválido.');
  if (setor && op.setores.indexOf(setor) < 0) setor = '';
  if (op.destinos.indexOf(destino) < 0) throw new Error('Área responsável inválida.');
  if ((op.categorias[destino] || []).indexOf(categoria) < 0) throw new Error('Categoria inválida.');
  if (URGENCIAS.indexOf(urgencia) < 0) throw new Error('Urgência inválida.');

  var fotoId = '';
  if (r.foto && r.foto.b64) fotoId = salvarFoto_(r.foto);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var protocolo = proximoProtocolo_(destino);
    var agora = new Date();
    var linha = {
      Protocolo: protocolo, DataAbertura: agora, Solicitante: nome, SetorSolicitante: setor,
      Local: local, DetalheLocal: detalhe, AreaDestino: destino, Categoria: categoria,
      Urgencia: urgencia, Descricao: descricao, FotoId: fotoId, Status: 'Aberto',
      DataInicio: '', DataResolucao: '', Resolucao: '', UltimaAtualizacao: agora, AtualizadoPor: nome
    };
    sh_(ABA.CHAMADOS).appendRow(COLS.map(function (c) { return linha[c]; }));
    registrarHist_(protocolo, nome + ' (solicitante)', '', 'Aberto', 'Chamado aberto');
    if (fotoId) {
      try { DriveApp.getFileById(fotoId).setName(protocolo + '.jpg'); } catch (e) {}
    }
    return { ok: true, protocolo: protocolo };
  } finally {
    lock.releaseLock();
  }
}

function consultar_(protocolo) {
  var p = String(protocolo || '').trim().toUpperCase();
  if (!p) return { ok: false, erro: 'Informe o protocolo.' };
  var achado = buscar_(p);
  if (!achado) return { ok: false, erro: 'Protocolo não encontrado. Confira o número informado.' };
  var c = achado.obj;
  var hist = historicoBruto_(p).map(function (h) {
    return { dataHora: iso_(h.DataHora), status: h.StatusNovo };
  });
  return {
    ok: true,
    chamado: {
      protocolo: c.Protocolo, dataAbertura: iso_(c.DataAbertura), local: c.Local,
      detalheLocal: c.DetalheLocal, areaDestino: c.AreaDestino, categoria: c.Categoria,
      urgencia: c.Urgencia, descricao: c.Descricao, status: c.Status,
      dataResolucao: iso_(c.DataResolucao), resolucao: c.Resolucao
    },
    historico: hist
  };
}

function salvarFoto_(f) {
  var mime = String(f.mime || 'image/jpeg');
  if (!/^image\/(jpeg|png|webp)$/.test(mime)) throw new Error('Formato de foto não suportado.');
  var bytes = Utilities.base64Decode(String(f.b64));
  if (bytes.length > MAX_FOTO_BYTES) throw new Error('Foto muito grande.');
  var blob = Utilities.newBlob(bytes, mime, 'foto.jpg');
  return pastaFotos_().createFile(blob).getId(); // arquivo privado, herda permissões da pasta
}

/* ===================== AUTENTICAÇÃO ===================== */

function verificarCpf_(cpf) {
  var u = usuario_(cpf);
  if (!u) return { ok: false, erro: 'CPF sem acesso ao sistema.' };
  if (bloqueado_(u.cpf)) return { ok: false, erro: 'Acesso bloqueado por tentativas incorretas. Aguarde 15 minutos.' };
  return { ok: true, nome: u.nome, temSenha: !!u.hash };
}

function definirSenha_(cpf, senha) {
  var u = usuario_(cpf);
  if (!u) return { ok: false, erro: 'CPF sem acesso ao sistema.' };
  if (u.hash) return { ok: false, erro: 'Senha já definida. Use a tela de login.' };
  if (!/^\d{4}$/.test(String(senha || ''))) return { ok: false, erro: 'A senha deve ter 4 dígitos numéricos.' };
  var salt = Utilities.getUuid();
  sh_(ABA.USU).getRange(u.linha, 5, 1, 2).setValues([[hash_(senha, salt), salt]]);
  return { ok: true, token: criarSessao_(u), usuario: publico_(u) };
}

function login_(cpf, senha) {
  var u = usuario_(cpf);
  if (!u) return { ok: false, erro: 'CPF sem acesso ao sistema.' };
  if (bloqueado_(u.cpf)) return { ok: false, erro: 'Acesso bloqueado por tentativas incorretas. Aguarde 15 minutos.' };
  if (!u.hash) return { ok: false, primeiroAcesso: true, erro: 'Defina sua senha no primeiro acesso.' };
  if (hash_(String(senha || ''), u.salt) !== u.hash) {
    var rest = registrarFalha_(u.cpf);
    return { ok: false, erro: rest > 0 ? 'Senha incorreta. Restam ' + rest + ' tentativa(s).' : 'Acesso bloqueado por 15 minutos.' };
  }
  CacheService.getScriptCache().remove('f_' + u.cpf);
  return { ok: true, token: criarSessao_(u), usuario: publico_(u) };
}

function usuario_(cpf) {
  var alvo = soDigitos_(cpf);
  if (alvo.length !== 11) return null;
  var dados = sh_(ABA.USU).getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var r = dados[i];
    if (soDigitos_(r[0]).padStart(11, '0') === alvo && String(r[6]).toLowerCase() !== 'não' && String(r[6]).toLowerCase() !== 'nao') {
      return { linha: i + 1, cpf: alvo, nome: r[1], perfil: String(r[2]).toLowerCase(), area: r[3], hash: r[4], salt: r[5] };
    }
  }
  return null;
}

function publico_(u) {
  return { nome: u.nome, perfil: u.perfil, area: u.area, podeEditar: u.perfil === 'gestor' };
}

function criarSessao_(u) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(publico_(u)), SESSAO_SEG);
  return token;
}

function sessao_(token) {
  var raw = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (!raw) { var e = new Error('Sessão expirada.'); e.sessao = true; throw e; }
  return JSON.parse(raw);
}

function hash_(senha, salt) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + senha, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function bloqueado_(cpf) {
  return Number(CacheService.getScriptCache().get('f_' + cpf) || 0) >= MAX_TENTATIVAS;
}

function registrarFalha_(cpf) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get('f_' + cpf) || 0) + 1;
  c.put('f_' + cpf, String(n), BLOQUEIO_SEG);
  return MAX_TENTATIVAS - n;
}

/* ===================== ÁREA RESTRITA ===================== */

function listar_(s) {
  var dados = sh_(ABA.CHAMADOS).getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    var c = obj_(dados[i]);
    if (!c.Protocolo) continue;
    if (s.area !== '*' && c.AreaDestino !== s.area) continue;
    lista.push(saida_(c));
  }
  return { ok: true, usuario: s, chamados: lista, sla: sla_(), status: STATUS, urgencias: URGENCIAS };
}

function historico_(s, protocolo) {
  var achado = buscar_(String(protocolo || '').toUpperCase());
  if (!achado || (s.area !== '*' && achado.obj.AreaDestino !== s.area)) return { ok: false, erro: 'Chamado não encontrado.' };
  return {
    ok: true,
    historico: historicoBruto_(achado.obj.Protocolo).map(function (h) {
      return { dataHora: iso_(h.DataHora), usuario: h.Usuario, de: h.StatusAnterior, para: h.StatusNovo, comentario: h.Comentario };
    })
  };
}

function atualizar_(s, r) {
  if (!s.podeEditar) return { ok: false, erro: 'Seu perfil permite apenas visualização.' };
  var novo = String(r.status || '');
  var coment = limpa_(r.comentario, 2000);
  if (STATUS.indexOf(novo) < 0) return { ok: false, erro: 'Status inválido.' };
  if (novo === 'Resolvido' && coment.length < 5) return { ok: false, erro: 'Descreva o que foi feito para resolver.' };
  if (novo === 'Cancelado' && coment.length < 5) return { ok: false, erro: 'Informe o motivo do cancelamento.' };
  if (novo === 'Aguardando material/terceiro' && coment.length < 5) return { ok: false, erro: 'Informe o que está sendo aguardado.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var achado = buscar_(String(r.protocolo || '').toUpperCase());
    if (!achado || (s.area !== '*' && achado.obj.AreaDestino !== s.area)) return { ok: false, erro: 'Chamado não encontrado.' };
    var c = achado.obj;
    var anterior = c.Status;
    if (anterior === novo && !coment) return { ok: false, erro: 'Nada a atualizar.' };

    var agora = new Date();
    if (novo !== 'Aberto' && novo !== 'Cancelado' && !c.DataInicio) c.DataInicio = agora;
    if (novo === 'Resolvido' || novo === 'Cancelado') {
      c.DataResolucao = agora;
      c.Resolucao = coment;
    } else if (anterior === 'Resolvido' || anterior === 'Cancelado') {
      c.DataResolucao = '';   // reabertura
      c.Resolucao = '';
    }
    c.Status = novo;
    c.UltimaAtualizacao = agora;
    c.AtualizadoPor = s.nome;

    sh_(ABA.CHAMADOS).getRange(achado.linha, 1, 1, COLS.length).setValues([COLS.map(function (k) { return c[k]; })]);
    registrarHist_(c.Protocolo, s.nome, anterior, novo, coment);
    return { ok: true, chamado: saida_(c) };
  } finally {
    lock.releaseLock();
  }
}

function foto_(s, protocolo) {
  var achado = buscar_(String(protocolo || '').toUpperCase());
  if (!achado || !achado.obj.FotoId) return { ok: false, erro: 'Sem foto.' };
  if (s.area !== '*' && achado.obj.AreaDestino !== s.area) return { ok: false, erro: 'Sem acesso.' };
  var blob = DriveApp.getFileById(achado.obj.FotoId).getBlob();
  return { ok: true, mime: blob.getContentType(), b64: Utilities.base64Encode(blob.getBytes()) };
}

/* ===================== AUXILIARES ===================== */

function proximoProtocolo_(destino) {
  var prefixo = 'CH';
  linhasAtivas_(ABA.DEST).forEach(function (r) { if (r[0] === destino) prefixo = String(r[1] || 'CH').toUpperCase(); });
  var ano = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy');
  var base = prefixo + '-' + ano + '-';
  var sh = sh_(ABA.CHAMADOS);
  var max = 0;
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var p = String(r[0]);
      if (p.indexOf(base) === 0) max = Math.max(max, Number(p.slice(base.length)) || 0);
    });
  }
  return base + ('000' + (max + 1)).slice(-4);
}

function buscar_(protocolo) {
  var dados = sh_(ABA.CHAMADOS).getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toUpperCase() === protocolo) return { linha: i + 1, obj: obj_(dados[i]) };
  }
  return null;
}

function historicoBruto_(protocolo) {
  var dados = sh_(ABA.HIST).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toUpperCase() !== protocolo) continue;
    var h = {};
    COLS_HIST.forEach(function (k, j) { h[k] = dados[i][j]; });
    out.push(h);
  }
  return out;
}

function registrarHist_(protocolo, usuario, de, para, coment) {
  sh_(ABA.HIST).appendRow([protocolo, new Date(), usuario, de, para, coment || '']);
}

function saida_(c) {
  return {
    protocolo: c.Protocolo, dataAbertura: iso_(c.DataAbertura), solicitante: c.Solicitante,
    setorSolicitante: c.SetorSolicitante, local: c.Local, detalheLocal: c.DetalheLocal,
    areaDestino: c.AreaDestino, categoria: c.Categoria, urgencia: c.Urgencia,
    descricao: c.Descricao, temFoto: !!c.FotoId, status: c.Status, dataInicio: iso_(c.DataInicio),
    dataResolucao: iso_(c.DataResolucao), resolucao: c.Resolucao,
    ultimaAtualizacao: iso_(c.UltimaAtualizacao), atualizadoPor: c.AtualizadoPor
  };
}

function obj_(linha) {
  var o = {};
  COLS.forEach(function (k, j) { o[k] = linha[j]; });
  return o;
}

function sla_() {
  var out = { 'Crítica': 24, 'Alta': 72, 'Normal': 168 };
  var dados = sh_(ABA.CFG).getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    var v = Number(dados[i][1]);
    if (!v) continue;
    if (dados[i][0] === 'SLA_Critica_h') out['Crítica'] = v;
    if (dados[i][0] === 'SLA_Alta_h') out['Alta'] = v;
    if (dados[i][0] === 'SLA_Normal_h') out['Normal'] = v;
  }
  return out;
}

function linhasAtivas_(aba) {
  var dados = sh_(aba).getDataRange().getValues().slice(1);
  return dados.filter(function (r) {
    var ativo = String(r[r.length - 1]).toLowerCase();
    return r[0] !== '' && ativo !== 'não' && ativo !== 'nao';
  });
}

function sh_(nome) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh) throw new Error('Aba "' + nome + '" não encontrada. Execute setup().');
  return sh;
}

function limpa_(v, max) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // impede fórmula injetada na planilha
  return s;
}

function soDigitos_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }

function iso_(d) { return d instanceof Date ? d.toISOString() : (d ? String(d) : ''); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
