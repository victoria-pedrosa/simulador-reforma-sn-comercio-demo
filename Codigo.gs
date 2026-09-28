/**
 * Simulador da Reforma SN 2027 - Comercio
 * Escritório Contábil Exemplo | Time de IA
 *
 * Compara o DAS integral de hoje com o cenario hibrido de 2027 (PIS/COFINS fora
 * do DAS e CBS por fora, com credito nas compras - art. 41, §3o da LC 214/2025).
 * A leitura dos arquivos e o calculo acontecem no navegador (Index.html). Este
 * arquivo cuida da planilha, da chave da IA e do registro das consultas.
 */

/* ===================== configuracao ===================== */

/* Cole aqui a chave do Gemini. Ela fica no script, nao na planilha, para nao
   aparecer para quem so usa a ferramenta. */
var CHAVE_GEMINI = 'COLE_AQUI_SUA_CHAVE';

/* Modelo padrao. Se a chave nao acessar este, a chamada cai para um valido. */
var MODELO_PADRAO = 'gemini-3.7-flash';

/* Deixe em branco para usar a planilha em que o script esta. Preencha com o ID
   da planilha quando publicar como aplicativo da web, senao o link abre sem
   enxergar as abas. O ID e o trecho entre /d/ e /edit na URL da planilha. */
var PLANILHA_ID = '';

var ABA_EMPRESAS    = 'EMPRESAS';
var ABA_PARAMETROS  = 'PARAMETROS';
var ABA_SIMULACOES  = 'SIMULACOES';
var ABA_FORNECEDOR  = 'FORNECEDORES';
var ABA_PRODUTOS    = 'PRODUTOS_SIM';
var ABA_LOG         = 'LOG_EXECUCAO';

var GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta';

/* ===================== pagina e menu ===================== */

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Reforma SN 2027')
      .addItem('Abrir o simulador', 'mostrarLateral')
      .addSeparator()
      .addItem('Criar as abas de apoio', 'prepararAbas')
      .addItem('Conferir a chave da IA', 'conferirChave')
      .addToUi();
  } catch (e) {}
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Reforma SN 2027 - Comercio')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function mostrarLateral() {
  var h = HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Reforma SN 2027 - Comercio').setWidth(1400);
  SpreadsheetApp.getUi().showModalDialog(h, 'Reforma SN 2027 - Comercio');
}

function conferirChave() {
  var r = temChaveIA();
  SpreadsheetApp.getUi().alert(r.tem
    ? 'Chave preenchida. Modelo padrao: ' + r.modelo
    : 'A constante CHAVE_GEMINI ainda esta com o texto de exemplo. Abra o Codigo.gs e cole a chave.');
}

/* ===================== planilha ===================== */

function ss_() {
  var ss = null;
  if (PLANILHA_ID && PLANILHA_ID.indexOf('COLE_AQUI') < 0) {
    try { ss = SpreadsheetApp.openById(PLANILHA_ID); } catch (e) { ss = null; }
  }
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abra pela planilha, ou preencha PLANILHA_ID no Codigo.gs para usar pelo link.');
  return ss;
}

/** A leitura tolera nao ter planilha: a ferramenta continua calculando. */
function ssOpcional_() {
  try { return ss_(); } catch (e) { return null; }
}

function aba_(nome, cabecalho) {
  var ss = ss_(), aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    if (cabecalho) {
      aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
         .setFontWeight('bold').setBackground('#14432D').setFontColor('#ffffff');
      aba.setFrozenRows(1);
      aba.setColumnWidths(1, cabecalho.length, 150);
    }
  }
  return aba;
}

function prepararAbas() {
  try {
    var e = aba_(ABA_EMPRESAS, ['Codigo', 'CNPJ', 'Razao social', 'Anexo',
                                'Responsavel', 'Observacao']);
    if (e.getLastRow() < 2) {
      e.getRange('A2:F2').setValues([['', '', '', 'I', '', 'preencha uma linha por cliente']]);
      e.getRange('D2:D').setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(['I', 'II'], true).build());
    }

    var p = aba_(ABA_PARAMETROS, ['Parametro', 'Valor', 'O que e']);
    if (p.getLastRow() < 2) {
      p.getRange(2, 1, 5, 3).setValues([
        ['CBS 2027 (%)', 8.8, 'aliquota de referencia menos 0,1 p.p. (art. 347)'],
        ['IBS teste (%)', 0.1, 'aliquota de teste de 2027 (art. 344)'],
        ['Somar IBS ao hibrido (S/N)', 'N', 'o IBS de 2027 e simbolico; por padrao fica fora'],
        ['Considerar reducao que depende da Anvisa (S/N)', 'S', 'Anexo IV, XII, XIV e art. 133'],
        ['Considerar reducao nos produtos a conferir (S/N)', 'N', 'aderencia baixa nao sustenta o enquadramento']
      ]);
    }

    aba_(ABA_SIMULACOES, ['ID', 'Data/Hora', 'Usuario', 'Empresa', 'CNPJ', 'Periodo',
      'Faturamento', 'RBT12', 'Anexo', 'Faixa', 'Aliquota efetiva',
      'DAS hoje', 'PIS+COFINS no DAS', 'DAS hibrido',
      'Compras', '% fornecedor regime normal', 'Reducao media saidas', 'Reducao media compras',
      'CBS debito', 'CBS credito', 'CBS a recolher', 'Credito acumulado', 'IBS teste',
      'Total 2027', 'Diferenca mes', 'Diferenca ano', 'Variacao p.p.', 'Recomendacao',
      'Notas lidas', 'NFC-e', 'NF-e', 'Canceladas', 'Devolucoes', 'Fornecedores', 'Produtos',
      'Interpretacao', 'Origem da interpretacao']);

    aba_(ABA_FORNECEDOR, ['ID da simulacao', 'CNPJ', 'Fornecedor', 'Regime', 'Notas', 'Valor']);

    aba_(ABA_PRODUTOS, ['ID da simulacao', 'NCM', 'Produto', 'Saidas', 'Entradas',
      'Anexo', 'Item', 'Aderencia', 'Reducao do anexo', 'Reducao aplicada', 'Situacao', 'Base legal']);

    aba_(ABA_LOG, ['Data/Hora', 'Usuario', 'Empresa', 'Periodo', 'Notas', 'Produtos',
                   'IA', 'Modelo', 'Resultado']);
    return { ok: true };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/** Lista as empresas cadastradas para o seletor da ferramenta. */
function lerEmpresas() {
  try {
    var ss = ssOpcional_();
    if (!ss) return { ok: true, empresas: [] };
    var aba = ss.getSheetByName(ABA_EMPRESAS);
    if (!aba || aba.getLastRow() < 2) return { ok: true, empresas: [] };
    var v = aba.getRange(2, 1, aba.getLastRow() - 1, 6).getValues();
    var out = [];
    for (var i = 0; i < v.length; i++) {
      var cnpj = String(v[i][1] || '').replace(/\D+/g, '');
      var nome = String(v[i][2] || '').trim();
      if (!nome && !cnpj) continue;
      out.push({ codigo: String(v[i][0] || '').trim(), cnpj: cnpj, nome: nome,
                 anexo: String(v[i][3] || 'I').trim() || 'I',
                 responsavel: String(v[i][4] || '').trim() });
    }
    return { ok: true, empresas: out };
  } catch (e) { return { ok: false, erro: String(e), empresas: [] }; }
}

/** Parametros globais; o que estiver em branco fica com o padrao da tela. */
function lerParametros() {
  try {
    var ss = ssOpcional_();
    if (!ss) return { ok: true, params: {} };
    var aba = ss.getSheetByName(ABA_PARAMETROS);
    if (!aba || aba.getLastRow() < 2) return { ok: true, params: {} };
    var v = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getValues(), p = {};
    for (var i = 0; i < v.length; i++) {
      var k = String(v[i][0] || '').toUpperCase();
      if (k.indexOf('CBS') === 0) p.cbs = Number(v[i][1]);
      else if (k.indexOf('IBS TESTE') === 0) p.ibs = Number(v[i][1]);
      else if (k.indexOf('SOMAR IBS') === 0) p.incIbs = /^S/i.test(String(v[i][1]));
      else if (k.indexOf('CONSIDERAR REDUCAO QUE DEPENDE') === 0) p.anvisa = /^S/i.test(String(v[i][1]));
      else if (k.indexOf('CONSIDERAR REDUCAO NOS PRODUTOS') === 0) p.conferir = /^S/i.test(String(v[i][1]));
    }
    return { ok: true, params: p };
  } catch (e) { return { ok: false, erro: String(e), params: {} }; }
}

/**
 * Grava a consulta. Uma linha em SIMULACOES e as tabelas filhas ligadas pelo ID,
 * para dar para voltar depois e reproduzir o numero que foi entregue ao cliente.
 */
function salvarSimulacao(d) {
  try {
    prepararAbas();
    var id = 'SIM' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var usuario = '';
    try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}

    aba_(ABA_SIMULACOES).appendRow([id, new Date(), usuario, d.empresa || '', d.cnpj || '',
      d.periodo || '', n_(d.fat), n_(d.rbt), d.anexo || '', d.faixa || '', n_(d.aliq),
      n_(d.das), n_(d.pisCof), n_(d.dasHib),
      n_(d.compras), n_(d.regular), n_(d.redSai), n_(d.redEnt),
      n_(d.cbsDeb), n_(d.cbsCred), n_(d.cbsLiq), n_(d.saldoCredor), n_(d.ibs),
      n_(d.hib), n_(d.dif), n_(d.difAno), n_(d.pp), d.recomendacao || '',
      n_(d.notas), n_(d.nfce), n_(d.nfe), n_(d.canceladas), n_(d.devolucoes),
      n_(d.qtdFornecedores), n_(d.qtdProdutos),
      String(d.interpretacao || ''), d.origemInterpretacao || 'regras']);

    if (d.fornecedores && d.fornecedores.length) {
      var lf = d.fornecedores.map(function (f) {
        return [id, f.cnpj || '', f.nome || '', f.regime || '', n_(f.notas), n_(f.valor)];
      });
      aba_(ABA_FORNECEDOR).getRange(aba_(ABA_FORNECEDOR).getLastRow() + 1, 1, lf.length, 6).setValues(lf);
    }
    if (d.produtos && d.produtos.length) {
      var lp = d.produtos.slice(0, 3000).map(function (p) {
        return [id, p.ncm || '', p.desc || '', n_(p.saidas), n_(p.entradas),
                p.anexo || '', p.item || '', n_(p.score), n_(p.red), n_(p.redAplicada),
                p.situacao || '', p.hipotese || ''];
      });
      aba_(ABA_PRODUTOS).getRange(aba_(ABA_PRODUTOS).getLastRow() + 1, 1, lp.length, 12).setValues(lp);
    }
    registrarLog({ empresa: d.empresa, periodo: d.periodo, notas: d.notas,
                   produtos: d.qtdProdutos, usouIA: d.origemInterpretacao === 'IA',
                   modelo: d.modelo, resultado: 'Simulacao gravada ' + id });
    return { ok: true, id: id };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

function n_(v) { var x = Number(v); return isFinite(x) ? x : 0; }

function registrarLog(d) {
  try {
    var usuario = '';
    try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
    aba_(ABA_LOG).appendRow([new Date(), usuario, d.empresa || '', d.periodo || '',
      n_(d.notas), n_(d.produtos), d.usouIA ? 'Sim' : 'Nao', d.modelo || '', d.resultado || '']);
    return { ok: true };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

/* ===================== interpretacao pela IA ===================== */

function temChaveIA() {
  return { ok: true, tem: !!chaveValida_(), modelo: MODELO_PADRAO };
}

function chaveValida_() {
  var k = String(CHAVE_GEMINI || '').trim();
  return (k && k.indexOf('COLE_AQUI') < 0 && k.length >= 20) ? k : '';
}

function chave_() {
  var k = chaveValida_();
  if (!k) throw new Error('Preencha a constante CHAVE_GEMINI no arquivo Codigo.gs.');
  return k;
}

function listarModelosIA() {
  try {
    var r = UrlFetchApp.fetch(GEMINI_BASE + '/models?key=' + encodeURIComponent(chave_()),
                              { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200)
      throw new Error('HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
    var ms = (JSON.parse(r.getContentText()).models || [])
      .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
      .map(function (m) { return m.name.replace('models/', ''); })
      .filter(function (n) { return n.indexOf('gemini') === 0; });
    return { ok: true, modelos: ms };
  } catch (e) { return { ok: false, erro: String(e) }; }
}

function modeloAlternativo_() {
  var r = listarModelosIA();
  if (!r.ok || !r.modelos.length) return '';
  var flash = r.modelos.filter(function (m) { return m.indexOf('flash') > 0 && m.indexOf('lite') < 0; });
  return (flash.length ? flash : r.modelos)[0];
}

/**
 * A IA escreve a leitura do mes - nunca os numeros. Ela recebe os valores ja
 * calculados e e proibida de criar, recalcular ou arredondar qualquer um deles.
 * Se a chamada falhar, a tela mantem a interpretacao por regras.
 */
var _jaTrocouModelo = false;
function interpretarIA(fatos, modelo) {
  try {
    modelo = modelo || MODELO_PADRAO;

    var regras =
      'Voce escreve a leitura de uma simulacao tributaria para a equipe fiscal da Escritório Contábil Exemplo.\n' +
      'REGRAS:\n' +
      '1. Use SOMENTE os numeros do JSON. Nunca calcule, estime, arredonde ou invente valor nenhum.\n' +
      '2. Escreva os valores exatamente como vieram no campo "formatado" quando existir.\n' +
      '3. Portugues do Brasil, tom direto, sem jargao tecnico e sem introducao.\n' +
      '4. Cada paragrafo comeca com uma afirmacao curta e termina explicando o porque.\n' +
      '5. Nao recomende a opcao pelo regime hibrido como decisao fechada: a analise e de um mes.\n' +
      '6. Nao cite artigo de lei que nao esteja no JSON.\n' +
      '7. De 3 a 5 paragrafos, cada um com no maximo 3 frases.\n' +
      '8. As ressalvas sao frases curtas, no maximo 6.\n' +
      'O QUE EXPLICAR, nesta ordem: o que acontece com a CBS nas vendas; o que o credito das ' +
      'compras faz; a conclusao do mes com o valor no ano; concentracao de fornecedor, se houver; ' +
      'o que mudaria o resultado.';

    var corpo = {
      systemInstruction: { parts: [{ text: regras }] },
      contents: [{ role: 'user', parts: [{ text: JSON.stringify(fatos) }] }],
      generationConfig: {
        temperature: 0,
        responseMimeType: 'application/json',
        responseSchema: {
          type: 'object',
          properties: {
            paragrafos: { type: 'array', items: { type: 'string' } },
            ressalvas: { type: 'array', items: { type: 'string' } }
          },
          required: ['paragrafos', 'ressalvas']
        }
      }
    };

    var r = UrlFetchApp.fetch(GEMINI_BASE + '/models/' + modelo + ':generateContent?key=' +
                              encodeURIComponent(chave_()), {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(corpo), muteHttpExceptions: true
    });
    var cod = r.getResponseCode();
    if (cod === 429) return { ok: false, erro: 'Limite de requisicoes da API atingido. Tente de novo em instantes.' };
    if (cod === 404 && !_jaTrocouModelo) {
      var alt = modeloAlternativo_();
      if (alt && alt !== modelo) {
        _jaTrocouModelo = true;
        var novo = interpretarIA(fatos, alt);
        if (novo.ok) novo.modeloUsado = alt;
        return novo;
      }
    }
    if (cod !== 200) throw new Error('HTTP ' + cod + ' ' + r.getContentText().slice(0, 300));

    var j = JSON.parse(r.getContentText());
    var txt = j.candidates && j.candidates[0] && j.candidates[0].content &&
              j.candidates[0].content.parts && j.candidates[0].content.parts[0].text;
    if (!txt) throw new Error('Resposta vazia do modelo.');
    var o = JSON.parse(txt);
    return { ok: true, paragrafos: o.paragrafos || [], ressalvas: o.ressalvas || [],
             modeloUsado: modelo };
  } catch (e) { return { ok: false, erro: String(e) }; }
}