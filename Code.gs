/**
 * New Blue Card Shop · Scanner ao vivo (servidor)
 * ------------------------------------------------
 * Cole este arquivo inteiro em Extensões › Apps Script da planilha
 * "New Blue Card Shop — Estoque". Ele:
 *   1. recebe as fotos que a página do scanner ao vivo envia do celular;
 *   2. identifica a carta (Pokémon TCG ou Disney Lorcana) com o Gemini;
 *   3. confere nome, coleção e número no banco de cartas oficial
 *      (Pokémon TCG API e Lorcast) e pega imagem e preço de referência;
 *   4. grava a carta em Produtos (se for nova) e a entrada em Movimentações.
 *
 * Nada aqui roda sozinho: só responde à página do scanner que tiver o
 * código de acesso que o menu Scanner › Configurar cria.
 */

var NB = {
  VERSION: '1.0',
  PROD: 'Produtos',
  MOV: 'Movimentações',
  CONFIG: 'Config',
  FIRST_ROW: 5,
  MOV_TYPE: 'Entrada da coleção',
  MOV_ORIGIN: 'Coleção própria',
  PAGE_URL: 'https://gustapaesleme.github.io/nbcs-scanner/',
  // modelos do Gemini, do mais rápido para o mais forte; o primeiro que existir na sua conta é usado
  MODELS: ['gemini-3.5-flash-lite', 'gemini-3.8-flash', 'gemini-3.6-flash', 'gemini-2.5-flash-lite', 'gemini-2.5-flash'],
  CAREFUL_MODELS: ['gemini-3.8-flash', 'gemini-3.6-flash', 'gemini-2.5-flash'],
};

// Produtos: colunas de entrada (as outras são fórmulas no cabeçalho e não podem receber valores)
// A SKU | B Jogo | C Tipo | D Produto | E Coleção / versão | F Nº da carta | G Raridade | H Idioma | I Condição
// P Estoque mín. | Q Preço de venda R$ | R Preço ref. R$ | S Custo previsto US$ | Y No site? | Z Link | AA Observações
// Movimentações: A Data | B Tipo | C SKU | E Qtd | F Recebido? | H Origem / loja | I Custo unit. US$
// K Dólar da compra | L Custo unit. R$ | M Frete e taxas R$ | O Pedido / nota | P Observações

/* ============================== menu ============================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scanner')
    .addItem('1. Configurar (chave do Gemini)', 'nbSetup')
    .addItem('2. Link para o celular', 'nbShowLink')
    .addSeparator()
    .addItem('Testar a chave do Gemini', 'nbTestKey')
    .addItem('Trocar o código de acesso', 'nbNewToken')
    .addItem('Endereço da página do scanner', 'nbPageUrl')
    .addItem('Trocar o URL do app da Web', 'nbChangeUrl')
    .addToUi();
}

function nbPageUrl() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty('PAGE_URL') || NB.PAGE_URL;
  var r = ui.prompt('Endereço da página do scanner', 'Onde você publicou o index.html do scanner ao vivo.\nAtual: ' + cur + '\nDeixe em branco para manter.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var v = r.getResponseText().trim();
  if (!v) return;
  if (!/^https:\/\//.test(v)) { ui.alert('O endereço precisa começar com https:// (a câmera só liga em páginas seguras).'); return; }
  props.setProperty('PAGE_URL', v.replace(/#.*$/, ''));
  ui.alert('Pronto. Gere o link de novo em Scanner › 2. Link para o celular.');
}

function nbSetup() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var has = !!props.getProperty('GEMINI_KEY');
  var r = ui.prompt('Chave do Gemini',
    (has ? 'Já existe uma chave salva. Cole uma nova para trocar, ou clique em Cancelar.\n\n' : '') +
    'Crie a chave em aistudio.google.com/apikey (entre com o mesmo Gmail) e cole aqui:',
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim()) {
    props.setProperty('GEMINI_KEY', r.getResponseText().trim());
    props.deleteProperty('GEMINI_MODEL_OK');
  }
  if (!props.getProperty('SCANNER_TOKEN')) props.setProperty('SCANNER_TOKEN', newToken_());
  var test = props.getProperty('GEMINI_KEY') ? testKey_() : 'sem chave';
  ui.alert('Scanner configurado',
    'Chave do Gemini: ' + test + '\n' +
    'Código de acesso do celular: ' + props.getProperty('SCANNER_TOKEN') + '\n\n' +
    'Próximo passo: no Apps Script, Implantar › Nova implantação › App da Web,\n' +
    'Executar como: Eu · Quem pode acessar: Qualquer pessoa. Copie o URL do app da Web\n' +
    'e use o menu Scanner › 2. Link para o celular.',
    ui.ButtonSet.OK);
}

function nbNewToken() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('Trocar o código de acesso?', 'O celular vai pedir o código novo (ou o link novo) para continuar escaneando.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().setProperty('SCANNER_TOKEN', newToken_());
  nbShowLink();
}

function nbTestKey() {
  SpreadsheetApp.getUi().alert('Chave do Gemini', testKey_(), SpreadsheetApp.getUi().ButtonSet.OK);
}

function nbChangeUrl() {
  PropertiesService.getScriptProperties().deleteProperty('WEBAPP_URL');
  nbShowLink(true);
}

function nbShowLink(askUrl) {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var EXEC = /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/;
  if (!props.getProperty('SCANNER_TOKEN')) props.setProperty('SCANNER_TOKEN', newToken_());
  var api = props.getProperty('WEBAPP_URL') || '';
  if (!api && askUrl !== true) {
    // after "Implantar › App da Web" the script knows its own /exec address
    try { var u = ScriptApp.getService().getUrl(); if (EXEC.test(u)) { api = u; props.setProperty('WEBAPP_URL', api); } } catch (e) {}
  }
  if (!api) {
    var r = ui.prompt('URL do app da Web', 'Cole o URL que termina em /exec (Apps Script › Implantar › Gerenciar implantações).', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    var typed = r.getResponseText().trim();
    if (!EXEC.test(typed)) {
      ui.alert('Esse URL não parece o do app da Web. Ele começa com https://script.google.com/macros/s/ e termina em /exec.');
      return;
    }
    api = typed; props.setProperty('WEBAPP_URL', api);
  }
  if (!api) { ui.alert('Falta o URL do app da Web.'); return; }
  var page = props.getProperty('PAGE_URL') || NB.PAGE_URL;
  var link = page + '#api=' + encodeURIComponent(api) + '&token=' + encodeURIComponent(props.getProperty('SCANNER_TOKEN'));
  var html = HtmlService.createHtmlOutput(
    '<div style="font:14px/1.45 system-ui,sans-serif;color:#0B1D24">' +
    '<p style="margin:0 0 10px">Aponte a câmera do celular para o código ou envie o link para você mesmo.</p>' +
    '<div id="qr" style="display:grid;place-items:center;min-height:260px"></div>' +
    '<textarea readonly style="width:100%;height:76px;font:12px monospace" onclick="this.select()">' + escHtml_(link) + '</textarea>' +
    '<p style="color:#4D6A74;font-size:12px;margin:8px 0 0">Página do scanner: ' + escHtml_(page) + '<br>Quem tiver este link consegue lançar cartas na planilha: não publique.</p></div>' +
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcode-generator/1.4.4/qrcode.min.js"></script>' +
    '<script>try{var q=qrcode(0,"M");q.addData(' + JSON.stringify(link) + ');q.make();document.getElementById("qr").innerHTML=q.createSvgTag({cellSize:4,margin:2});}catch(e){document.getElementById("qr").textContent="Copie o link abaixo.";}</script>'
  ).setWidth(420).setHeight(470);
  ui.showModalDialog(html, 'Scanner ao vivo no celular');
}

/* ============================== web app ============================== */

function doGet() {
  return json_({ ok: true, app: 'nbcs-scanner', version: NB.VERSION, hint: 'Use a página do scanner ao vivo.' });
}

function doPost(e) {
  var req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, code: 'bad_json', message: 'Pedido inválido.' }); }
  var token = PropertiesService.getScriptProperties().getProperty('SCANNER_TOKEN');
  if (!token || String(req.token || '') !== token) return json_({ ok: false, code: 'bad_token', message: 'Código de acesso inválido. Gere o link de novo no menu Scanner da planilha.' });
  try {
    switch (req.action) {
      case 'ping': return json_(ping_());
      case 'scan': return json_(scan_(req));
      case 'log': return json_(logCard_(req));
      case 'undo': return json_(undo_(req));
      default: return json_({ ok: false, code: 'bad_action', message: 'Ação desconhecida.' });
    }
  } catch (err) {
    return json_({ ok: false, code: err.code || 'server_error', message: String(err.message || err) });
  }
}

function ping_() {
  var ss = SpreadsheetApp.getActive();
  var p = ss.getSheetByName(NB.PROD), m = ss.getSheetByName(NB.MOV);
  if (!p || !m) throw fail_('no_tabs', 'A planilha não tem as abas Produtos e Movimentações.');
  return {
    ok: true, version: NB.VERSION, sheet: ss.getName(),
    products: lastDataRow_(p) - NB.FIRST_ROW + 1,
    hasKey: !!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY'),
    model: PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL_OK') || '',
  };
}

/* ============================== scan ============================== */

function scan_(req) {
  var id = cleanId_(req.id);
  var cache = CacheService.getScriptCache();
  var prev = cache.get('res:' + id);
  if (prev) return JSON.parse(prev);                                   // the phone retried: same answer, nothing written twice

  var images = (req.images || []).filter(function (b) { return typeof b === 'string' && b.length > 100; }).slice(0, 3);
  if (!images.length) throw fail_('no_image', 'A foto não chegou.');
  var prefs = req.prefs || {};

  var t0 = Date.now();
  var ai = geminiIdentify_(images, !!req.careful);
  var t1 = Date.now();
  var r = ai.data || {};
  if (r.isCard === false) {
    var res0 = { ok: true, isCard: false, check: str_(r.check) || 'Não vi uma carta de Pokémon ou Lorcana.', model: ai.model, ms: { ai: t1 - t0 } };
    cache.put('res:' + id, JSON.stringify(res0), 21600);
    return res0;
  }
  var card = normalizeAi_(r, prefs);
  var look = { verified: false, candidates: [] };
  try { look = card.game === 'Lorcana' ? lookupLorcana_(card) : lookupPokemon_(card); } catch (err) { look = { verified: false, candidates: [], error: String(err.message || err) }; }
  if (look.pick) applyPick_(card, look.pick);
  card.verified = !!look.verified;
  card.sku = buildSku_(card);
  card.obs = obsFor_(card);
  var t2 = Date.now();

  var products = readProducts_();
  var match = findMatch_(card, products);
  var avoid = String(prefs.avoidSku || '');
  var sku = match ? match.sku : card.sku;
  var confident = card.verified || card.confidence === 'alta';
  var dup = avoid && avoid === sku;
  var res = {
    ok: true, isCard: true, card: card, match: match, sku: sku,
    candidates: look.candidates.slice(0, 4), verified: card.verified,
    model: ai.model, logged: null, needsReview: false, reason: '',
  };
  if (prefs.autoLog !== false && confident && card.name && !dup) {
    res.logged = writeRows_(card, id, prefs, match);
  } else {
    res.needsReview = true;
    res.reason = dup ? 'duplicate' : (!card.name ? 'no_name' : (!confident ? 'low_confidence' : 'manual'));
  }
  res.ms = { ai: t1 - t0, lookup: t2 - t1, total: Date.now() - t0 };
  cache.put('res:' + id, JSON.stringify(res), 21600);
  return res;
}

function logCard_(req) {
  var id = cleanId_(req.id);
  var c = req.card || {};
  if (!str_(c.name)) throw fail_('no_name', 'Falta o nome da carta.');
  var card = {
    game: c.game === 'Lorcana' ? 'Lorcana' : 'Pokémon',
    kind: oneOf_(c.kind, ['Carta avulsa', 'Promo', 'Slab (graduada)'], 'Carta avulsa'),
    name: str_(c.name), subtitle: str_(c.subtitle), setName: str_(c.setName), setCode: str_(c.setCode),
    number: str_(c.number), rarity: str_(c.rarity), finish: oneOf_(c.finish, ['Normal', 'Holo', 'Reverse Holo', 'Foil', 'Full Art'], 'Normal'),
    language: (str_(c.language) || 'PT').toUpperCase().slice(0, 2), condition: str_(c.condition) || 'NM',
    grading: c.grading && typeof c.grading === 'object' ? c.grading : {}, image: str_(c.image), obs: str_(c.obs),
    priceUsd: num_(c.priceUsd), stats: str_(c.stats), illustrator: str_(c.illustrator),
  };
  card.sku = str_(c.sku) || buildSku_(card);
  if (!card.obs) card.obs = obsFor_(card);
  var products = readProducts_();
  var match = req.useSku ? { sku: String(req.useSku), label: '' } : findMatch_(card, products);
  var logged = writeRows_(card, id, req.prefs || {}, match);
  var res = { ok: true, isCard: true, card: card, match: match, sku: logged.sku, logged: logged, needsReview: false };
  CacheService.getScriptCache().put('res:' + id, JSON.stringify(res), 21600);
  return res;
}

function undo_(req) {
  var id = cleanId_(req.id);
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ss = SpreadsheetApp.getActive();
    var m = ss.getSheetByName(NB.MOV), p = ss.getSheetByName(NB.PROD);
    var mLast = lastDataRow_(m);
    if (mLast < NB.FIRST_ROW) return { ok: true, removed: null };
    var n = mLast - NB.FIRST_ROW + 1;
    var skus = m.getRange(NB.FIRST_ROW, 3, n, 1).getValues();
    var notes = m.getRange(NB.FIRST_ROW, 16, n, 1).getValues();
    var row = -1, sku = '';
    for (var i = n - 1; i >= 0; i--) if (String(notes[i][0]).indexOf('scan:' + id) >= 0) { row = NB.FIRST_ROW + i; sku = String(skus[i][0]); break; }
    if (row < 0) return { ok: true, removed: null };
    m.deleteRow(row);
    var removedProd = null;
    var made = CacheService.getScriptCache().get('made:' + id);
    if (made && made === sku) {
      var others = skus.some(function (r, j) { return NB.FIRST_ROW + j !== row && String(r[0]) === sku; });
      if (!others) {
        var pLast = lastDataRow_(p);
        if (pLast >= NB.FIRST_ROW) {
          var ps = p.getRange(NB.FIRST_ROW, 1, pLast - NB.FIRST_ROW + 1, 1).getValues();
          for (var k = ps.length - 1; k >= 0; k--) if (String(ps[k][0]) === sku) { p.deleteRow(NB.FIRST_ROW + k); removedProd = NB.FIRST_ROW + k; break; }
        }
      }
      CacheService.getScriptCache().remove('made:' + id);
    }
    CacheService.getScriptCache().remove('res:' + id);
    SpreadsheetApp.flush();
    return { ok: true, removed: { movRow: row, prodRow: removedProd, sku: sku } };
  } finally { lock.releaseLock(); }
}

/* ============================== sheet writes ============================== */

function writeRows_(card, id, prefs, match) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var ss = SpreadsheetApp.getActive();
    var p = ss.getSheetByName(NB.PROD), m = ss.getSheetByName(NB.MOV);
    if (!p || !m) throw fail_('no_tabs', 'A planilha não tem as abas Produtos e Movimentações.');

    // already written for this scan? (a retry after a timeout): report the row, write nothing
    var mLast = lastDataRow_(m);
    if (mLast >= NB.FIRST_ROW) {
      var notes = m.getRange(NB.FIRST_ROW, 16, mLast - NB.FIRST_ROW + 1, 1).getValues();
      for (var i = notes.length - 1; i >= 0; i--) {
        if (String(notes[i][0]).indexOf('scan:' + id) >= 0) return { movRow: NB.FIRST_ROW + i, prodRow: null, sku: String(m.getRange(NB.FIRST_ROW + i, 3).getValue()), already: true };
      }
    }

    var pLast = lastDataRow_(p);
    var skus = pLast >= NB.FIRST_ROW ? p.getRange(NB.FIRST_ROW, 1, pLast - NB.FIRST_ROW + 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
    var sku = match && match.sku ? match.sku : card.sku;
    var prodRow = null;
    if (skus.indexOf(sku) < 0) {
      prodRow = pLast + 1;
      var colecao = [card.setName, card.setCode ? '(' + card.setCode + ')' : ''].filter(Boolean).join(' ');
      var raridade = [card.rarity, card.finish && card.finish !== 'Normal' ? card.finish : ''].filter(Boolean).join(' · ');
      var g = card.grading || {};
      var cond = card.kind === 'Slab (graduada)' && g.grade ? ((g.company || '') + ' ' + g.grade).trim() : (card.condition || 'NM');
      p.getRange(prodRow, 1, 1, 9).setValues([[
        txt_(sku), card.game, card.kind, txt_(productName_(card)), txt_(colecao),
        card.number ? "'" + card.number : '', txt_(raridade), card.language || '', txt_(cond),
      ]]);
      p.getRange(prodRow, 16, 1, 4).setValues([['', numOrBlank_(prefs.price), '', '']]);       // P..S
      p.getRange(prodRow, 25, 1, 3).setValues([[
        prefs.site === 'Não' ? 'Não' : 'Sim', card.image || '',
        txt_([card.obs, 'Scanner ao vivo ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy')].filter(Boolean).join(' · ')),
      ]]);                                                                                      // Y..AA
      CacheService.getScriptCache().put('made:' + id, sku, 21600);
    }

    var movRow = mLast < NB.FIRST_ROW ? NB.FIRST_ROW : mLast + 1;
    var qty = Math.max(1, parseInt(prefs.qty, 10) || 1);
    var cost = num_(prefs.cost);
    m.getRange(movRow, 1, 1, 3).setValues([[todaySerial_(ss), NB.MOV_TYPE, txt_(sku)]]);       // A..C
    m.getRange(movRow, 5, 1, 2).setValues([[qty, 'Sim']]);                                     // E..F
    m.getRange(movRow, 8, 1, 1).setValues([[NB.MOV_ORIGIN]]);                                  // H
    if (cost !== null) m.getRange(movRow, 12, 1, 1).setValues([[cost]]);                       // L
    m.getRange(movRow, 16, 1, 1).setValues([['Scanner ao vivo · scan:' + id]]);                // P
    SpreadsheetApp.flush();
    return { movRow: movRow, prodRow: prodRow, sku: sku };
  } finally { lock.releaseLock(); }
}

function readProducts_() {
  var p = SpreadsheetApp.getActive().getSheetByName(NB.PROD);
  if (!p) throw fail_('no_tabs', 'A planilha não tem a aba Produtos.');
  var last = lastDataRow_(p);
  return last >= NB.FIRST_ROW ? p.getRange(NB.FIRST_ROW, 1, last - NB.FIRST_ROW + 1, 9).getDisplayValues() : [];
}

function lastDataRow_(sh) {
  var max = sh.getLastRow();
  if (max < NB.FIRST_ROW) return NB.FIRST_ROW - 1;
  var vals = sh.getRange(NB.FIRST_ROW, 1, max - NB.FIRST_ROW + 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim() !== '') return NB.FIRST_ROW + i;
  return NB.FIRST_ROW - 1;
}

function todaySerial_(ss) {
  var d = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd').split('-');
  return Math.round((Date.UTC(+d[0], +d[1] - 1, +d[2]) - Date.UTC(1899, 11, 30)) / 86400000);
}

/* ============================== Gemini ============================== */

var PROMPT = [
  'You identify ONE trading card for a Brazilian card shop\'s stock sheet.',
  'Image 1 is the card as framed by the phone camera. Image 2 zooms into the bottom of the same card: card number, set code, language, rarity symbol, illustrator.',
  'Supported games: Pokémon TCG and Disney Lorcana.',
  'Rules:',
  '- Report only what is printed or clearly visible. Use "" when unsure.',
  '- name: exactly as printed, in the printed language. nameEn: the official English card name if you know it (same as name for English cards).',
  '- number: exactly as printed with the total, e.g. "025/198", "TG05/TG30", "185/167". For black-star promos the promo number, e.g. "SVP 085".',
  '- Pokémon setCode: the letters printed next to the number on modern cards (e.g. SVI, PAL, OBF, MEW, PAR, PAF, TEF, TWM, SFA, SCR, SSP, PRE, JTG, DRI, BLK, WHT, MEG). Older cards without printed code: "".',
  '- Lorcana: the bottom line reads like "123/204 · EN · 5": number "123/204", language "EN", setCode "5" (the set number).',
  '- condition: judge only visible wear. For a graded slab, condition is the label grade, e.g. "PSA 10", and fill grading.',
  '- If the image does not show a single Pokémon or Lorcana card (empty table, hand, other game), set isCard to false and say what you see in "check".',
  'Reply with only this JSON object:',
  '{"isCard":true,"game":"Pokémon","kind":"Carta avulsa","name":"","nameEn":"","subtitle":"","setName":"","setCode":"","number":"","rarity":"","finish":"Normal","language":"PT","illustrator":"","stats":"","grading":{"company":"","grade":"","cert":""},"condition":"NM","confidence":"alta","check":""}',
  'Field notes: game is "Pokémon" or "Lorcana". kind is "Carta avulsa", "Promo" or "Slab (graduada)". subtitle is the Lorcana version line (e.g. "Snow Queen"), else "". finish is one of Normal, Holo, Reverse Holo, Foil, Full Art. language is one of PT, EN, JP, ES, FR, DE, IT, KO, ZH. stats is a short Portuguese summary such as "HP 120 · Fogo" or "Custo 5 · Âmbar". confidence is "alta" only when name AND number are clearly legible, else "média" or "baixa". check is one short sentence in Brazilian Portuguese saying what to double-check, or "".',
].join('\n');

function geminiIdentify_(images, careful) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_KEY');
  if (!key) throw fail_('no_key', 'Falta a chave do Gemini: menu Scanner › 1. Configurar, na planilha.');
  var parts = [{ text: PROMPT }];
  images.forEach(function (b64) { parts.push({ inlineData: { mimeType: 'image/jpeg', data: b64.replace(/^data:image\/\w+;base64,/, '') } }); });
  var bodyFor = function (model, thinking) {
    var gc = { responseMimeType: 'application/json' };
    if (thinking && /^gemini-3/.test(model)) gc.thinkingConfig = { thinkingLevel: 'low' };   // faster answers on Gemini 3.x
    return JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: gc });
  };
  var custom = props.getProperty('GEMINI_MODEL');
  var ok = props.getProperty('GEMINI_MODEL_OK');
  var list = careful ? NB.CAREFUL_MODELS.slice() : NB.MODELS.slice();
  if (!careful && ok) list.unshift(ok);
  if (custom) list.unshift(custom);
  list = list.filter(function (m, i) { return m && list.indexOf(m) === i; });

  var lastErr = null;
  var call = function (model, thinking) {
    return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
      payload: bodyFor(model, thinking), muteHttpExceptions: true,
    });
  };
  for (var i = 0; i < list.length; i++) {
    var model = list[i];
    var res = call(model, true);
    if (res.getResponseCode() === 400 && /thinking/i.test(res.getContentText())) res = call(model, false);
    var code = res.getResponseCode();
    var text = res.getContentText();
    if (code === 200) {
      var j = JSON.parse(text);
      var cand = (j.candidates || [])[0] || {};
      var out = ((cand.content || {}).parts || []).filter(function (p) { return !p.thought; }).map(function (p) { return p.text || ''; }).join('');
      var data = parseJsonLoose_(out);
      if (!data) { lastErr = fail_('bad_reply', 'O Gemini respondeu fora do formato.'); continue; }
      if (!careful && ok !== model) props.setProperty('GEMINI_MODEL_OK', model);
      return { data: data, model: model };
    }
    var msg = '';
    try { msg = (JSON.parse(text).error || {}).message || ''; } catch (e) {}
    if (code === 404 || (code === 400 && /model|not found|not supported/i.test(msg) && !/api key/i.test(msg))) { lastErr = fail_('no_model', 'Modelo ' + model + ' indisponível.'); continue; }
    if (code === 429) throw fail_('quota', 'Limite do Gemini atingido (plano gratuito). Espere 1 minuto; se continuar, o limite do dia acabou.');
    if (code === 400 && /api key/i.test(msg) || code === 401 || code === 403) throw fail_('bad_key', 'A chave do Gemini foi recusada. Crie outra em aistudio.google.com/apikey e use Scanner › 1. Configurar.');
    if (code >= 500) { lastErr = fail_('ai_down', 'O Gemini não respondeu (' + code + ').'); continue; }
    lastErr = fail_('ai_error', 'Gemini: ' + (msg || ('erro ' + code)));
  }
  throw lastErr || fail_('ai_error', 'Não consegui usar o Gemini.');
}

function testKey_() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
    if (!key) return 'sem chave';
    var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return 'recusada (' + res.getResponseCode() + ')';
    var names = (JSON.parse(res.getContentText()).models || []).map(function (m) { return String(m.name).replace('models/', ''); });
    var first = NB.MODELS.filter(function (m) { return names.indexOf(m) >= 0; })[0];
    if (first) PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL_OK', first);
    return first ? 'ok, modelo ' + first : 'ok, mas nenhum modelo da lista apareceu (' + names.slice(0, 4).join(', ') + '…)';
  } catch (e) { return 'erro: ' + e.message; }
}

function normalizeAi_(r, prefs) {
  var g = r.grading && typeof r.grading === 'object' ? r.grading : {};
  return {
    game: /lorcana/i.test(r.game) ? 'Lorcana' : 'Pokémon',
    kind: oneOf_(r.kind, ['Carta avulsa', 'Promo', 'Slab (graduada)'], 'Carta avulsa'),
    name: str_(r.name), nameEn: str_(r.nameEn), subtitle: str_(r.subtitle),
    setName: str_(r.setName), setCode: str_(r.setCode).toUpperCase(), number: str_(r.number).replace(/\s*\/\s*/, '/'),
    rarity: str_(r.rarity), finish: oneOf_(r.finish, ['Normal', 'Holo', 'Reverse Holo', 'Foil', 'Full Art'], 'Normal'),
    language: (str_(r.language) || 'PT').toUpperCase().slice(0, 2), illustrator: str_(r.illustrator), stats: str_(r.stats),
    grading: { company: str_(g.company), grade: str_(g.grade), cert: str_(g.cert) },
    condition: str_(prefs.condition) || str_(r.condition) || 'NM',
    confidence: /alta/i.test(r.confidence) ? 'alta' : /baixa/i.test(r.confidence) ? 'baixa' : 'média',
    check: str_(r.check), image: '', priceUsd: null, dbId: '',
  };
}

/* ============================== card databases ============================== */

var PTCG_RARITY = {
  'common': 'Comum', 'uncommon': 'Incomum', 'rare': 'Rara', 'rare holo': 'Rara Holo', 'double rare': 'Rara Dupla',
  'ultra rare': 'Ultra Rara', 'illustration rare': 'Ilustração Rara', 'special illustration rare': 'Ilustração Rara Especial',
  'hyper rare': 'Hiper Rara', 'ace spec rare': 'ACE SPEC', 'shiny rare': 'Rara Brilhante', 'shiny ultra rare': 'Ultra Rara Brilhante',
  'promo': 'Promo', 'rare holo v': 'Rara Holo V', 'rare holo vmax': 'Rara Holo VMAX', 'rare ultra': 'Ultra Rara', 'rare secret': 'Secreta',
  'rare rainbow': 'Rara Arco-íris', 'trainer gallery rare holo': 'Galeria de Treinadores', 'mega hyper rare': 'Mega Hiper Rara',
};

function lookupPokemon_(c) {
  var parsed = splitNumber_(c.number);
  var n = parsed.n, total = parsed.total;
  if (!n) return { verified: false, candidates: [] };
  var queries = [];
  if (c.setCode) queries.push('set.ptcgoCode:' + c.setCode + ' number:' + n);
  if (/^\d+$/.test(total)) queries.push('number:' + n + ' set.printedTotal:' + total);
  var en = c.nameEn || (c.language === 'EN' ? c.name : '');
  if (en) queries.push('name:"' + en.replace(/"/g, '') + '" number:' + n);
  var seen = {}, cands = [];
  for (var i = 0; i < queries.length && cands.length < 12; i++) {
    var list = ptcgSearch_(queries[i]);
    list.forEach(function (x) { if (!seen[x.id]) { seen[x.id] = 1; cands.push(x); } });
    if (i === 0 && cands.length === 1) break;
  }
  if (!cands.length) return { verified: false, candidates: [] };
  var nn = norm_(en || c.name);
  cands.forEach(function (x) {
    var s = 0;
    var xn = norm_(x.name);
    if (nn && (xn === nn || xn.indexOf(nn) >= 0 || nn.indexOf(xn) >= 0)) s += 3;
    if (c.setCode && x.set && String(x.set.ptcgoCode || '').toUpperCase() === c.setCode) s += 2;
    if (total && x.set && String(x.set.printedTotal) === String(total).replace(/^0+/, '')) s += 1;
    x._score = s;
  });
  cands.sort(function (a, b) { return b._score - a._score; });
  var best = cands[0], second = cands[1];
  var unique = !second || best._score > second._score;
  var verified = unique && best._score >= 3 || (cands.length === 1 && best._score >= 2);
  return { verified: verified, pick: unique && best._score >= 2 ? ptcgPick_(best) : null, candidates: cands.slice(0, 4).map(ptcgPick_) };
}

function ptcgSearch_(q) {
  var cache = CacheService.getScriptCache();
  var ck = 'ptcg:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, q));
  var hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  var headers = {};
  var k = PropertiesService.getScriptProperties().getProperty('PTCG_KEY');
  if (k) headers['X-Api-Key'] = k;
  var res = UrlFetchApp.fetch('https://api.pokemontcg.io/v2/cards?pageSize=12&select=id,name,number,rarity,set,images,tcgplayer,cardmarket&q=' + encodeURIComponent(q), { headers: headers, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  var data = (JSON.parse(res.getContentText()).data || []).map(function (x) {
    var pr = (x.tcgplayer && x.tcgplayer.prices) || {};
    var market = null;
    ['holofoil', 'normal', 'reverseHolofoil', '1stEditionHolofoil', 'unlimitedHolofoil'].some(function (k2) { if (pr[k2] && pr[k2].market) { market = pr[k2].market; return true; } return false; });
    return {
      id: x.id, name: x.name, number: x.number, rarity: x.rarity || '',
      set: { id: x.set && x.set.id, name: x.set && x.set.name, ptcgoCode: x.set && x.set.ptcgoCode, printedTotal: x.set && x.set.printedTotal },
      image: (x.images && (x.images.small || x.images.large)) || '', priceUsd: market,
    };
  });
  try { cache.put(ck, JSON.stringify(data), 21600); } catch (e) {}
  return data;
}

function ptcgPick_(x) {
  return {
    db: 'pokemontcg', id: x.id, name: x.name, number: x.number + (x.set && x.set.printedTotal ? '/' + x.set.printedTotal : ''),
    setName: x.set && x.set.name, setCode: (x.set && (x.set.ptcgoCode || String(x.set.id || '').toUpperCase())) || '',
    rarity: PTCG_RARITY[String(x.rarity).toLowerCase()] || x.rarity, image: x.image, priceUsd: x.priceUsd,
  };
}

function lookupLorcana_(c) {
  var parsed = splitNumber_(c.number);
  var n = String(parsed.n || '').replace(/^0+(?=\d)/, '');
  var set = String(c.setCode || '').replace(/^S(?=\d)/i, '').trim();
  var hits = [];
  if (n && set) {
    var one = lorcanaCard_(set, n);                       // GET /v0/cards/:set/:number
    if (one) hits = [one];
    else hits = lorcanaSet_(set).filter(function (x) { return String(x.n) === n; });
  }
  if (!hits.length && c.name) {
    hits = lorcanaSearch_(c.name).filter(function (x) { return !n || String(x.n) === n; });
    if (set) { var inSet = hits.filter(function (x) { return String(x.set) === set; }); if (inSet.length) hits = inSet; }
  }
  if (!hits.length) return { verified: false, candidates: [] };
  var nn = norm_(c.name);
  hits.forEach(function (x) { x._score = (nn && norm_(x.name).indexOf(nn.split(' ')[0]) >= 0 ? 2 : 0) + (set && String(x.set) === set ? 1 : 0); });
  hits.sort(function (a, b) { return b._score - a._score; });
  var best = hits[0];
  var unique = hits.length === 1 || best._score > hits[1]._score;
  return { verified: unique && best._score >= 2, pick: unique ? lorPick_(best, c) : null, candidates: hits.slice(0, 4).map(function (x) { return lorPick_(x, c); }) };
}

function lorcanaCard_(set, n) {
  var cache = CacheService.getScriptCache();
  var key = 'lor:card:' + set + ':' + n;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  var res = UrlFetchApp.fetch('https://api.lorcast.com/v0/cards/' + encodeURIComponent(set) + '/' + encodeURIComponent(n), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  var raw = JSON.parse(res.getContentText());
  if (!raw || !raw.name) return null;
  var slim = lorSlim_(raw);
  try { cache.put(key, JSON.stringify(slim), 21600); } catch (e) {}
  return slim;
}

function lorcanaSet_(set) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('lor:set:' + set);
  if (hit) return JSON.parse(hit);
  var res = UrlFetchApp.fetch('https://api.lorcast.com/v0/sets/' + encodeURIComponent(set) + '/cards', { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  var raw = JSON.parse(res.getContentText());
  var list = (Array.isArray(raw) ? raw : (raw.results || raw.data || [])).map(lorSlim_);
  try { cache.put('lor:set:' + set, JSON.stringify(list), 21600); } catch (e) {}
  return list;
}

function lorcanaSearch_(name) {
  var res = UrlFetchApp.fetch('https://api.lorcast.com/v0/cards/search?q=' + encodeURIComponent(name), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  var raw = JSON.parse(res.getContentText());
  return (Array.isArray(raw) ? raw : (raw.results || raw.data || [])).map(lorSlim_);
}

function lorSlim_(x) {
  var iu = (x.image_uris && (x.image_uris.digital || x.image_uris)) || {};
  var pr = x.prices || {};
  return {
    name: x.name || '', version: x.version || '', n: String(x.collector_number || '').replace(/^0+(?=\d)/, ''),
    set: String((x.set && (x.set.code || x.set.id)) || ''), setName: (x.set && x.set.name) || '',
    rarity: String(x.rarity || '').replace(/_/g, ' '), image: iu.small || iu.normal || '',
    usd: pr.usd ? Number(pr.usd) : null, usdFoil: pr.usd_foil ? Number(pr.usd_foil) : null,
  };
}

function lorPick_(x, c) {
  var foil = /foil/i.test(c.finish || '');
  return {
    db: 'lorcast', id: x.set + '-' + x.n, name: x.name, subtitle: x.version, number: x.n,
    setName: x.setName, setCode: x.set, rarity: x.rarity.replace(/\b\w/g, function (s) { return s.toUpperCase(); }),
    image: x.image, priceUsd: foil && x.usdFoil ? x.usdFoil : x.usd,
  };
}

function applyPick_(card, pk) {
  var printedNumber = card.number;
  if (card.language === 'EN' || !card.name) card.name = pk.name || card.name;
  if (card.game === 'Lorcana') {
    if (pk.subtitle && (!card.subtitle || card.language === 'EN')) card.subtitle = pk.subtitle;
    var tot = splitNumber_(printedNumber).total;
    card.number = pk.number + (tot ? '/' + tot : '');
  } else {
    card.number = printedNumber || pk.number;
  }
  card.setName = pk.setName || card.setName;
  card.setCode = pk.setCode || card.setCode;
  card.rarity = pk.rarity || card.rarity;
  card.image = pk.image || '';
  card.priceUsd = pk.priceUsd || null;
  card.dbId = pk.id || '';
}

/* ============================== SKU, names, matching ============================== */

function productName_(c) { return c.game === 'Lorcana' && c.subtitle ? c.name + ' – ' + c.subtitle : (c.name || ''); }

function buildSku_(c) {
  var g = c.game === 'Lorcana' ? 'LOR' : 'POK';
  var set = String(c.setCode || '').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 6);
  if (c.game === 'Lorcana' && /^\d+$/.test(set)) set = 'S' + set;
  var n = String(c.number || '').split('/')[0].toUpperCase().replace(/[^A-Z0-9]/g, '').replace(/^0+(?=\d)/, '') || '0';
  var parts = [g, set || 'X', n, (c.language || 'XX').toUpperCase().slice(0, 2)];
  if (/reverse/i.test(c.finish || '')) parts.push('RH');
  if (c.game === 'Lorcana' && /foil/i.test(c.finish || '')) parts.push('F');
  var gr = c.grading || {};
  if (c.kind === 'Slab (graduada)' && (gr.company || gr.grade)) parts.push(String((gr.company || '') + (gr.grade || '')).toUpperCase().replace(/[^A-Z0-9]/g, ''));
  return parts.join('-');
}

function obsFor_(c) {
  var bits = [];
  if (c.stats) bits.push(c.stats);
  if (c.illustrator) bits.push('Ilustr.: ' + c.illustrator);
  if (c.priceUsd) {
    var usd = Number(c.priceUsd);
    var fx = dollar_();
    bits.push('Mercado EUA US$ ' + usd.toFixed(2).replace('.', ',') + (fx ? ' (≈ R$ ' + Math.round(usd * fx) + ')' : ''));
  }
  var g = c.grading || {};
  if (g.cert) bits.push(('Cert ' + (g.company || '') + ' ' + g.cert).replace(/\s+/g, ' '));
  if (c.verified === false && c.game) bits.push('conferir: não achei no banco de cartas');
  return bits.join(' · ');
}

function dollar_() {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(NB.CONFIG);
    if (!sh) return null;
    var vals = sh.getRange('A3:B30').getValues();
    for (var i = 0; i < vals.length; i++) if (/d[oó]lar usado/i.test(String(vals[i][0]))) { var v = Number(vals[i][1]); return v > 0 ? v : null; }
  } catch (e) {}
  return null;
}

function findMatch_(c, products) {
  var sku = c.sku || buildSku_(c);
  var nm = norm_(productName_(c)), nh = numHead_(c.number);
  var best = null;
  for (var i = 0; i < products.length; i++) {
    var p = products[i];
    var psku = String(p[0] || '');
    if (!psku) continue;
    if (psku === sku) return { sku: psku, label: String(p[3] || psku) };
    if (String(p[1] || '') !== c.game) continue;
    if (String(p[7] || '') && c.language && String(p[7]).toUpperCase() !== c.language) continue;
    var pn = norm_(p[3]), ph = numHead_(p[5]);
    var words = nm.split(' ').filter(function (w) { return w.length > 2; });
    var overlap = words.length ? words.filter(function (w) { return pn.indexOf(w) >= 0; }).length / words.length : 0;
    var okNum = nh && ph ? nh === ph : false;
    if (okNum && overlap >= 0.5) {
      var score = overlap + 1;
      if (!best || score > best.score) best = { sku: psku, label: String(p[3] || '') + (p[5] ? ' · ' + String(p[5]) : ''), score: score };
    }
  }
  return best ? { sku: best.sku, label: best.label } : null;
}

/* ============================== helpers ============================== */

function splitNumber_(s) {
  var t = String(s || '').toUpperCase().replace(/\s+/g, '');
  var m = t.match(/([A-Z]{0,4}\d{1,4}[A-Z]?)\/([A-Z]{0,4}\d{1,4})/);
  var n = m ? m[1] : ((t.match(/[A-Z]{0,4}\d{1,4}[A-Z]?/) || [''])[0]);
  n = n.replace(/^0+(?=\d)/, '');
  if (/^SVP/.test(n)) n = n.replace(/^SVP0*/, 'SVP');
  return { n: n, total: m ? m[2].replace(/^0+(?=\d)/, '') : '' };
}
function norm_(v) { return String(v || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim(); }
function numHead_(v) { return String(v || '').split('/')[0].replace(/^'/, '').replace(/^0+/, '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function str_(v) { return typeof v === 'string' ? v.trim() : (v === null || v === undefined ? '' : String(v).trim()); }
function oneOf_(v, list, d) { return list.indexOf(v) >= 0 ? v : d; }
function num_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var n = Number(String(v).replace(/\s|R\$/g, '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
  return isFinite(n) ? n : null;
}
function numOrBlank_(v) { var n = num_(v); return n === null ? '' : n; }
function txt_(v) { var s = String(v === null || v === undefined ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function cleanId_(v) { var s = String(v || '').replace(/[^a-z0-9]/gi, '').slice(0, 40); if (s.length < 6) throw fail_('bad_id', 'Pedido sem identificador.'); return s; }
function parseJsonLoose_(t) {
  if (!t) return null;
  try { return JSON.parse(t); } catch (e) {}
  var a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a >= 0 && b > a) { try { return JSON.parse(t.slice(a, b + 1)); } catch (e2) {} }
  return null;
}
function newToken_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 24); }
function escHtml_(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function fail_(code, message) { var e = new Error(message); e.code = code; return e; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
