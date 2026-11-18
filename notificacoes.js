// ===================== Aba Notificações =====================
// Lê os PDFs gerados na aba Relatório (dados embutidos), agrupa por município (empresa)
// e preenche o modelo Word (modelos/modelo-notificacao.docx) trocando os campos {{...}}.

let nfItems = [];   // [{key, empresa, cidade, uf, rows:[{placa, enq, pontos, valor, data}], sel}]
let nfMeta = {emitido: ''};
const nfLoaded = new Set();   // conteúdo dos PDFs já lidos — o mesmo PDF duas vezes não duplica

const MESES = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
const NIC_RE = /NAO IDENTIFICACAO DO CONDUTOR/;   // multa por não identificação do condutor (pessoa jurídica)

// ---------- abas ----------
document.querySelectorAll('.tab').forEach(t => t.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(o => { o.classList.toggle('active', o === t); o.setAttribute('aria-selected', o === t); });
  ['tabRelatorio','tabNotif','tabComp'].forEach(id => $(id).hidden = id !== t.dataset.tab);
  if (t.dataset.tab === 'tabNotif') $('nfFromSheet').hidden = !allRows.length;
}));

// ---------- utilitários de texto ----------
const LOWER = new Set(['de','da','do','das','dos','e']);
function titleCase(s){
  return s.toLowerCase().split(' ').map((w, i) => (i && LOWER.has(w)) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}
// "PREFEITURA MUNICIPAL DE MATHIAS LOBATO" -> "Mathias Lobato"
function cidadeDe(empresa){
  let s = String(empresa).replace(/\s+/g,' ').trim();
  s = s.replace(/^(prefeitura\s+municipal\s+de|prefeitura\s+de|munic[ií]pio\s+de|c[aâ]mara\s+municipal\s+de|fundo\s+municipal\s+de\s+sa[uú]de\s+de)\s+/i, '');
  s = s.replace(/\s*[-–\/]\s*[A-Z]{2}$/,'');           // tira "/MG" ou "- MG" no fim
  return titleCase(s);
}
const listaBR = arr => arr.length <= 1 ? (arr[0] || '') : arr.slice(0, -1).join(', ') + ' e ' + arr[arr.length - 1];
const brD = iso => iso ? iso.split('-').reverse().join('/') : '';
const extenso = iso => { const [y,m,d] = iso.split('-'); return `${+d} de ${MESES[m-1]} de ${y}`; };
const plural = (n, um, varios) => `${n.toLocaleString('pt-BR')} ${n === 1 ? um : varios}`;
const xmlEsc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const money = n => n.toLocaleString('pt-BR',{minimumFractionDigits:2, maximumFractionDigits:2});

// ---------- carregar dados ----------
function addEmpresas(empresas){
  const map = new Map(nfItems.map(it => [it.key, it]));
  for (const e of empresas) {
    const key = norm(e.n);
    if (!map.has(key)) map.set(key, {key, empresa: e.n, cidade: cidadeDe(e.n), uf: 'MG', rows: [], sel: true});
    const it = map.get(key);
    for (const [placa, enq, pontos, valor, data] of e.r) it.rows.push({placa, enq, pontos, valor, data});
  }
  nfItems = [...map.values()].sort((a,b) => a.cidade.localeCompare(b.cidade,'pt-BR'));
}

function readPdfData(buf){
  // o jsPDF grava o campo Keywords sem compressão: "/Keywords (HSDATA1:<base64>)"
  const txt = new TextDecoder('latin1').decode(buf);
  const m = /\/Keywords\s*\(HSDATA1:([A-Za-z0-9+\/=]+)\)/.exec(txt);
  if (!m) return null;
  if (nfLoaded.has(m[1])) return {dup: true};
  nfLoaded.add(m[1]);
  const bin = atob(m[1]), bytes = Uint8Array.from(bin, c => c.charCodeAt(0));
  return JSON.parse(new TextDecoder().decode(bytes));
}

function nfMsg(text, type){ const m = $('nfMsg'); m.textContent = text; m.className = 'msg ' + type; m.hidden = !text; }

async function loadPdfs(files){
  const ok = [], bad = [];
  for (const f of files) {
    try {
      const data = readPdfData(await f.arrayBuffer());
      if (!data) { bad.push(f.name); continue; }
      if (data.dup) { toast(`${f.name} já tinha sido carregado.`); continue; }
      addEmpresas(data.empresas);
      if (data.emitido && (!nfMeta.emitido || data.emitido > nfMeta.emitido)) nfMeta.emitido = data.emitido;
      ok.push(f.name);
    } catch { bad.push(f.name); }
  }
  if (bad.length) nfMsg(`Não reconheci: ${bad.join(', ')}. Use um PDF gerado por este sistema (aba Relatório) — PDFs gerados antes desta versão precisam ser gerados de novo.`, 'err');
  else if (ok.length) nfMsg(`${ok.length} PDF(s) lido(s) com sucesso.`, 'ok');
  if (ok.length) showNotif();
}

function fromSheet(){
  const list = selected();
  if (!list.length) return toast('Nenhuma empresa selecionada na aba Relatório.', 'err');
  nfItems = []; nfLoaded.clear();   // planilha substitui o que estava carregado
  addEmpresas(list.map(g => ({n: g.name, r: g.rows.map(r => [String(r.PLACA ?? '').trim(), String(r.ENQUADRAMENTO ?? '').trim(), toNum(r.PONTOS), toNum(r['VALOR INFRACAO']), isoDate(r['DATA INFRACAO'])])})));
  nfMeta.emitido = isoDate(new Date());
  nfMsg(`Dados da planilha carregados${periodText() ? ' (' + periodText() + ')' : ''}.`, 'ok');
  showNotif();
}

function showNotif(){
  const today = isoDate(new Date());
  if (!$('nfDataRel').value) $('nfDataRel').value = nfMeta.emitido || today;
  if (!$('nfDataCom').value) $('nfDataCom').value = today;
  const n = nfItems.reduce((a, it) => a + it.rows.length, 0);
  $('nfInfo').textContent = `${n} infrações · ${nfItems.length} município(s)`;
  $('nfClear').hidden = false;
  renderNotif(true);
}

// ---------- tabela ----------
function stats(it){
  const nic = it.rows.filter(r => NIC_RE.test(norm(r.enq)));
  return {qtd: it.rows.length, valor: it.rows.reduce((a, r) => a + r.valor, 0), nic};
}

function renderNotif(animate){
  const tb = $('nfRows'); tb.innerHTML = '';
  nfItems.forEach((it, i) => {
    const s = stats(it);
    const tr = document.createElement('tr');
    if (animate) { tr.className = 'reveal'; tr.style.setProperty('--i', i); }
    tr.innerHTML = `<td class="chk"><input type="checkbox" data-i="${i}" ${it.sel ? 'checked' : ''}></td>
      <td><input class="cell" data-f="cidade" data-i="${i}"><span class="sub"></span></td>
      <td><input class="cell uf" data-f="uf" data-i="${i}" maxlength="2"></td>
      <td class="num">${s.qtd}</td><td class="num">${brl(s.valor)}</td>
      <td>${s.nic.length ? `<span class="tag warn">Sim · ${s.nic.length} multa(s)</span>` : '<span class="tag none">Não — parágrafo removido</span>'}</td>
      <td class="act"><button class="btn-link" data-doc="${i}">Word</button><button class="btn-link" data-pdf="${i}">PDF</button></td>`;
    tr.querySelector('[data-f=cidade]').value = it.cidade;
    tr.querySelector('[data-f=uf]').value = it.uf;
    tr.querySelector('.sub').textContent = it.empresa;
    tb.appendChild(tr);
  });
  $('nfChkAll').checked = nfItems.every(it => it.sel);
  $('nfAll').disabled = !nfItems.some(it => it.sel);
  $('nfPanel').hidden = false;
}

// ---------- geração do Word ----------
function campos(it){
  const s = stats(it);
  const pontos = it.rows.reduce((a, r) => a + r.pontos, 0);
  const placas = [...new Set(it.rows.map(r => r.placa).filter(Boolean))].sort();
  const datas = it.rows.map(r => r.data).filter(Boolean).sort();
  const ini = datas[0], fim = datas[datas.length - 1];
  const nicVals = [...new Set(s.nic.map(r => r.valor))].sort((a,b) => a - b).map(v => 'R$ ' + money(v));
  return {
    CIDADE: it.cidade.trim(), UF: (it.uf || '').trim().toUpperCase(),
    DATA_RELATORIO: brD($('nfDataRel').value), DATA_EXTENSO: extenso($('nfDataCom').value),
    INFRACOES: plural(s.qtd, 'infração', 'infrações'),
    PONTOS: plural(pontos, 'ponto', 'pontos'),
    VALOR_TOTAL: money(s.valor),
    PLACAS: (placas.length === 1 ? 'a placa ' : 'as placas ') + listaBR(placas),
    PERIODO: !ini ? '' : ini === fim ? `em ${brD(ini)}` : `entre ${brD(ini)} e ${brD(fim)}`,
    // null = parágrafo inteiro sai do documento
    NIC_VALORES: nicVals.length ? `${nicVals.length === 1 ? 'com valor de' : 'com valores de'} ${listaBR(nicVals)}` : null
  };
}

function preencher(xml, vals){
  // 1) remove parágrafos que contêm um campo sem valor (ex.: não identificação do condutor)
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    const tags = p.match(/\{\{(\w+)\}\}/g) || [];
    return tags.some(t => vals[t.slice(2, -2)] === null) ? '' : p;
  });
  // 2) troca os campos restantes
  return xml.replace(/\{\{(\w+)\}\}/g, (m, k) => k in vals ? xmlEsc(vals[k] ?? '') : m);
}

let modeloZip = null;
async function modelo(){
  if (!modeloZip) {
    const bin = atob(MODELO_NOTIFICACAO);
    modeloZip = Uint8Array.from(bin, c => c.charCodeAt(0));
  }
  return JSZip.loadAsync(modeloZip);
}

async function preenchido(it){
  const zip = await modelo();
  const vals = campos(it);
  for (const f of Object.keys(zip.files).filter(f => /^word\/(document|header\d*|footer\d*)\.xml$/.test(f))) {
    zip.file(f, preencher(await zip.file(f).async('string'), vals));
  }
  return zip;
}
async function gerarDocx(it){
  return (await preenchido(it)).generateAsync({type:'blob', mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression:'DEFLATE'});
}
async function gerarPdf(it){ return docxParaPdf(await preenchido(it)); }

const docName = it => `Comunicacao_Administrativa_${safeName(it.cidade || it.empresa)}.docx`;
function baixar(blob, nome){
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nome;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function validar(list){
  if (!$('nfDataRel').value || !$('nfDataCom').value) { toast('Preencha a data do relatório e a data da comunicação.', 'err'); return false; }
  const semNome = list.filter(it => !it.cidade.trim());
  if (semNome.length) { toast('Preencha o nome do município em todas as linhas selecionadas.', 'err'); return false; }
  return true;
}

// ---------- eventos ----------
$('nfFile').addEventListener('change', e => { if (e.target.files.length) loadPdfs([...e.target.files]); e.target.value = ''; });
const nfDrop = $('nfDrop');
['dragenter','dragover'].forEach(ev => nfDrop.addEventListener(ev, e => { e.preventDefault(); nfDrop.classList.add('over'); }));
['dragleave','drop'].forEach(ev => nfDrop.addEventListener(ev, e => { e.preventDefault(); nfDrop.classList.remove('over'); }));
nfDrop.addEventListener('drop', e => e.dataTransfer.files.length && loadPdfs([...e.dataTransfer.files]));
$('nfFromSheet').addEventListener('click', fromSheet);

$('nfRows').addEventListener('input', e => { const i = e.target.dataset.i, f = e.target.dataset.f; if (f) nfItems[i][f] = e.target.value; });
$('nfRows').addEventListener('change', e => {
  if (e.target.type === 'checkbox') { nfItems[e.target.dataset.i].sel = e.target.checked; $('nfChkAll').checked = nfItems.every(it => it.sel); $('nfAll').disabled = !nfItems.some(it => it.sel); }
});
$('nfRows').addEventListener('click', e => {
  const i = e.target.dataset.doc ?? e.target.dataset.pdf; if (i == null) return;
  const it = nfItems[i]; if (!validar([it])) return;
  const pdf = e.target.dataset.pdf != null;
  busy(e.target, async () => {
    if (pdf) baixar(await gerarPdf(it), docName(it).replace(/\.docx$/, '.pdf'));
    else baixar(await gerarDocx(it), docName(it));
    toast(`Comunicação de ${it.cidade} gerada.`, 'ok');
  });
});
$('nfChkAll').addEventListener('change', e => { nfItems.forEach(it => it.sel = e.target.checked); renderNotif(); });
$('nfAll').addEventListener('click', e => {
  const list = nfItems.filter(it => it.sel); if (!validar(list)) return;
  busy(e.currentTarget, async () => {
    // zip com Word e PDF de cada município, em pastas separadas
    const out = new JSZip();
    for (const it of list) {
      out.file('Word/' + docName(it), await gerarDocx(it));
      out.file('PDF/' + docName(it).replace(/\.docx$/, '.pdf'), await gerarPdf(it));
    }
    baixar(await out.generateAsync({type:'blob'}), `Comunicacoes_${$('nfDataCom').value}.zip`);
    toast(`${list.length} comunicação(ões) gerada(s).`, 'ok');
  });
});
$('nfClear').addEventListener('click', () => {
  nfItems = []; nfMeta = {emitido: ''}; nfLoaded.clear();
  $('nfRows').innerHTML = ''; $('nfPanel').hidden = true; $('nfClear').hidden = true;
  $('nfDataRel').value = $('nfDataCom').value = '';
  $('nfInfo').textContent = 'Nenhum PDF carregado'; nfMsg('', '');
});
