// Número do WhatsApp do chefe (DDI + DDD + número, só dígitos)
const WPP_CHEFE = '5533999090467';

// Colunas exportadas (cabeçalho na planilha -> rótulo no PDF)
const COLS = [
  ['PLACA','Placa'],['ENQUADRAMENTO','Enquadramento'],['PONTOS','Pts'],
  ['LOCAL INFRACAO','Local'],['MUNICIPIO INFRACAO','Município'],['GRAVIDADE','Gravidade'],
  ['INFRATOR','Infrator'],['DATA INFRACAO','Data'],['HORARIO INFRACAO','Hora'],
  ['VALOR INFRACAO','Valor'],['DATA VENCIMENTO','Vencimento']
];
const REQUIRED = ['EMPRESA', ...COLS.map(c => c[0])];

const $ = id => document.getElementById(id);
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ').trim().toUpperCase();
const brl = n => n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const toNum = v => { if (typeof v === 'number') return v; const s = String(v ?? '').replace(/[R$\s]/g,''); if (!s) return 0; return parseFloat(s.includes(',') ? s.replace(/\./g,'').replace(',','.') : s) || 0; };
const pad = n => String(n).padStart(2,'0');
const fmtCell = (key, v) => {
  if (v instanceof Date) return key.startsWith('HORARIO') ? `${pad(v.getHours())}:${pad(v.getMinutes())}` : `${pad(v.getDate())}/${pad(v.getMonth()+1)}/${v.getFullYear()}`;
  if (key === 'VALOR INFRACAO') return brl(toNum(v));
  let s = String(v ?? '').trim();
  if (key === 'HORARIO INFRACAO' && /^\d{2}:\d{2}:\d{2}$/.test(s)) s = s.slice(0,5);
  return s;
};
const parseDate = s => { const m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(s); return m ? new Date(+m[3], m[2]-1, +m[1]).getTime() : 0; };
const safeName = s => s.normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^\w\- ]+/g,'').trim().replace(/\s+/g,'_').slice(0,80);

let groups = [];   // [{name, rows:[...], pts, total}]
let fileName = '';

function showMsg(text, type){ const m = $('msg'); m.textContent = text; m.className = 'msg ' + type; m.hidden = !text; }

async function loadFile(file){
  showMsg('', '');
  fileName = file.name;
  $('fileInfo').textContent = 'Lendo ' + file.name + '…';
  try {
    const wb = XLSX.read(await file.arrayBuffer(), {type:'array', cellDates:true});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, {defval:'', raw:true});
    if (!raw.length) throw new Error('A planilha está vazia.');
    // mapeia cabeçalhos normalizados -> nome real
    const map = {}; Object.keys(raw[0]).forEach(k => map[norm(k)] = k);
    const missing = REQUIRED.filter(c => !(c in map));
    if (missing.length) throw new Error('Colunas não encontradas: ' + missing.join(', '));

    const byCo = new Map();
    for (const r of raw) {
      const emp = String(r[map.EMPRESA] ?? '').replace(/\s+/g,' ').trim();
      if (!emp && !r[map.PLACA]) continue;
      const key = norm(emp) || '(SEM EMPRESA)';
      if (!byCo.has(key)) byCo.set(key, {name: emp || '(Sem empresa)', rows: [], pts:0, total:0});
      const g = byCo.get(key);
      const row = {}; REQUIRED.forEach(c => row[c] = r[map[c]]);
      g.rows.push(row);
      g.pts += toNum(row.PONTOS); g.total += toNum(row['VALOR INFRACAO']);
    }
    groups = [...byCo.values()].sort((a,b) => a.name.localeCompare(b.name,'pt-BR'));
    groups.forEach(g => g.rows.sort((a,b) => String(a.PLACA).localeCompare(String(b.PLACA)) || parseDate(fmtCell('DATA INFRACAO',a['DATA INFRACAO'])) - parseDate(fmtCell('DATA INFRACAO',b['DATA INFRACAO']))));
    const n = groups.reduce((s,g) => s + g.rows.length, 0);
    $('fileInfo').textContent = `${file.name} — ${n} infrações, ${groups.length} empresas`;
    showMsg('Planilha carregada com sucesso.', 'ok');
    $('btnClear').hidden = false;
    render(true);
  } catch (e) {
    groups = []; $('resultPanel').hidden = true;
    $('fileInfo').textContent = 'Nenhum arquivo carregado';
    showMsg('Não foi possível ler o arquivo: ' + e.message, 'err');
  }
}

function render(animate){
  const q = norm($('search').value);
  const tb = $('rows'); tb.innerHTML = '';
  groups.forEach((g, i) => {
    if (q && !norm(g.name).includes(q)) return;
    const tr = document.createElement('tr');
    if (animate) { tr.className = 'reveal'; tr.style.setProperty('--i', tb.children.length); }
    tr.innerHTML = `<td class="chk"><input type="checkbox" data-i="${i}" ${g.sel !== false ? 'checked' : ''}></td>
      <td></td><td class="num">${g.rows.length}</td><td class="num">${g.pts}</td><td class="num">${brl(g.total)}</td>
      <td class="act"><button class="btn-link" data-pdf="${i}">PDF</button></td>`;
    tr.children[1].textContent = g.name;
    tb.appendChild(tr);
  });
  updateFooter(animate);
  const panel = $('resultPanel');
  if (panel.hidden) { panel.hidden = false; panel.classList.remove('enter'); void panel.offsetWidth; panel.classList.add('enter'); }
}

function selected(){ return groups.filter(g => g.sel !== false); }
// Efeito 4 — Number ticker: totais contam até o novo valor
function tick(el, to, fmt){
  const from = el._v ?? 0; el._v = to;
  if (matchMedia('(prefers-reduced-motion: reduce)').matches || from === to) { el.textContent = fmt(to); return; }
  const t0 = performance.now(), d = 450;
  cancelAnimationFrame(el._raf);
  const step = t => { const k = Math.min(1, (t - t0) / d), e = 1 - Math.pow(1 - k, 3);
    el.textContent = fmt(from + (to - from) * e); if (k < 1) el._raf = requestAnimationFrame(step); };
  el._raf = requestAnimationFrame(step);
}
function updateFooter(){
  const s = selected();
  $('ftLabel').textContent = `${s.length} de ${groups.length} selecionadas`;
  const int = v => String(Math.round(v));
  tick($('ftQtd'), s.reduce((a,g) => a + g.rows.length, 0), int);
  tick($('ftPts'), s.reduce((a,g) => a + g.pts, 0), int);
  tick($('ftVal'), s.reduce((a,g) => a + g.total, 0), brl);
  $('btnAll').disabled = $('btnEach').disabled = $('btnWpp').disabled = !s.length;
  $('chkAll').checked = s.length === groups.length;
}

// ---------- PDF ----------
function newDoc(){ return new jspdf.jsPDF({orientation:'landscape', unit:'mm', format:'a4'}); }
const NOW = () => new Date().toLocaleString('pt-BR',{dateStyle:'short',timeStyle:'short'});

function drawCompany(doc, g, first){
  if (!first) doc.addPage();
  const W = doc.internal.pageSize.getWidth();
  const header = () => {
    doc.addImage(LOGO, 'PNG', 10, 8, 12, 12);
    doc.setFont('helvetica','bold'); doc.setFontSize(12); doc.setTextColor(31,41,55);
    doc.text(g.name, 25, 13);
    doc.setFont('helvetica','normal'); doc.setFontSize(8); doc.setTextColor(107,114,128);
    doc.text('Relatório de Infrações', 25, 18);
    doc.text('Emitido em ' + NOW(), W - 10, 13, {align:'right'});
    doc.text(`${g.rows.length} infrações · ${g.pts} pontos · ${brl(g.total)}`, W - 10, 18, {align:'right'});
    doc.setDrawColor(226,229,233); doc.line(10, 22, W - 10, 22);
  };
  doc.autoTable({
    startY: 25,
    margin: {top: 25, left: 10, right: 10, bottom: 12},
    head: [COLS.map(c => c[1])],
    body: g.rows.map(r => COLS.map(([k]) => fmtCell(k, r[k]))),
    foot: [[{content:'Total', colSpan:2}, String(g.pts), {content:'', colSpan:6}, brl(g.total), '']],
    showFoot: 'lastPage',
    theme: 'grid',
    styles: {font:'helvetica', fontSize:7, cellPadding:1.4, textColor:[31,41,55], lineColor:[226,229,233], lineWidth:.1, valign:'middle'},
    headStyles: {fillColor:[30,78,140], textColor:255, fontStyle:'bold', fontSize:7},
    footStyles: {fillColor:[234,242,255], textColor:[31,41,55], fontStyle:'bold'},
    alternateRowStyles: {fillColor:[249,250,251]},
    columnStyles: {
      0:{cellWidth:17, fontStyle:'bold'}, 1:{cellWidth:62}, 2:{cellWidth:9, halign:'center'},
      3:{cellWidth:'auto'}, 4:{cellWidth:28}, 5:{cellWidth:17}, 6:{cellWidth:20},
      7:{cellWidth:17, halign:'center'}, 8:{cellWidth:11, halign:'center'},
      9:{cellWidth:20, halign:'right'}, 10:{cellWidth:18, halign:'center'}
    },
    didDrawPage: header
  });
}

function addPageNumbers(doc){
  const n = doc.getNumberOfPages(), W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i); doc.setFontSize(7); doc.setTextColor(107,114,128);
    doc.text(`Página ${i} de ${n}`, W - 10, H - 6, {align:'right'});
  }
}

function buildPdf(list){
  const doc = newDoc();
  list.forEach((g, i) => drawCompany(doc, g, i === 0));
  addPageNumbers(doc);
  return doc;
}
function pdfFor(list, name){ buildPdf(list).save(name); }

// estado de carregamento nos botões (shimmer) — deixa o navegador pintar antes do trabalho pesado
async function busy(btn, fn){
  btn.classList.add('busy');
  await new Promise(r => setTimeout(r, 30));
  try { await fn(); } catch (e) { toast('Erro ao gerar PDF: ' + e.message, 'err'); }
  finally { btn.classList.remove('busy'); }
}

function toast(text, type = '', ms = 5000){
  const t = document.createElement('div'); t.className = 'toast ' + type; t.textContent = text;
  $('toasts').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 200); }, ms);
}

function clearAll(){
  groups = []; fileName = '';
  $('file').value = ''; $('search').value = '';
  $('rows').innerHTML = '';
  ['ftQtd','ftPts','ftVal'].forEach(id => $(id)._v = 0);
  $('resultPanel').hidden = true; $('btnClear').hidden = true;
  $('fileInfo').textContent = 'Nenhum arquivo carregado';
  showMsg('', '');
  toast('Tela limpa. Pode carregar outra planilha.');
}

async function sendWhatsApp(){
  const list = selected();
  const nome = list.length === 1 ? `Infracoes_${safeName(list[0].name)}.pdf` : `Infracoes_${new Date().toISOString().slice(0,10)}.pdf`;
  const qtd = list.reduce((a,g) => a + g.rows.length, 0), total = list.reduce((a,g) => a + g.total, 0);
  const texto = `Segue relatório de infrações (${list.length} empresa${list.length > 1 ? 's' : ''}, ${qtd} infrações, ${brl(total)}).\n` +
                list.map(g => `• ${g.name}: ${g.rows.length} — ${brl(g.total)}`).join('\n');
  const url = `https://web.whatsapp.com/send?phone=${WPP_CHEFE}&text=${encodeURIComponent(texto)}`;
  buildPdf(list).save(nome);
  // abre dentro do próprio clique para o navegador não bloquear o pop-up
  if (!window.open(url, 'whatsapp')) location.href = url;
  toast(`PDF "${nome}" baixado. No WhatsApp Web, arraste o arquivo da pasta Downloads para a conversa e envie.`, 'ok', 9000);
}

// ---------- eventos ----------
$('file').addEventListener('change', e => e.target.files[0] && loadFile(e.target.files[0]));
const drop = $('drop');
['dragenter','dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));
$('search').addEventListener('input', render);
$('rows').addEventListener('change', e => { const i = e.target.dataset.i; if (i != null) { groups[i].sel = e.target.checked; updateFooter(); } });
$('rows').addEventListener('click', e => { const i = e.target.dataset.pdf; if (i != null) pdfFor([groups[i]], `Infracoes_${safeName(groups[i].name)}.pdf`); });
$('chkAll').addEventListener('change', e => { groups.forEach(g => g.sel = e.target.checked); render(); });
$('btnAll').addEventListener('click', e => busy(e.currentTarget, () => pdfFor(selected(), `Infracoes_${new Date().toISOString().slice(0,10)}.pdf`)));
$('btnEach').addEventListener('click', e => busy(e.currentTarget, async () => {
  for (const g of selected()) { pdfFor([g], `Infracoes_${safeName(g.name)}.pdf`); await new Promise(r => setTimeout(r, 400)); }
  toast(`${selected().length} PDF(s) baixado(s).`, 'ok');
}));
$('btnWpp').addEventListener('click', sendWhatsApp);
$('btnClear').addEventListener('click', clearAll);
