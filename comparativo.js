// ===================== Aba Comparativo =====================
// Lê PDFs gerados no site (dados embutidos), um por período, e compara uma empresa mês a mês.

let cpPeriods = [];   // [{id, label, file, emitido, empresas: Map(key -> {name, qtd, pontos, valor, nic})}]
const MES = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
const CP_ALL = '__todas__';

function cpParse(buf){
  const txt = new TextDecoder('latin1').decode(buf);
  const m = /\/Keywords\s*\(HSDATA1:([A-Za-z0-9+\/=]+)\)/.exec(txt);
  if (!m) return null;
  const bin = atob(m[1]);
  return {sig: m[1], data: JSON.parse(new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0))))};
}

// nome sugerido do período: mês predominante das datas das infrações, senão o mês de emissão
function cpLabel(data){
  const cnt = {};
  data.empresas.forEach(e => e.r.forEach(r => { const k = (r[4] || '').slice(0, 7); if (k) cnt[k] = (cnt[k] || 0) + 1; }));
  let best = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0]?.[0];
  if (!best || Object.keys(cnt).length > 3) best = (data.emitido || '').slice(0, 7);   // período misto: usa a emissão
  if (!best) return 'Período';
  const [y, mo] = best.split('-'); return `${MES[+mo - 1]}/${y}`;
}

async function cpLoad(files){
  const bad = [];
  for (const f of files) {
    try {
      const r = cpParse(await f.arrayBuffer());
      if (!r) { bad.push(f.name); continue; }
      if (cpPeriods.some(p => p.sig === r.sig)) { toast(`${f.name} já está na lista.`); continue; }
      const empresas = new Map();
      for (const e of r.data.empresas) {
        const key = norm(e.n);
        const it = empresas.get(key) || {name: e.n, qtd: 0, pontos: 0, valor: 0, nic: 0};
        for (const [, enq, pts, val] of e.r) { it.qtd++; it.pontos += pts; it.valor += val; if (/NAO IDENTIFICACAO DO CONDUTOR/.test(norm(enq))) it.nic++; }
        empresas.set(key, it);
      }
      cpPeriods.push({sig: r.sig, label: cpLabel(r.data), file: f.name, emitido: r.data.emitido || '', empresas});
    } catch { bad.push(f.name); }
  }
  const msg = $('cpMsg');
  if (bad.length) { msg.textContent = `Não reconheci: ${bad.join(', ')}. Use PDFs gerados pela aba Relatório deste site.`; msg.className = 'msg err'; msg.hidden = false; }
  else { msg.hidden = true; }
  cpRender();
}

function cpRenderList(){
  const tb = $('cpList'); tb.innerHTML = '';
  cpPeriods.forEach((p, i) => {
    const n = [...p.empresas.values()].reduce((a, e) => a + e.qtd, 0);
    const tr = document.createElement('tr');
    tr.innerHTML = `<td><input class="cell" data-i="${i}"></td><td class="sub-file"></td><td class="num">${n}</td>
      <td class="act"><button class="ord" data-up="${i}" title="Mover para cima" aria-label="Mover para cima">↑</button>
      <button class="ord" data-down="${i}" title="Mover para baixo" aria-label="Mover para baixo">↓</button>
      <button class="btn-link" data-rm="${i}">Remover</button></td>`;
    tr.querySelector('input').value = p.label;
    tr.querySelector('.sub-file').textContent = p.file;
    tb.appendChild(tr);
  });
  $('cpListWrap').hidden = $('cpClear').hidden = !cpPeriods.length;
}

function cpRenderEmpresas(){
  const sel = $('cpEmp'), atual = sel.value;
  const map = new Map();
  cpPeriods.forEach(p => p.empresas.forEach((e, k) => { if (!map.has(k)) map.set(k, e.name); }));
  const opts = [...map.entries()].sort((a, b) => a[1].localeCompare(b[1], 'pt-BR'));
  sel.innerHTML = `<option value="${CP_ALL}">Todas as empresas (soma)</option>` + opts.map(([k]) => `<option value="${k}"></option>`).join('');
  opts.forEach(([k, n], i) => sel.options[i + 1].textContent = n);
  sel.value = map.has(atual) || atual === CP_ALL ? atual : (opts.length === 1 ? opts[0][0] : CP_ALL);
}

function cpData(key){
  return cpPeriods.map(p => {
    if (key === CP_ALL) return [...p.empresas.values()].reduce((a, e) => ({qtd: a.qtd + e.qtd, pontos: a.pontos + e.pontos, valor: a.valor + e.valor, nic: a.nic + e.nic}), {qtd: 0, pontos: 0, valor: 0, nic: 0});
    return p.empresas.get(key) || {qtd: 0, pontos: 0, valor: 0, nic: 0};
  });
}

const cpEsc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function cpChart(title, values, fmt, short){
  const n = values.length, W = 360, H = 190, padT = 24, padB = 34, padX = 12;
  const max = Math.max(...values, 0) || 1;
  const slot = (W - padX * 2) / n, bw = Math.min(64, slot * 0.6);
  const last = values[n - 1], prev = values[n - 2];
  let delta = '';
  if (n >= 2) {
    const d = prev ? (last - prev) / prev * 100 : (last ? 100 : 0);
    const cls = Math.abs(d) < 0.5 ? 'eq' : d > 0 ? 'up' : 'down';
    delta = `<span class="cp-delta ${cls}" title="Em relação a ${cpEsc(cpPeriods[n - 2].label)}">${cls === 'eq' ? '= 0%' : (d > 0 ? '▲ +' : '▼ ') + d.toLocaleString('pt-BR', {maximumFractionDigits: 1}) + '%'}</span>`;
  }
  const bars = values.map((v, i) => {
    const h = (H - padT - padB) * v / max, x = padX + slot * i + (slot - bw) / 2, y = H - padB - h;
    const isLast = i === n - 1;
    const color = isLast ? '#1E4E8C' : '#9FB6D8';
    return `<g class="cp-bar"><title>${cpEsc(cpPeriods[i].label)}: ${cpEsc(fmt(v))}</title>
      <rect class="cp-grow" style="animation-delay:${i * 80}ms" x="${x}" y="${y}" width="${bw}" height="${Math.max(h, v ? 2 : 0)}" rx="3" fill="${color}"/>
      <text x="${x + bw / 2}" y="${y - 6}" text-anchor="middle" font-size="11" font-weight="600" fill="#1F2937">${cpEsc(short(v))}</text>
      <text x="${x + bw / 2}" y="${H - padB + 16}" text-anchor="middle" font-size="11" fill="#6B7280">${cpEsc(cpPeriods[i].label)}</text></g>`;
  }).join('');
  return `<div class="cp-card"><h3>${title}</h3><div><span class="cp-big">${cpEsc(fmt(last))}</span>${delta}</div>
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${cpEsc(title)}">
      <line x1="${padX}" x2="${W - padX}" y1="${H - padB}" y2="${H - padB}" stroke="#E2E5E9"/>${bars}</svg></div>`;
}

function cpRender(){
  cpRenderList();
  const ok = cpPeriods.length > 0;
  $('cpPanel').hidden = !ok;
  if (!ok) return;
  cpRenderEmpresas();
  const key = $('cpEmp').value, d = cpData(key);
  const int = v => v.toLocaleString('pt-BR');
  const kbrl = v => v >= 1e6 ? 'R$ ' + (v / 1e6).toLocaleString('pt-BR', {maximumFractionDigits: 1}) + ' mi' : v >= 1e4 ? 'R$ ' + (v / 1e3).toLocaleString('pt-BR', {maximumFractionDigits: 1}) + ' mil' : brl(v);
  const nome = key === CP_ALL ? 'Todas as empresas' : $('cpEmp').selectedOptions[0].textContent;
  $('cpCharts').innerHTML =
    `<div class="print-head"><img src="logo.png" alt=""><div><strong>${cpEsc(nome)}</strong><br><small>Comparativo: ${cpEsc(cpPeriods.map(p => p.label).join(' × '))}</small></div></div>` +
    cpChart('Valor total', d.map(x => x.valor), brl, kbrl) +
    cpChart('Infrações', d.map(x => x.qtd), int, int) +
    cpChart('Pontos', d.map(x => x.pontos), int, int) +
    cpChart('Não identificação do condutor', d.map(x => x.nic), int, int);
  if (cpPeriods.length < 2) $('cpCharts').insertAdjacentHTML('afterbegin', '<div class="msg warn" style="grid-column:1/-1">Carregue pelo menos mais um PDF (outro mês) para comparar.</div>');
}

// ---------- eventos ----------
$('cpFile').addEventListener('change', e => { if (e.target.files.length) cpLoad([...e.target.files]); e.target.value = ''; });
const cpDrop = $('cpDrop');
['dragenter','dragover'].forEach(ev => cpDrop.addEventListener(ev, e => { e.preventDefault(); cpDrop.classList.add('over'); }));
['dragleave','drop'].forEach(ev => cpDrop.addEventListener(ev, e => { e.preventDefault(); cpDrop.classList.remove('over'); }));
cpDrop.addEventListener('drop', e => e.dataTransfer.files.length && cpLoad([...e.dataTransfer.files]));
$('cpList').addEventListener('input', e => { const i = e.target.dataset.i; if (i != null) { cpPeriods[i].label = e.target.value; clearTimeout(cpList._t); cpList._t = setTimeout(() => { const f = document.activeElement; const idx = f?.dataset?.i; cpRender(); if (idx != null) { const el = $('cpList').querySelector(`[data-i="${idx}"]`); el.focus(); el.setSelectionRange(el.value.length, el.value.length); } }, 400); } });
$('cpList').addEventListener('click', e => {
  const d = e.target.dataset, swap = (a, b) => { if (b < 0 || b >= cpPeriods.length) return; [cpPeriods[a], cpPeriods[b]] = [cpPeriods[b], cpPeriods[a]]; cpRender(); };
  if (d.up != null) swap(+d.up, +d.up - 1);
  else if (d.down != null) swap(+d.down, +d.down + 1);
  else if (d.rm != null) { cpPeriods.splice(+d.rm, 1); cpRender(); }
});
$('cpEmp').addEventListener('change', cpRender);
$('cpPrint').addEventListener('click', () => window.print());
$('cpClear').addEventListener('click', () => { cpPeriods = []; $('cpMsg').hidden = true; cpRender(); });
