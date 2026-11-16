// ===================== Word -> PDF =====================
// Renderizador enxuto para o modelo da Comunicação Administrativa: lê o .docx já preenchido
// (parágrafos, negrito, sublinhado, listas, alinhamento, quebra de página, timbre, linha da assinatura)
// e desenha o mesmo documento com o jsPDF. Assim Word e PDF saem sempre do mesmo modelo.

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const TW = 1 / 1440 * 25.4;            // twip -> mm
const PT = 25.4 / 72;                  // pt -> mm

function wEls(node, name){ return node ? [...node.getElementsByTagNameNS(W_NS, name)] : []; }
function wChild(node, name){ return node ? [...node.childNodes].find(n => n.localName === name && n.namespaceURI === W_NS) : null; }
function wAttr(node, name){ return node ? node.getAttributeNS(W_NS, name) || node.getAttribute('w:' + name) : null; }
function isOn(rPr, name){ const e = wChild(rPr, name); return !!e && !['0','false','none'].includes(wAttr(e, 'val') || ''); }

async function docxParaPdf(zip){
  const parse = async f => zip.file(f) ? new DOMParser().parseFromString(await zip.file(f).async('string'), 'application/xml') : null;
  const doc = await parse('word/document.xml');
  const numXml = await parse('word/numbering.xml');

  // numId -> formato da lista (bullet / decimal)
  const absFmt = {};
  wEls(numXml, 'abstractNum').forEach(a => {
    const lvl = wEls(a, 'lvl').find(l => wAttr(l, 'ilvl') === '0');
    absFmt[wAttr(a, 'abstractNumId')] = wAttr(wChild(lvl, 'numFmt'), 'val') || 'bullet';
  });
  const numFmt = {};
  wEls(numXml, 'num').forEach(n => numFmt[wAttr(n, 'numId')] = absFmt[wAttr(wChild(n, 'abstractNumId'), 'val')]);

  // página e margens
  const sect = wEls(doc, 'sectPr').pop();
  const pg = wChild(sect, 'pgSz'), mar = wChild(sect, 'pgMar');
  const PW = +wAttr(pg, 'w') * TW, PH = +wAttr(pg, 'h') * TW;
  const ML = +wAttr(mar, 'left') * TW, MR = +wAttr(mar, 'right') * TW, MT = +wAttr(mar, 'top') * TW, MB = +wAttr(mar, 'bottom') * TW;
  const CW = PW - ML - MR;

  // timbre (imagem do cabeçalho)
  let timbre = null;
  const hdrRels = zip.file('word/_rels/header1.xml.rels') && await zip.file('word/_rels/header1.xml.rels').async('string');
  const m = hdrRels && /Target="(media\/[^"]+\.png)"/.exec(hdrRels);
  if (m) timbre = 'data:image/png;base64,' + await zip.file('word/' + m[1]).async('base64');

  const FS = 11, LH = FS * PT * 1.08 * 1.17, AFTER = 160 * TW;   // Calibri 11, espaçamento 1,08, 8 pt depois
  const pdf = new jspdf.jsPDF({unit: 'mm', format: [PW, PH]});
  pdf.setFont('helvetica', 'normal'); pdf.setFontSize(FS); pdf.setTextColor(0, 0, 0);
  const setStyle = s => pdf.setFont('helvetica', s.b ? 'bold' : 'normal');
  const width = (t, s) => { setStyle(s); return pdf.getTextWidth(t); };

  let y = MT, page = 1;
  const decorate = () => { if (timbre) pdf.addImage(timbre, 'PNG', 0, 0, PW, PW * 295 / 2480); };
  const newPage = () => { pdf.addPage([PW, PH]); page++; decorate(); y = MT; };
  decorate();

  const counters = {};
  const body = wChild(doc.documentElement, 'body');
  for (const p of [...body.childNodes].filter(n => n.localName === 'p')) {
    const pPr = wChild(p, 'pPr');
    const jc = wAttr(wChild(pPr, 'jc'), 'val') || 'left';
    const numPr = wChild(pPr, 'numPr');
    if (wChild(pPr, 'pageBreakBefore') && y > MT + 0.1) newPage();

    // runs -> tokens (palavras com estilo, espaços e quebras)
    const tokens = []; let hasDrawing = false;
    for (const r of wEls(p, 'r')) {
      const rPr = wChild(r, 'rPr');
      const s = {b: isOn(rPr, 'b'), u: !!wChild(rPr, 'u') && wAttr(wChild(rPr, 'u'), 'val') !== 'none', caps: isOn(rPr, 'caps')};
      for (const c of r.childNodes) {
        if (c.localName === 't') {
          let t = c.textContent; if (s.caps) t = t.toUpperCase();
          t.split(/(\s+)/).forEach(part => { if (part) tokens.push({t: part, s, sp: /^\s+$/.test(part)}); });
        } else if (c.localName === 'br') tokens.push({br: true});
        else if (c.localName === 'tab') tokens.push({t: ' ', s, sp: true});
        else if (c.localName === 'drawing' || c.localName === 'pict') hasDrawing = true;
      }
    }

    // recuo e marcador da lista
    let indent = 0, marker = null;
    if (numPr) {
      const id = wAttr(wChild(numPr, 'numId'), 'val');
      indent = 720 * TW;
      counters[id] = (counters[id] || 0) + 1;
      marker = numFmt[id] === 'decimal' ? counters[id] + '.' : '•';
    }
    const x0 = ML + indent, maxW = CW - indent;

    if (!tokens.some(t => !t.sp && !t.br)) {          // parágrafo vazio (ou só a linha da assinatura)
      if (y + LH > PH - MB) newPage();
      if (hasDrawing) { pdf.setDrawColor(140); pdf.setLineWidth(0.25); pdf.line(ML, y + LH * 0.5, ML + CW, y + LH * 0.5); }
      y += LH + AFTER; continue;
    }

    // quebra em linhas
    const lines = []; let cur = [], curW = 0;
    const flush = hard => { while (cur.length && cur[cur.length - 1].sp) cur.pop(); lines.push({items: cur, hard}); cur = []; curW = 0; };
    for (const tk of tokens) {
      if (tk.br) { flush(true); continue; }
      if (tk.sp && !cur.length) continue;
      const w = width(tk.t, tk.s);
      if (!tk.sp && curW + w > maxW && cur.length) flush(false);
      if (tk.sp && !cur.length) continue;
      cur.push({...tk, w}); curW += w;
    }
    flush(true);

    lines.forEach((ln, li) => {
      if (y + LH > PH - MB) newPage();
      const base = y + FS * PT * 0.8;
      if (li === 0 && marker) { setStyle({}); pdf.text(marker, x0 - (marker === '•' ? 4.5 : 5.5), base); }
      const natural = ln.items.reduce((a, it) => a + it.w, 0);
      const spaces = ln.items.filter(it => it.sp).length;
      let x = x0, extra = 0;
      if (jc === 'center') x = ML + (CW - natural) / 2;
      else if (jc === 'right') x = ML + CW - natural;
      else if (jc === 'both' && !ln.hard && spaces) extra = (maxW - natural) / spaces;
      for (const it of ln.items) {
        const w = it.w + (it.sp ? extra : 0);
        if (!it.sp) { setStyle(it.s); pdf.text(it.t, x, base); }
        if (it.s && it.s.u) { pdf.setDrawColor(0); pdf.setLineWidth(0.2); pdf.line(x, base + 0.6, x + w, base + 0.6); }
        x += w;
      }
      y += LH;
    });
    y += AFTER;
  }

  // número da página no rodapé (como no Word)
  const n = pdf.getNumberOfPages();
  for (let i = 1; i <= n; i++) { pdf.setPage(i); pdf.setFont('helvetica', 'normal'); pdf.setFontSize(9); pdf.setTextColor(90); pdf.text(String(i), PW - MR, PH - 12, {align: 'right'}); }
  return pdf.output('blob');
}
