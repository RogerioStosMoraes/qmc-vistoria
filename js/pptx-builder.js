// Gera o relatório .pptx a partir do modelo QMC (template.pptx enxuto, sem slides).
// Funciona no navegador e no Node (JSZip é injetado).
import { STATUS, YN, normAnswer, expandGroup, scoreGroup, hasContent, instTitle, OPERADORAS, fmtDate, fmtTime } from './model.js';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Geometria do modelo (EMU)
const FRAME = { x: 293238, y: 809008, w: 9607639, h: 4677391 };
const PHOTO_BOX = { w: 3244510, h: 2453969, y: 923132, xs: [1002549, 5703138] };
const OBS_BOX = { w: 3244510, h: 1802294, y: 3491225 };

const FONT = 'Encode Sans';
const FONT_XB = 'Encode Sans ExtraBold';
const TEAL = '0E5676';
const GREY = '363738';

export function esc(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ---------- helpers XML ----------
class Slide {
  constructor(layout) {
    this.layout = layout; // 12 | 13
    this.shapes = [];
    this.rels = []; // {id,type,target,external}
    this.nid = 2;
  }
  id() { return this.nid++; }
  rel(type, target, external = false) {
    const id = 'rId' + (this.rels.length + 2); // rId1 = layout
    this.rels.push({ id, type, target, external });
    return id;
  }
  xml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${this.shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
  }
  relsXml() {
    const r = [`<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout${this.layout}.xml"/>`];
    for (const x of this.rels) {
      r.push(`<Relationship Id="${x.id}" Type="${x.type}" Target="${esc(x.target)}"${x.external ? ' TargetMode="External"' : ''}/>`);
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${r.join('')}</Relationships>`;
  }
}

function rPr({ sz, b, color, font, link }) {
  const fill = color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : '';
  const lat = font ? `<a:latin typeface="${font}"/><a:ea typeface="${font}"/><a:cs typeface="${font}"/>` : '';
  const hl = link ? `<a:hlinkClick r:id="${link}"/>` : '';
  return `<a:rPr lang="pt-BR"${sz ? ` sz="${sz}"` : ''}${b !== undefined ? ` b="${b ? 1 : 0}"` : ''} dirty="0">${fill}${lat}${hl}</a:rPr>`;
}
function run(t, o = {}) { return `<a:r>${rPr(o)}<a:t>${esc(t)}</a:t></a:r>`; }
function para(runs, { algn = 'l', spcAft, spcBef } = {}) {
  const sp = (spcBef != null ? `<a:spcBef><a:spcPts val="${spcBef}"/></a:spcBef>` : '') + (spcAft != null ? `<a:spcAft><a:spcPts val="${spcAft}"/></a:spcAft>` : '');
  return `<a:p><a:pPr algn="${algn}">${sp}</a:pPr>${runs.join('')}</a:p>`;
}
function textBox(s, { x, y, w, h, paras, name = 'Texto', anchor = 't', ins = [75600, 37775, 75600, 37775] }) {
  const id = s.id();
  s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)} ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="${ins[0]}" tIns="${ins[1]}" rIns="${ins[2]}" bIns="${ins[3]}" rtlCol="0" anchor="${anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>${paras.join('')}</p:txBody></p:sp>`);
}
function titleShape(s, text, { x = 293238, y = 118617, w = 8100000, h = 548640, sz } = {}) {
  const id = s.id();
  const n = text.length;
  const size = sz || (n <= 36 ? 2600 : n <= 46 ? 2200 : 1800);
  s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="pt-BR" sz="${size}" dirty="0"/><a:t>${esc(text)}</a:t></a:r></a:p></p:txBody></p:sp>`);
}
function frameShapes(s) {
  const a = s.id(), b = s.id();
  s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${a}" name="Moldura ${a}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="179748" y="741488"/><a:ext cx="9777052" cy="4810445"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent4"><a:alpha val="20000"/></a:schemeClr></a:solidFill><a:ln w="63500"><a:solidFill><a:srgbClr val="FDBD52"/></a:solidFill></a:ln><a:effectLst><a:outerShdw blurRad="177800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="pt-BR" sz="772" dirty="0"/></a:p></p:txBody></p:sp>`);
  s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${b}" name="Fundo ${b}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${FRAME.x}" y="${FRAME.y}"/><a:ext cx="${FRAME.w}" cy="${FRAME.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="pt-BR" dirty="0"/></a:p></p:txBody></p:sp>`);
}
function picShape(s, relId, { x, y, w, h, name = 'Imagem', descr = '', line = false }) {
  const id = s.id();
  const ln = line ? '<a:ln w="9525"><a:solidFill><a:srgbClr val="D0D0D0"/></a:solidFill></a:ln>' : '';
  s.shapes.push(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)} ${id}" descr="${esc(descr)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:pic>`);
}
function fit(boxW, boxH, w, h) {
  const r = Math.min(boxW / w, boxH / h);
  return { w: w * r, h: h * r };
}

// ---------- tabela ----------
// níveis de densidade: o maior que couber em um slide é escolhido
const DENS = [
  { sz: 1000, rowH: 250000, extraH: 150000, cpl: 48 },
  { sz: 900, rowH: 212000, extraH: 135000, cpl: 54 },
  { sz: 800, rowH: 176000, extraH: 118000, cpl: 60 },
];
const TABLE_H = 3570000;

function cell(text, { fill, color = GREY, b = false, algn = 'l', sz = 800, bottom = 'E3E3E3' } = {}) {
  const ln = `<a:lnL w="0"><a:noFill/></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnT w="0"><a:noFill/></a:lnT><a:lnB w="6350"><a:solidFill><a:srgbClr val="${bottom}"/></a:solidFill></a:lnB>`;
  const f = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '<a:noFill/>';
  const p = text === '' ? `<a:p><a:pPr algn="${algn}"/><a:endParaRPr lang="pt-BR" sz="${sz}" dirty="0"/></a:p>` : `<a:p><a:pPr algn="${algn}"/>${run(text, { sz, b, color, font: FONT })}</a:p>`;
  return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${p}</a:txBody><a:tcPr marL="54000" marR="54000" marT="14000" marB="14000" anchor="ctr">${ln}${f}</a:tcPr></a:tc>`;
}
function rowHeight(text, d) {
  const lines = Math.max(1, Math.ceil(String(text).length / d.cpl));
  return d.rowH + (lines - 1) * d.extraH;
}
function tableShape(s, { x, y, colW, rows }) {
  const id = s.id();
  const total = rows.reduce((a, r) => a + r.h, 0);
  const cx = colW.reduce((a, b) => a + b, 0);
  const grid = colW.map((w) => `<a:gridCol w="${w}"/>`).join('');
  const trs = rows.map((r) => `<a:tr h="${r.h}">${r.cells.join('')}</a:tr>`).join('');
  s.shapes.push(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Tabela ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${total}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
}

/** Linhas do checklist de uma instância (item, extra, seção de valores). */
function buildRows(def, inst) {
  const rows = [];
  for (const it of def.items) {
    const a = it.tipo === 'yn' ? inst.answers?.[it.key] : normAnswer(inst.answers?.[it.key]);
    if (it.tipo === 'yn') {
      const st = a ? YN[a] : null;
      rows.push({ kind: 'item', text: it.text, st: st ? { ...st } : null });
    } else {
      rows.push({ kind: 'item', text: it.text, st: a && STATUS[a] ? STATUS[a] : null });
      if (it.extra) {
        const e = inst.answers?.[it.key + '#extra'];
        if (e) rows.push({ kind: 'item', text: '   ↳ ' + it.extra, st: YN[e] });
      }
    }
  }
  const vals = [];
  for (const c of def.campos) {
    const v = inst.values?.[c.key];
    if (v && String(v.v ?? '').trim() !== '') vals.push({ c, v });
  }
  if (vals.length) {
    if (def.items.length) rows.push({ kind: 'sub', text: def.tipo === 'medicao' ? 'Medições' : 'Amperagens dos disjuntores' });
    for (const { c, v } of vals) {
      const val = String(v.v).trim() + (c.unit ? ' ' + c.unit : '');
      rows.push({ kind: 'val', text: c.label, value: val, diff: !!v.diff });
    }
  }
  return rows;
}

function rowToTr(r, d) {
  const sz = d.sz;
  if (r.kind === 'sub') {
    return { h: d.rowH, cells: [cell(r.text, { sz, fill: 'F3F3F3', b: true, color: TEAL }), cell('', { sz, fill: 'F3F3F3' })], kind: 'sub' };
  }
  if (r.kind === 'val') {
    const t = r.diff ? `${r.value} (≠ projeto)` : r.value;
    return { h: rowHeight(r.text, d), cells: [cell(r.text, { sz }), cell(t, { sz, b: true, algn: 'ctr', color: r.diff ? 'B3261E' : GREY, fill: r.diff ? 'FBDADA' : undefined })], kind: 'val' };
  }
  const st = r.st;
  const sc = st
    ? cell(st.label, { sz, fill: st.fill, color: st.color, b: true, algn: 'ctr' })
    : cell('Pendente', { sz, color: '939598', algn: 'ctr' });
  return { h: rowHeight(r.text, d), cells: [cell(r.text, { sz }), sc], kind: 'item' };
}

// ---------- slides ----------
function headerStrip(s, def, inst, sc) {
  let t;
  if (def.items.some((i) => i.tipo === 'status')) {
    const pct = sc.pct == null ? '—' : sc.pct + '%';
    t = [
      run('Conformidade: ', { sz: 900, b: true, color: TEAL, font: FONT }),
      run(`${pct}`, { sz: 900, b: true, color: GREY, font: FONT }),
      run(`   (${sc.aplicaveis} itens aplicáveis)    OK ${sc.ok}  ·  Não conf. ${sc.nok}  ·  N/A ${sc.na}${sc.pend ? `  ·  Pendentes ${sc.pend}` : ''}`, { sz: 900, color: GREY, font: FONT }),
    ];
  } else {
    t = [run(def.tipo === 'medicao' ? 'Medições elétricas registradas em campo' : 'Valores registrados em campo', { sz: 900, b: true, color: TEAL, font: FONT })];
  }
  textBox(s, { x: FRAME.x + 100000, y: FRAME.y + 40000, w: FRAME.w - 200000, h: 260000, paras: [para(t)], name: 'Resumo', anchor: 'ctr' });
}

function obsBox(s, text) {
  const lines = String(text || '').trim() ? String(text).trim().split(/\r?\n/) : ['—'];
  const paras = [para([run('Observação:', { sz: 1100, color: TEAL, font: FONT_XB })], { spcAft: 100 })];
  const sz = text && text.length > 420 ? 800 : 900;
  for (const l of lines) paras.push(para([run(l, { sz, b: true, color: GREY, font: FONT })]));
  textBox(s, { x: FRAME.x + 100000, y: 4740000, w: FRAME.w - 200000, h: 700000, paras, name: 'Observação do grupo' });
}

function checklistSlides(def, inst) {
  const sc = scoreGroup(def, inst);
  const rawRows = buildRows(def, inst);
  let d = DENS[DENS.length - 1], trs = rawRows.map((r) => rowToTr(r, d));
  for (const cand of DENS) {
    const t = rawRows.map((r) => rowToTr(r, cand));
    const tot = t.reduce((a, r) => a + r.h, 0);
    const mx = Math.max(...t.map((r) => r.h), 0);
    // cabe em uma página, dividido em duas colunas (ou uma, se poucas linhas)?
    const cap = TABLE_H - cand.rowH;
    const need = t.length < 6 ? tot : tot / 2 + mx * 1.5;
    if (need <= cap) { d = cand; trs = t; break; }
  }
  const ROW_H = d.rowH;
  const colW = [3500000, 1143819];
  const gap = 120000;
  const x0 = FRAME.x + 100000;
  const y0 = 1130000;
  const capacity = TABLE_H - ROW_H; // altura útil por coluna, sem o cabeçalho
  const head = () => ({ h: ROW_H, cells: [cell('Item', { sz: d.sz, fill: TEAL, color: 'FFFFFF', b: true }), cell('Situação', { sz: d.sz, fill: TEAL, color: 'FFFFFF', b: true, algn: 'ctr' })] });

  const pages = [];
  let page = [[], []], used = [0, 0], col = 0;
  const newPage = () => { pages.push(page); page = [[], []]; used = [0, 0]; col = 0; };
  const T = trs.reduce((a, r) => a + r.h, 0);
  const pagesN = Math.max(1, Math.ceil(T / (2 * capacity)));
  const target = pagesN === 1 && trs.length < 6 ? capacity : Math.min(capacity, T / (pagesN * 2) + ROW_H * 0.6);
  for (let i = 0; i < trs.length; i++) {
    const tr = trs[i];
    const need = tr.kind === 'sub' && trs[i + 1] ? tr.h + trs[i + 1].h : tr.h;
    if (used[col] > 0 && used[col] + need > (col === 0 ? target : capacity)) {
      if (col === 0) col = 1; else newPage();
      if (col === 1 && used[1] > 0 && used[1] + need > capacity) newPage();
    }
    page[col].push(tr); used[col] += tr.h;
  }
  if (page[0].length || page[1].length || !pages.length) pages.push(page);

  const slides = [];
  const base = 'Check-List ' + instTitle(def, inst);
  pages.forEach((pg, pi) => {
    const s = new Slide(13);
    const t = pages.length > 1 ? `${base} (${pi + 1}/${pages.length})` : base;
    titleShape(s, t);
    frameShapes(s);
    headerStrip(s, def, inst, sc);
    pg.forEach((colRows, ci) => {
      if (!colRows.length) return;
      tableShape(s, { x: x0 + ci * (colW[0] + colW[1] + gap), y: y0, colW, rows: [head(), ...colRows] });
    });
    if (pi === pages.length - 1) obsBox(s, inst.obs);
    slides.push(s);
  });
  return slides;
}

function evidenceSlides(def, inst, ctx) {
  const photos = inst.photos || [];
  const out = [];
  const pages = Math.ceil(photos.length / 2);
  for (let p = 0; p < pages; p++) {
    const s = new Slide(13);
    const base = 'Evidências - ' + instTitle(def, inst);
    titleShape(s, pages > 1 ? `${base} (${p + 1}/${pages})` : base);
    frameShapes(s);
    const pagePhotos = photos.slice(p * 2, p * 2 + 2);
    pagePhotos.forEach((ph, i) => {
      const x = pagePhotos.length === 1 ? FRAME.x + (FRAME.w - PHOTO_BOX.w) / 2 : PHOTO_BOX.xs[i];
      const data = ctx.photoBytes.get(ph.id);
      if (data) {
        const name = ctx.addMedia(data, 'jpg');
        const rid = s.rel(`${REL}/image`, `../media/${name}`);
        const f = fit(PHOTO_BOX.w, PHOTO_BOX.h, ph.w || 4, ph.h || 3);
        picShape(s, rid, { x: x + (PHOTO_BOX.w - f.w) / 2, y: PHOTO_BOX.y + (PHOTO_BOX.h - f.h) / 2, w: f.w, h: f.h, name: 'Foto', descr: ph.obs || '', line: true });
      }
      const obs = String(ph.obs || '').trim();
      const lines = obs ? obs.split(/\r?\n/) : ['—'];
      const sz = obs.length > 380 ? 800 : obs.length > 220 ? 900 : 1000;
      const paras = [para([run('Observação:', { sz: 1400, color: TEAL, font: FONT_XB })])];
      for (const l of lines) paras.push(para([run(l, { sz, b: true, color: GREY, font: FONT })]));
      textBox(s, { x, y: OBS_BOX.y, w: OBS_BOX.w, h: OBS_BOX.h, paras, name: 'Observação da foto' });
    });
    out.push(s);
  }
  return out;
}

function coverSlide(v, ctx) {
  const s = new Slide(12);
  const h = v.header || {};
  // gradiente escuro à esquerda (como no modelo)
  const gid = s.id();
  s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${gid}" name="Overlay Gradiente Capa"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6249987" cy="5670550"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="051024"><a:alpha val="78000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="051024"><a:alpha val="0"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill><a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="pt-BR"/></a:p></p:txBody></p:sp>`);

  const name = (h.empreendimento || '').trim() || 'Empreendimento';
  const title = `Vistoria – ${name}`;
  const tsz = title.length <= 28 ? 2600 : title.length <= 40 ? 2200 : 1800;
  const l2 = [];
  if (h.siteId) l2.push(h.siteId);
  const proj = [h.projeto && `Projeto ${h.projeto}`, h.implantadora].filter(Boolean).join(' · ');
  if (proj) l2.push(proj);
  const when = [fmtDate(v.startAt), v.startAt && v.endAt ? `${fmtTime(v.startAt)} – ${fmtTime(v.endAt)}` : fmtTime(v.startAt)].filter(Boolean).join(' · ');
  if (when) l2.push(when);
  const tid = s.id();
  const sub = [];
  l2.forEach((t, i) => {
    sub.push(run(t, { sz: 1600, b: false, color: 'F58220' }));
    if (i < l2.length - 1) sub.push('<a:br><a:rPr lang="pt-BR" sz="1600"/></a:br>');
  });
  s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${tid}" name="Título ${tid}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="554854" y="695331"/><a:ext cx="6784800" cy="1650000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="pt-BR" sz="${tsz}" dirty="0"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:rPr><a:t>${esc(title)}</a:t></a:r></a:p><a:p><a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>${sub.join('')}</a:p></p:txBody></p:sp>`);

  // logos das operadoras sobre placas brancas
  const plateW = 1750000, plateH = 760000, gap = 150000, x0 = 554854, y0 = 2450000;
  let i = 0;
  for (const key of h.operadoras || []) {
    const lg = ctx.logos?.[key];
    if (!lg) continue;
    const px = x0 + i * (plateW + gap);
    const pid = s.id();
    s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${pid}" name="Placa ${key}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${px}" y="${y0}"/><a:ext cx="${plateW}" cy="${plateH}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 14000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="92000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="pt-BR"/></a:p></p:txBody></p:sp>`);
    const name = ctx.addMedia(lg.bytes, 'png');
    const rid = s.rel(`${REL}/image`, `../media/${name}`);
    const f = fit(plateW - 220000, plateH - 170000, lg.w, lg.h);
    picShape(s, rid, { x: px + (plateW - f.w) / 2, y: y0 + (plateH - f.h) / 2, w: f.w, h: f.h, name: 'Logo ' + (OPERADORAS[key]?.nome || key), descr: 'Logo ' + (OPERADORAS[key]?.nome || key) });
    i++;
  }
  if ((h.outraOperadora || '').trim()) {
    const px = x0 + i * (plateW + gap);
    const pid = s.id();
    s.shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${pid}" name="Placa outra"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${px}" y="${y0}"/><a:ext cx="${plateW}" cy="${plateH}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 14000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="92000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="pt-BR"/></a:p></p:txBody></p:sp>`);
    textBox(s, { x: px, y: y0, w: plateW, h: plateH, paras: [para([run(h.outraOperadora.trim(), { sz: 1400, b: true, color: '1A1A1A', font: FONT })], { algn: 'ctr' })], name: 'Outra operadora', anchor: 'ctr' });
  }

  // GPS (link para o mapa)
  if (v.gps && v.gps.lat != null) {
    const rid = s.rel(`${REL}/hyperlink`, `https://www.google.com/maps?q=${v.gps.lat},${v.gps.lon}`, true);
    const txt = `GPS: ${Number(v.gps.lat).toFixed(5)}, ${Number(v.gps.lon).toFixed(5)}`;
    textBox(s, { x: 554854, y: 5180000, w: 5200000, h: 260000, paras: [para([run(txt, { sz: 900, color: 'FFFFFF', font: FONT, link: rid })])], name: 'GPS', anchor: 'ctr' });
  }
  return s;
}

// ---------- montagem do pacote ----------
export async function buildPptx({ JSZip, template, vistoria, checklist, photoBytes = new Map(), logos = {}, coverBytes = null }) {
  const zip = await JSZip.loadAsync(template);
  const defs = new Map(checklist.grupos.map((g) => [g.id, expandGroup(checklist, g)]));
  let mediaN = 0;
  const media = [];
  const ctx = {
    photoBytes, logos,
    addMedia(bytes, ext) { const name = `rep${++mediaN}.${ext}`; media.push([name, bytes]); return name; },
  };

  const slides = [coverSlide(vistoria, ctx)];
  // ordem: a do checklist.json, instâncias na ordem de criação
  const order = checklist.grupos.map((g) => g.id);
  const insts = [...(vistoria.groups || [])].sort((a, b) => order.indexOf(a.groupId) - order.indexOf(b.groupId));
  for (const inst of insts) {
    const def = defs.get(inst.groupId);
    if (!def || !hasContent(def, inst)) continue;
    slides.push(...checklistSlides(def, inst));
    slides.push(...evidenceSlides(def, inst, ctx));
  }

  if (coverBytes) zip.file('ppt/media/image20.jpg', coverBytes);
  for (const [name, bytes] of media) zip.file('ppt/media/' + name, bytes);

  let ct = await zip.file('[Content_Types].xml').async('string');
  let pres = await zip.file('ppt/presentation.xml').async('string');
  let prels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  const ids = [], newRels = [], overrides = [];
  slides.forEach((s, i) => {
    const n = i + 1;
    zip.file(`ppt/slides/slide${n}.xml`, s.xml());
    zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, s.relsXml());
    overrides.push(`<Override PartName="/ppt/slides/slide${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
    newRels.push(`<Relationship Id="rId${1000 + n}" Type="${REL}/slide" Target="slides/slide${n}.xml"/>`);
    ids.push(`<p:sldId id="${255 + n}" r:id="rId${1000 + n}"/>`);
  });
  ct = ct.replace('</Types>', overrides.join('') + '</Types>');
  pres = pres.replace('<p:sldIdLst/>', `<p:sldIdLst>${ids.join('')}</p:sldIdLst>`);
  prels = prels.replace('</Relationships>', newRels.join('') + '</Relationships>');
  zip.file('[Content_Types].xml', ct);
  zip.file('ppt/presentation.xml', pres);
  zip.file('ppt/_rels/presentation.xml.rels', prels);

  const core = zip.file('docProps/core.xml');
  if (core) {
    let c = await core.async('string');
    const title = `Vistoria – ${vistoria.header?.empreendimento || ''} ${vistoria.header?.siteId || ''}`.trim();
    if (/<dc:title>.*?<\/dc:title>|<dc:title\/>/s.test(c)) c = c.replace(/<dc:title>.*?<\/dc:title>|<dc:title\/>/s, `<dc:title>${esc(title)}</dc:title>`);
    zip.file('docProps/core.xml', c);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 4 } });
}
