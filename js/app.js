import { STATUS, STATUS_ORDER, YN, expandAll, scoreGroup, hasContent, needsPhoto, normAnswer, instTitle, OPERADORAS, uid, fmtDate, fmtTime, pad } from './model.js';
import * as db from './db.js';
import * as cloud from './cloud.js';
import { buildPptx } from './pptx-builder.js';

const $ = (s, r = document) => r.querySelector(s);
const view = $('#view');
const E = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

let checklist, defs, defMap;
let V = null; // vistoria aberta
let deferredInstall = null;
const urls = new Map();

// ---------- utilidades ----------
let toastT;
function toast(msg, ms = 2800) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('show'), ms);
}

function askText(title, value = '') {
  const d = $('#dlg'); $('#dlgTitle').textContent = title; const i = $('#dlgInput'); i.value = value;
  return new Promise((resolve) => {
    d.addEventListener('close', () => resolve(d.returnValue === 'ok' ? i.value.trim() : null), { once: true });
    d.showModal(); setTimeout(() => { i.focus(); i.select(); }, 50);
  });
}

let saveT, syncT;
function touch() {
  if (!V) return;
  V.updatedAt = new Date().toISOString(); V.dirty = true;
  clearTimeout(saveT); saveT = setTimeout(() => db.putVistoria(V), 250);
  scheduleSync();
}
async function flush() { clearTimeout(saveT); if (V) await db.putVistoria(V); }
function scheduleSync() {
  if (!cloud.user() || !navigator.onLine) return;
  clearTimeout(syncT); syncT = setTimeout(() => doSync(false), 12000);
}
let syncing = false;
async function doSync(manual) {
  if (syncing || !cloud.user()) return;
  if (!navigator.onLine) { if (manual) toast('Sem internet — sincroniza quando voltar a conexão'); return; }
  syncing = true; setSyncBtn(true);
  try {
    await flush();
    const r = await cloud.syncAll();
    if (V) { const fresh = await db.getVistoria(V.id); if (fresh && !V.dirty) V = fresh; }
    if (manual || r.pulled) toast(`Sincronizado: ${r.pushed} enviada(s), ${r.pulled} recebida(s)`);
    if (location.hash === '' || location.hash === '#/') route();
  } catch (e) { console.error(e); if (manual) toast('Falha ao sincronizar: ' + e.message, 5000); }
  finally { syncing = false; setSyncBtn(false); }
}
function setSyncBtn(on) { const b = $('[data-act=sync]'); if (b) { b.disabled = on; b.textContent = on ? '⟳ …' : '⟳'; } }

function photoURL(id) {
  if (urls.has(id)) return Promise.resolve(urls.get(id));
  return db.getPhoto(id).then((r) => { if (!r?.blob) return ''; const u = URL.createObjectURL(r.blob); urls.set(id, u); return u; });
}

// ---------- imagens ----------
async function bitmapOf(file) {
  if (window.createImageBitmap) { try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); } catch (e) { /* fallback */ } }
  return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = URL.createObjectURL(file); });
}
async function toJpeg(file, maxDim = 1600, q = 0.8) {
  const bm = await bitmapOf(file);
  const w0 = bm.width || bm.naturalWidth, h0 = bm.height || bm.naturalHeight;
  const s = Math.min(1, maxDim / Math.max(w0, h0));
  const w = Math.round(w0 * s), h = Math.round(h0 * s);
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  c.getContext('2d').drawImage(bm, 0, 0, w, h);
  const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', q));
  if (bm.close) bm.close();
  return { blob, w, h };
}
async function coverJpeg(file) {
  const bm = await bitmapOf(file);
  const w0 = bm.width || bm.naturalWidth, h0 = bm.height || bm.naturalHeight;
  const W = 1600, H = 900, r = Math.max(W / w0, H / h0);
  const c = document.createElement('canvas'); c.width = W; c.height = H;
  const ctx = c.getContext('2d');
  ctx.drawImage(bm, (W - w0 * r) / 2, (H - h0 * r) / 2, w0 * r, h0 * r);
  const blob = await new Promise((res) => c.toBlob(res, 'image/jpeg', 0.85));
  if (bm.close) bm.close();
  return { blob, w: W, h: H };
}

function getGPS() {
  return new Promise((res) => {
    if (!navigator.geolocation) return res(null);
    navigator.geolocation.getCurrentPosition(
      (p) => res({ lat: p.coords.latitude, lon: p.coords.longitude, acc: Math.round(p.coords.accuracy), ts: new Date(p.timestamp).toISOString() }),
      () => res(null), { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 });
  });
}

// ---------- modelo ----------
const newInst = (d, label = '') => ({ gid: uid(), groupId: d.id, label, na: false, answers: {}, values: {}, obs: '', photos: [] });
const instsOf = (d) => V.groups.filter((g) => g.groupId === d.id);
const isUnique = (d) => d.instancias === 'única';
const DEFAULT_LABEL = { idf: 'IDF', miniqdca: 'IDF', n3ru: 'Unidade', gab_poi: 'Gab.', gab_ftta: 'Gab.', gab_5g: 'Gab.', hbts_op: 'Operadora', qdca_op: 'Operadora' };

function newVistoria() {
  const now = new Date().toISOString();
  const v = { id: uid(), createdAt: now, updatedAt: now, dirty: true, startAt: now, endAt: null, gps: null, cover: null, removedPhotos: [],
    header: { siteId: '', empreendimento: '', projeto: '', implantadora: '', operadoras: [], outraOperadora: '' }, groups: [] };
  for (const d of defs) if (isUnique(d)) v.groups.push(newInst(d));
  return v;
}

function opLabel(key) { return OPERADORAS[key]?.nome || key; }
function syncOperatorGroups() {
  const keys = V.header.operadoras.map(opLabel);
  for (const id of ['hbts_op', 'qdca_op']) {
    const d = defMap.get(id);
    for (const k of keys) if (!V.groups.some((g) => g.groupId === id && g.label === k)) V.groups.push(newInst(d, k));
    V.groups = V.groups.filter((g) => !(g.groupId === id && Object.values(OPERADORAS).some((o) => o.nome === g.label) && !keys.includes(g.label) && !hasContent(d, g)));
  }
}

function progressOf(d, inst) {
  const st = d.items.filter((i) => i.tipo === 'status' || i.tipo === 'yn');
  const done = st.filter((i) => inst.answers[i.key]).length;
  let total = st.length;
  if (!total) { // grupos só com campos numéricos
    total = d.campos.length; const dn = d.campos.filter((c) => String(inst.values[c.key]?.v ?? '').trim() !== '').length;
    return { done: dn, total, pct: total ? Math.round((dn / total) * 100) : 0 };
  }
  return { done, total, pct: Math.round((done / total) * 100) };
}
function overall() {
  let done = 0, total = 0;
  for (const g of V.groups) { if (g.na) continue; const p = progressOf(defMap.get(g.groupId), g); done += p.done; total += p.total; }
  return { done, total, pct: total ? Math.round((done / total) * 100) : 0 };
}

// ---------- layout (barra superior) ----------
function setTop({ back, title, right = '' }) {
  $('#top').innerHTML = (back ? `<a class="icon-btn" href="${back}" aria-label="Voltar" style="text-decoration:none;display:grid;place-items:center">←</a>` : `<img src="assets/qmc-logo.png" alt="QMC">`) + `<h1>${E(title)}</h1>${right}`;
}

// ---------- HOME ----------
async function viewHome() {
  V = null;
  const list = (await db.listVistorias()).sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || ''));
  const u = cloud.user();
  let acct = '';
  if (cloud.configured()) {
    acct = u
      ? `<span class="pill" title="${E(u.username)}">${E((u.name || u.username || '').split(' ')[0])}</span><button class="icon-btn" data-act="sync" aria-label="Sincronizar">⟳</button><button class="icon-btn" data-act="logout" aria-label="Sair" title="Sair">⎋</button>`
      : `<button class="btn primary small" data-act="login">Entrar com Microsoft</button>`;
  }
  setTop({ title: 'Vistorias', right: acct });
  const note = !cloud.configured()
    ? `<div class="warnbox">Modo local: o login Microsoft ainda não foi configurado (arquivo <b>config.js</b>). As vistorias ficam só neste aparelho.</div>`
    : (!u ? `<div class="warnbox">Entre com sua conta Microsoft para guardar as vistorias no seu OneDrive e abri-las em outros aparelhos.</div>` : '');
  view.innerHTML = `
    ${note}
    <div class="row" style="margin-bottom:14px">
      <button class="btn primary grow" data-act="new">＋ Nova vistoria</button>
      ${deferredInstall ? `<button class="btn ghost" data-act="install">Instalar app</button>` : ''}
    </div>
    ${list.length ? `<div class="card vlist">${list.map(vItem).join('')}</div>` : `<div class="empty">Nenhuma vistoria ainda.<br>Toque em “Nova vistoria” para começar.</div>`}`;
}
function vItem(v) {
  const h = v.header || {};
  const done = !!v.endAt;
  let pct = 0, total = 0, d0 = 0;
  for (const g of v.groups || []) { if (g.na) continue; const p = progressOf(defMap.get(g.groupId), g); d0 += p.done; total += p.total; }
  pct = total ? Math.round((d0 / total) * 100) : 0;
  const sync = !cloud.user() ? '' : (v.dirty ? '<span class="badge warn">não enviada</span>' : '<span class="badge sync">☁ ok</span>');
  return `<a class="item" href="#/v/${v.id}"><div class="grow"><div class="t">${E(h.empreendimento || 'Sem nome')}</div><div class="s">${E(h.siteId || 'sem ID')} · ${fmtDate(v.startAt)} ${fmtTime(v.startAt)} · ${pct}%</div></div>${sync}<span class="badge ${done ? 'done' : 'go'}">${done ? 'Finalizada' : 'Em andamento'}</span></a>`;
}

// ---------- VISTORIA (visão geral) ----------
async function viewVistoria() {
  setTop({ back: '#/', title: V.header.empreendimento || 'Nova vistoria', right: `<span class="pill" id="ovPct">${overall().pct}%</span>` });
  const h = V.header;
  const coverUrl = V.cover ? await photoURL(V.cover.id) : '';
  view.innerHTML = `
  <div class="card"><h2>Dados da vistoria</h2>
    <div class="grid">
      <div><label class="f">ID do Site</label><input type="text" data-h="siteId" value="${E(h.siteId)}" placeholder="ex.: BR-SP-ACD-03170-QMC" autocapitalize="characters"></div>
      <div><label class="f">Empreendimento</label><input type="text" data-h="empreendimento" value="${E(h.empreendimento)}"></div>
      <div><label class="f">Número do projeto</label><input type="text" data-h="projeto" value="${E(h.projeto)}"></div>
      <div><label class="f">Empresa implantadora</label><input type="text" data-h="implantadora" value="${E(h.implantadora)}"></div>
    </div>
    <label class="f" style="margin-top:14px">Operadora(s)</label>
    <div class="chips">${Object.entries(OPERADORAS).map(([k, o]) => `<button class="chip ${h.operadoras.includes(k) ? 'on' : ''}" data-op="${k}">${E(o.nome)}</button>`).join('')}</div>
    <div style="margin-top:10px"><input type="text" data-h="outraOperadora" value="${E(h.outraOperadora)}" placeholder="Outra operadora (opcional)"></div>
  </div>

  <div class="card"><h2>Foto do empreendimento</h2>
    ${coverUrl ? `<img class="cover-prev" src="${coverUrl}" alt="Foto do empreendimento">` : `<div class="cover-prev" style="display:grid;place-items:center;color:#7b8794">Sem foto — vai para a capa do relatório</div>`}
    <div class="row"><label class="btn teal small" style="cursor:pointer">${V.cover ? 'Trocar foto' : 'Enviar foto'}<input type="file" accept="image/*" id="coverIn" hidden></label>${V.cover ? '<button class="btn ghost small" data-act="rmcover">Remover</button>' : ''}</div>
  </div>

  <div class="card"><h2>Registro automático</h2>
    <div class="auto" id="auto">${autoHTML()}</div>
    <div class="row" style="margin-top:10px"><button class="btn ghost small" data-act="gps">📍 Atualizar localização</button></div>
  </div>

  <div id="groups">${groupsHTML()}</div>

  <div class="card"><button class="btn danger small" data-act="delv">Excluir esta vistoria</button></div>
  <div class="sticky"><a class="btn primary" style="text-decoration:none;text-align:center;display:block;padding-top:14px" href="#/v/${V.id}/fim">Revisar e gerar relatório →</a></div>`;
}
function autoHTML() {
  const g = V.gps;
  return `<div>Início: <b>${fmtDate(V.startAt)} ${fmtTime(V.startAt)}</b></div>
    <div>Fim: <b>${V.endAt ? fmtDate(V.endAt) + ' ' + fmtTime(V.endAt) : 'ao finalizar'}</b></div>
    <div>Localização: <b>${g ? `${g.lat.toFixed(5)}, ${g.lon.toFixed(5)} (±${g.acc} m)` : 'não obtida'}</b></div>`;
}
function groupsHTML() {
  let out = '';
  for (const d of defs) {
    const ins = instsOf(d);
    if (isUnique(d)) { if (ins[0]) out += `<div class="gh" style="display:none"></div>` + grpRow(d, ins[0]); continue; }
    out += `<div class="gh"><h3>${E(d.grupo)}</h3><button class="btn ghost small" data-add="${d.id}">＋ Adicionar</button></div>`;
    out += ins.length ? ins.map((i) => grpRow(d, i)).join('') : `<div class="muted" style="margin:-2px 0 8px 2px">Nenhum adicionado.</div>`;
  }
  return `<div class="card"><h2>Grupos do checklist</h2>${out}</div>`;
}
function grpRow(d, inst) {
  const p = progressOf(d, inst);
  const ph = inst.photos.length;
  return `<a class="grp ${inst.na ? 'na' : ''}" href="#/v/${V.id}/g/${inst.gid}"><div class="grow"><div class="t">${E(instTitle(d, inst))}</div>
    <div class="s">${inst.na ? 'Não se aplica' : `${p.done}/${p.total} · ${ph} foto${ph === 1 ? '' : 's'}`}${needsPhoto(d, inst) ? ' · <b style="color:var(--nok)">foto obrigatória</b>' : ''}</div>
    ${inst.na ? '' : `<div class="bar ${p.pct === 100 ? 'full' : ''}"><i style="width:${p.pct}%"></i></div>`}</div><span class="chev">›</span></a>`;
}
function refreshOverviewBits() {
  const g = $('#groups'); if (g) g.innerHTML = groupsHTML();
  const a = $('#auto'); if (a) a.innerHTML = autoHTML();
  const p = $('#ovPct'); if (p) p.textContent = overall().pct + '%';
}

// ---------- GRUPO ----------
function segHTML(key, val) {
  return `<div class="seg">${STATUS_ORDER.map((s) => `<button type="button" data-k="${E(key)}" data-v="${s}" class="${normAnswer(val) === s ? 'on' : ''}">${STATUS[s].label}</button>`).join('')}</div>`;
}
function ynHTML(key, val) {
  return `<div class="seg yn">${['sim', 'nao'].map((s) => `<button type="button" data-k="${E(key)}" data-v="${s}" class="${val === s ? 'on' : ''}">${YN[s].label}</button>`).join('')}</div>`;
}
function itemHTML(it, inst) {
  const a = inst.answers[it.key];
  let h = `<div class="it" data-item="${E(it.key)}"><div class="q">${E(it.text)}</div>`;
  if (it.anexo && it.tipo !== 'yn' && !it.extra) h += `<div class="hint">📎 ${E(it.anexo)}</div>`;
  h += it.tipo === 'yn' ? ynHTML(it.key, a) : segHTML(it.key, a);
  if (it.tipo === 'yn' && a === 'sim' && it.anexo) h += `<div class="hint" style="margin-top:8px">📎 ${E(it.anexo)}</div>`;
  if (it.extra) {
    const e = inst.answers[it.key + '#extra'];
    h += `<div class="sub-q"><div class="q">${E(it.extra)}</div>${ynHTML(it.key + '#extra', e)}${e === 'sim' && it.anexo ? `<div class="hint" style="margin-top:8px">📎 ${E(it.anexo)}</div>` : ''}</div>`;
  }
  return h + '</div>';
}
function campoHTML(c, inst) {
  const v = inst.values[c.key] || {};
  const warn = v.diff;
  return `<div class="num" data-ck="${E(c.key)}"><div class="l">${E(c.label)}</div>
    <div class="unit"><input type="number" inputmode="decimal" step="any" data-nv="${E(c.key)}" value="${E(v.v ?? '')}">${c.unit ? `<span>${c.unit}</span>` : ''}</div>
    ${c.amperagem ? `<label class="diff ${warn ? 'warn' : ''}"><input type="checkbox" data-nd="${E(c.key)}" ${warn ? 'checked' : ''}> Valor difere do projeto (anexar foto/print de conformidade)</label>` : ''}</div>`;
}

let curInst = null, curDef = null;
async function viewGroup(gid) {
  const inst = V.groups.find((g) => g.gid === gid);
  if (!inst) { location.hash = `#/v/${V.id}`; return; }
  const d = defMap.get(inst.groupId);
  curInst = inst; curDef = d;
  setTop({ back: `#/v/${V.id}`, title: instTitle(d, inst) });
  const hasSt = d.items.length > 0;
  view.innerHTML = `
  <div class="card" style="display:flex;align-items:center;gap:12px">
    <div class="grow"><b id="gprog"></b><div class="bar" id="gbar"><i></i></div></div>
    <label class="row" style="gap:8px;font-size:14px"><input type="checkbox" id="naChk" ${inst.na ? 'checked' : ''} style="width:20px;height:20px"> Não se aplica</label>
  </div>
  <div class="need" id="need" hidden></div>
  <div id="body" ${inst.na ? 'hidden' : ''}>
    ${hasSt ? `<div class="card"><h2>Checklist</h2>
      <div class="row" style="margin-bottom:6px"><button class="btn ghost small" data-act="allok">Marcar pendentes como OK</button></div>
      <div id="items">${d.items.map((it) => itemHTML(it, inst)).join('')}</div></div>` : ''}
    ${d.campos.length ? `<div class="card"><h2>${d.tipo === 'medicao' ? 'Medições' : 'Amperagens dos disjuntores'}</h2>${d.campos.map((c) => campoHTML(c, inst)).join('')}</div>` : ''}
    <div class="card"><h2>Observação do grupo</h2><textarea data-obs placeholder="Observações gerais deste grupo (vai no slide do checklist)">${E(inst.obs)}</textarea></div>
    <div class="card"><h2>Fotos <span class="muted" style="text-transform:none;letter-spacing:0">(2 por slide)</span></h2>
      <div class="row" style="margin-bottom:12px">
        <label class="btn primary small" style="cursor:pointer">📷 Tirar foto<input type="file" accept="image/*" capture="environment" id="camIn" hidden></label>
        <label class="btn teal small" style="cursor:pointer">🖼 Galeria<input type="file" accept="image/*" multiple id="galIn" hidden></label>
      </div>
      <div class="photos" id="photos"></div>
    </div>
  </div>
  <div class="sticky"><a class="btn teal" style="text-decoration:none;text-align:center;display:block;padding-top:14px" href="#/v/${V.id}">Concluir grupo ✓</a></div>`;
  updateGroupProgress();
  renderPhotos();
}
function updateGroupProgress() {
  const p = progressOf(curDef, curInst);
  const el = $('#gprog'); if (!el) return;
  const sc = scoreGroup(curDef, curInst);
  el.textContent = curDef.items.length ? `${p.done} de ${p.total} respondidos${sc.pct != null ? ` · conformidade ${sc.pct}%` : ''}` : `${p.done} de ${p.total} preenchidos`;
  const nd = $('#need'); if (nd) { const need = needsPhoto(curDef, curInst); nd.hidden = !need; if (need) nd.textContent = `📷 Há ${sc.nok} item(ns) “Não conf.” — foto de evidência obrigatória neste grupo.`; }
  const bar = $('#gbar'); bar.className = 'bar ' + (p.pct === 100 ? 'full' : ''); bar.firstElementChild.style.width = p.pct + '%';
}
async function renderPhotos() {
  const box = $('#photos'); if (!box) return;
  updateGroupProgress();
  const ps = curInst.photos;
  if (!ps.length) { box.innerHTML = '<div class="muted">Nenhuma foto neste grupo.</div>'; return; }
  const srcs = await Promise.all(ps.map((p) => photoURL(p.id)));
  box.innerHTML = ps.map((p, i) => `<div class="ph" data-pid="${p.id}">
    ${srcs[i] ? `<img src="${srcs[i]}" alt="Foto ${i + 1}">` : `<div class="muted" style="padding:40px;text-align:center;background:#eee">Foto ainda não baixada</div>`}
    <div class="b"><label class="f">Observação da foto ${i + 1}</label><textarea data-po="${p.id}" placeholder="O que esta foto mostra?">${E(p.obs)}</textarea>
    <div class="acts"><span class="n">Slide ${Math.floor(i / 2) + 1} · ${i % 2 ? 'direita' : 'esquerda'}</span>
      <button class="btn ghost small" data-pm="up" data-pid="${p.id}" ${i === 0 ? 'disabled' : ''}>↑</button>
      <button class="btn ghost small" data-pm="down" data-pid="${p.id}" ${i === ps.length - 1 ? 'disabled' : ''}>↓</button>
      <button class="btn danger small" data-pd="${p.id}">Excluir</button></div></div></div>`).join('');
}
async function addPhotos(files) {
  if (!files?.length) return;
  toast('Processando foto…', 1500);
  for (const f of files) {
    try {
      const { blob, w, h } = await toJpeg(f);
      const id = uid();
      await db.putPhoto({ id, vid: V.id, blob, w, h, synced: false });
      curInst.photos.push({ id, obs: '', w, h });
    } catch (e) { console.error(e); toast('Não foi possível ler uma das fotos'); }
  }
  touch(); renderPhotos();
}
async function removePhoto(id) {
  const rec = await db.getPhoto(id);
  if (rec?.synced) (V.removedPhotos ||= []).push(id);
  await db.deletePhoto(id);
  if (urls.has(id)) { URL.revokeObjectURL(urls.get(id)); urls.delete(id); }
}

// ---------- FINALIZAÇÃO ----------
function collectWarnings() {
  const w = [], block = [];
  let pend = 0;
  for (const g of V.groups) {
    if (g.na) continue;
    const d = defMap.get(g.groupId); const sc = scoreGroup(d, g);
    if (hasContent(d, g) || sc.pend < sc.total) {
      pend += sc.pend;
      if (needsPhoto(d, g)) block.push(`${instTitle(d, g)}: ${sc.nok} item(ns) "Não conf." — adicione ao menos 1 foto de evidência`);
      const diff = d.campos.filter((c) => g.values[c.key]?.diff).length;
      if (diff && !g.photos.length) w.push(`${instTitle(d, g)}: amperagem difere do projeto, sem print/foto anexado`);
    }
  }
  return { w, block, pend };
}
async function viewFinish() {
  setTop({ back: `#/v/${V.id}`, title: 'Relatório' });
  const rows = [], vazios = [];
  for (const d of defs) for (const g of instsOf(d)) {
    if (g.na) continue;
    const sc = scoreGroup(d, g);
    if (!hasContent(d, g)) { vazios.push(instTitle(d, g)); continue; }
    rows.push(`<tr><td>${E(instTitle(d, g))}</td><td class="c">${sc.ok}</td><td class="c">${sc.nok}</td><td class="c">${sc.na}</td><td class="c">${sc.pend}</td><td class="c">${sc.pct ?? '—'}${sc.pct != null ? '%' : ''}</td><td class="c">${g.photos.length}</td></tr>`);
  }
  const { w, block, pend } = collectWarnings();
  const canShare = !!(navigator.canShare && navigator.share);
  view.innerHTML = `
  ${block.length ? `<div class="need"><b>Foto obrigatória antes de gerar o relatório:</b><ul style="margin:6px 0 0 18px;padding:0;font-weight:500">${block.map((x) => `<li>${E(x)}</li>`).join('')}</ul></div>` : ''}
  ${w.length ? `<div class="warnbox"><b>Atenção:</b><ul style="margin:6px 0 0 18px;padding:0">${w.map((x) => `<li>${E(x)}</li>`).join('')}</ul></div>` : (block.length ? '' : `<div class="okbox">Nenhuma pendência crítica encontrada.</div>`)}
  ${pend ? `<div class="warnbox">Há <b>${pend}</b> item(ns) sem resposta nos grupos preenchidos — aparecerão como “Pendente”.</div>` : ''}
  <div class="card"><h2>Resumo por grupo</h2>
    <div style="overflow:auto"><table class="sum"><thead><tr><th>Grupo</th><th class="c">OK</th><th class="c">NC</th><th class="c">N/A</th><th class="c">Pend.</th><th class="c">%</th><th class="c">Fotos</th></tr></thead><tbody>${rows.join('') || '<tr><td colspan="7" class="muted">Nenhum grupo preenchido ainda.</td></tr>'}</tbody></table></div>
    ${vazios.length ? `<p class="muted"><b>Fora do relatório</b> (sem respostas, fotos ou observação): ${E(vazios.join(' · '))}</p>` : ''}
    <p class="muted">NC = não conforme.</p>
  </div>
  <div class="card"><h2>Horário</h2><div class="auto">${autoHTML()}</div></div>
  <div class="card"><div class="row">
    <button class="btn primary grow" data-act="gen" ${block.length ? 'disabled' : ''}>${V.endAt ? 'Gerar PPTX novamente' : 'Finalizar vistoria e gerar PPTX'}</button>
    ${V.endAt ? '<button class="btn ghost" data-act="reopen">Reabrir</button>' : ''}</div>
    <div class="row" style="margin-top:10px">
      ${canShare ? '<button class="btn ghost small" data-act="share" id="shareBtn" disabled>Compartilhar</button>' : ''}
      ${cloud.user() ? '<button class="btn ghost small" data-act="onedrive" id="odBtn" disabled>Salvar no OneDrive</button>' : ''}
    </div>
    <p class="muted" id="genMsg"></p></div>`;
}

let lastPptx = null;
async function generate() {
  const btn = $('[data-act=gen]'); const msg = $('#genMsg');
  if (collectWarnings().block.length) { msg.textContent = 'Há itens "Não conf." sem foto.'; return; }
  btn.disabled = true; msg.textContent = 'Gerando relatório…';
  try {
    if (!V.endAt) { V.endAt = new Date().toISOString(); touch(); await flush(); }
    const tpl = await (await fetch('assets/template.pptx')).arrayBuffer();
    const photoBytes = new Map(); let missing = 0;
    for (const g of V.groups) for (const p of g.photos) {
      const rec = await db.getPhoto(p.id);
      if (rec?.blob) photoBytes.set(p.id, new Uint8Array(await rec.blob.arrayBuffer())); else missing++;
    }
    const logos = {};
    for (const k of V.header.operadoras) {
      const o = OPERADORAS[k]; if (!o) continue;
      const blob = await (await fetch(o.logo)).blob();
      const bm = await createImageBitmap(blob);
      logos[k] = { bytes: new Uint8Array(await blob.arrayBuffer()), w: bm.width, h: bm.height };
    }
    let coverBytes = null;
    if (V.cover) { const rec = await db.getPhoto(V.cover.id); if (rec?.blob) coverBytes = new Uint8Array(await rec.blob.arrayBuffer()); }
    const bytes = await buildPptx({ JSZip: window.JSZip, template: tpl, vistoria: V, checklist, photoBytes, logos, coverBytes });
    const safe = (V.header.siteId || 'vistoria').replace(/[^\w.-]+/g, '_');
    const name = `Vistoria_${safe}_${V.startAt.slice(0, 10).replace(/-/g, '')}.pptx`;
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    lastPptx = { blob, bytes, name };
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
    msg.textContent = `Relatório gerado: ${name} (${(bytes.length / 1048576).toFixed(1)} MB)` + (missing ? ` — ${missing} foto(s) ainda não baixadas ficaram de fora` : '');
    $('#shareBtn')?.removeAttribute('disabled'); $('#odBtn')?.removeAttribute('disabled');
    btn.textContent = 'Gerar PPTX novamente';
  } catch (e) { console.error(e); msg.textContent = 'Erro ao gerar: ' + e.message; }
  finally { btn.disabled = false; }
}

// ---------- eventos ----------
async function onClick(e) {
  const t = e.target;
  // checklist segmentado
  const sb = t.closest('.seg button');
  if (sb && curInst) {
    const key = sb.dataset.k, val = sb.dataset.v;
    if (curInst.answers[key] === val) delete curInst.answers[key]; else curInst.answers[key] = val;
    const baseKey = key.replace('#extra', '');
    const it = curDef.items.find((i) => i.key === baseKey);
    const card = sb.closest('.it');
    if (it && (it.extra || it.tipo === 'yn')) card.outerHTML = itemHTML(it, curInst);
    else sb.parentElement.querySelectorAll('button').forEach((b) => b.classList.toggle('on', curInst.answers[key] === b.dataset.v));
    updateGroupProgress(); touch(); return;
  }
  const opb = t.closest('[data-op]');
  if (opb) {
    const k = opb.dataset.op, a = V.header.operadoras;
    const i = a.indexOf(k); if (i >= 0) a.splice(i, 1); else a.push(k);
    opb.classList.toggle('on'); syncOperatorGroups(); touch(); refreshOverviewBits(); return;
  }
  const add = t.closest('[data-add]');
  if (add) {
    const d = defMap.get(add.dataset.add); const n = instsOf(d).length + 1;
    const label = await askText(`Identificação (${d.grupo})`, `${DEFAULT_LABEL[d.id] || 'Item'} ${n}`);
    if (label === null) return;
    V.groups.push(newInst(d, label)); touch(); refreshOverviewBits(); return;
  }
  const pm = t.closest('[data-pm]');
  if (pm) {
    const i = curInst.photos.findIndex((p) => p.id === pm.dataset.pid); const j = pm.dataset.pm === 'up' ? i - 1 : i + 1;
    if (j >= 0 && j < curInst.photos.length) { [curInst.photos[i], curInst.photos[j]] = [curInst.photos[j], curInst.photos[i]]; touch(); renderPhotos(); }
    return;
  }
  const pd = t.closest('[data-pd]');
  if (pd) {
    if (!confirm('Excluir esta foto?')) return;
    const id = pd.dataset.pd; curInst.photos = curInst.photos.filter((p) => p.id !== id);
    await removePhoto(id); touch(); renderPhotos(); return;
  }
  const act = t.closest('[data-act]')?.dataset.act;
  if (!act) return;
  switch (act) {
    case 'new': {
      V = newVistoria(); await db.putVistoria(V); location.hash = `#/v/${V.id}`;
      getGPS().then((g) => { if (g && V) { V.gps = g; touch(); refreshOverviewBits(); } else if (!g) toast('Localização não disponível — permita o acesso ao GPS'); });
      break;
    }
    case 'install': deferredInstall?.prompt(); deferredInstall = null; route(); break;
    case 'login': toast('Abrindo login da Microsoft…', 4000); cloud.signIn().catch((e) => toast('Erro no login: ' + (e.errorCode ? e.errorCode + ' – ' : '') + (e.message || e), 12000)); break;
    case 'logout': if (confirm('Sair da conta Microsoft? As vistorias continuam no aparelho.')) cloud.signOut(); break;
    case 'sync': doSync(true); break;
    case 'gps': { toast('Obtendo localização…', 1500); const g = await getGPS(); if (g) { V.gps = g; touch(); refreshOverviewBits(); toast('Localização atualizada'); } else toast('Não foi possível obter a localização'); break; }
    case 'rmcover': if (V.cover) { await removePhoto(V.cover.id); V.cover = null; touch(); viewVistoria(); } break;
    case 'allok':
      for (const it of curDef.items) if (it.tipo === 'status' && !curInst.answers[it.key]) curInst.answers[it.key] = 'ok';
      touch(); $('#items').innerHTML = curDef.items.map((it) => itemHTML(it, curInst)).join(''); updateGroupProgress(); break;
    case 'delv':
      if (confirm('Excluir esta vistoria? Isto também remove do OneDrive.')) {
        const v = V; V = null;
        try { if (cloud.user() && navigator.onLine) await cloud.removeRemote(v); } catch (err) { console.warn(err); }
        await db.deleteVistoria(v.id); location.hash = '#/'; toast('Vistoria excluída');
      } break;
    case 'gen': generate(); break;
    case 'reopen': V.endAt = null; touch(); viewFinish(); break;
    case 'share':
      if (lastPptx) { try { await navigator.share({ files: [new File([lastPptx.blob], lastPptx.name, { type: lastPptx.blob.type })], title: lastPptx.name }); } catch (err) { /* cancelado */ } }
      break;
    case 'onedrive':
      if (lastPptx) { try { toast('Enviando…', 1500); await cloud.savePptx(V, lastPptx.bytes, lastPptx.name); toast('PPTX salvo no OneDrive'); } catch (err) { toast('Falha: ' + err.message, 5000); } }
      break;
  }
}
view.addEventListener('click', onClick);
$('#top').addEventListener('click', onClick);

view.addEventListener('input', (e) => {
  const t = e.target;
  if (t.dataset.h) { V.header[t.dataset.h] = t.value; if (t.dataset.h === 'empreendimento') $('#top h1').textContent = t.value || 'Nova vistoria'; touch(); }
  else if (t.dataset.obs !== undefined) { curInst.obs = t.value; touch(); }
  else if (t.dataset.po) { const p = curInst.photos.find((x) => x.id === t.dataset.po); if (p) { p.obs = t.value; touch(); } }
  else if (t.dataset.nv) { const v = (curInst.values[t.dataset.nv] ||= {}); v.v = t.value; touch(); updateGroupProgress(); }
});
view.addEventListener('change', async (e) => {
  const t = e.target;
  if (t.id === 'naChk') { curInst.na = t.checked; $('#body').hidden = t.checked; touch(); }
  else if (t.dataset.nd) { const v = (curInst.values[t.dataset.nd] ||= {}); v.diff = t.checked; t.closest('.diff').classList.toggle('warn', t.checked); touch(); }
  else if (t.id === 'camIn' || t.id === 'galIn') { await addPhotos([...t.files]); t.value = ''; }
  else if (t.id === 'coverIn' && t.files[0]) {
    try {
      toast('Processando foto…', 1500);
      const { blob, w, h } = await coverJpeg(t.files[0]);
      if (V.cover) await removePhoto(V.cover.id);
      const id = uid(); await db.putPhoto({ id, vid: V.id, blob, w, h, synced: false });
      V.cover = { id, w, h }; touch(); viewVistoria();
    } catch (err) { console.error(err); toast('Não foi possível ler a foto'); }
  }
});

// ---------- roteador ----------
function parse() {
  const p = location.hash.replace(/^#\/?/, '').split('/');
  if (p[0] === 'v' && p[1]) { if (p[2] === 'g' && p[3]) return { v: 'group', id: p[1], gid: p[3] }; if (p[2] === 'fim') return { v: 'finish', id: p[1] }; return { v: 'vistoria', id: p[1] }; }
  return { v: 'home' };
}
async function route() {
  await flush();
  const r = parse();
  window.scrollTo(0, 0);
  curInst = null; curDef = null;
  if (r.v === 'home') return viewHome();
  if (!V || V.id !== r.id) V = await db.getVistoria(r.id);
  if (!V) { location.hash = '#/'; return; }
  if (r.v === 'vistoria') return viewVistoria();
  if (r.v === 'group') return viewGroup(r.gid);
  if (r.v === 'finish') return viewFinish();
}

window.addEventListener('hashchange', route);
window.addEventListener('beforeinstallprompt', (e) => { e.preventDefault(); deferredInstall = e; if (parse().v === 'home') viewHome(); });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') flush(); });
window.addEventListener('online', () => doSync(false));

// ---------- início ----------
(async function init() {
  try {
    checklist = await (await fetch('data/checklist.json')).json();
    defs = expandAll(checklist); defMap = new Map(defs.map((d) => [d.id, d]));
    if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW', e));
    try { await cloud.initAuth(); } catch (e) { toast('Erro no login Microsoft: ' + (e.message || e), 6000); }
    await route();
    if (cloud.user() && navigator.onLine) doSync(false);
  } catch (e) {
    console.error(e);
    view.innerHTML = `<div class="warnbox">Não foi possível iniciar o app: ${E(e.message)}</div>`;
  }
})();
