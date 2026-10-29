// Modelo de dados do checklist (compartilhado entre o app e o gerador de PPTX)

export const STATUS = {
  ok:  { label: 'OK',        fill: 'DDF1E3', color: '1E6B34' },
  na:  { label: 'N/A',       fill: 'EAEAEA', color: '6D6E71' },
  nok: { label: 'Não conf.', fill: 'FBDADA', color: 'B3261E' },
};
export const STATUS_ORDER = ['ok', 'nok', 'na'];
/** Respostas antigas "bom" contam como OK. */
export const normAnswer = (a) => (a === 'bom' ? 'ok' : a);

export const YN = {
  sim: { label: 'Sim', fill: 'FDE9D2', color: '9A4A00' },
  nao: { label: 'Não', fill: 'EAEAEA', color: '6D6E71' },
};

export function slug(s) {
  return String(s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
    .slice(0, 60);
}

const UNIT = { numero_A: 'A', numero_V: 'V' };

/** Expande um grupo do checklist.json: blocos reutilizáveis + itens próprios + campos. */
export function expandGroup(checklist, g) {
  const bloco = g.usa_bloco ? (checklist.blocos_reutilizaveis[g.usa_bloco] || []) : [];
  const raw = [...bloco, ...(g.itens || [])];
  const items = raw.map((it) => ({
    key: slug(it.item),
    text: it.item,
    anexo: it.anexo_obrigatorio_se || null,
    tipo: it.pergunta_extra === true ? 'yn' : 'status',
    extra: typeof it.pergunta_extra === 'string' ? it.pergunta_extra : null,
  }));
  const campos = (g.campos || []).map((c) => ({
    key: 'c-' + slug(c.campo),
    label: c.campo,
    unit: UNIT[c.tipo] || '',
    anexo: c.anexo_obrigatorio_se || null,
    amperagem: /amperagem/i.test(c.campo),
  }));
  return { id: g.id, grupo: g.grupo, instancias: g.instancias, items, campos, tipo: g.tipo || 'checklist' };
}

export function expandAll(checklist) {
  return checklist.grupos.map((g) => expandGroup(checklist, g));
}

/** Itens do modelo + itens adicionados em campo (inst.extras). */
export function itemsOf(def, inst) {
  const ex = (inst?.extras || []).map((x) => ({ key: x.k, text: x.text, anexo: null, tipo: 'status', extra: null, custom: true }));
  return ex.length ? [...def.items, ...ex] : def.items;
}

export function scoreGroup(def, inst) {
  const c = { ok: 0, na: 0, nok: 0, pend: 0 };
  for (const it of itemsOf(def, inst)) {
    if (it.tipo !== 'status') continue;
    const a = normAnswer(inst.answers?.[it.key]);
    if (a && c[a] !== undefined) c[a]++; else c.pend++;
  }
  const apl = c.ok + c.nok;
  const total = c.ok + c.na + c.nok + c.pend;
  return { ...c, total, aplicaveis: apl, pct: apl ? Math.round((c.ok / apl) * 100) : null };
}

export function hasContent(def, inst) {
  if (inst.na) return false;
  if (def.tipo === 'acomp' && def.multi) return true; // quadros / IDFs adicionados sempre entram no relatório
  if (inst.obs && inst.obs.trim()) return true;
  if ((inst.photos || []).length) return true;
  if (Object.keys(inst.answers || {}).length) return true;
  if ((inst.extras || []).length) return true;
  if (Object.entries(inst.quadro || {}).some(([k, x]) => ['inst', 'conec', 'energ'].includes(k) ? !!x : (k === 'cabos' && String(x ?? '').trim() !== ''))) return true;
  if (Object.values(inst.qty || {}).some((q) => Object.values(q || {}).some((x) => String(x ?? '').trim() !== ''))) return true;
  return Object.values(inst.values || {}).some((v) => v && String(v.v ?? '').trim() !== '');
}

export function instTitle(def, inst) {
  if (def.tipo === 'acomp') return acTitle(def, inst);
  return inst.label ? `${def.grupo} – ${inst.label}` : def.grupo;
}
/** Título de um grupo/subgrupo do acompanhamento. */
export function acTitle(def, inst) {
  const l = (inst?.label || '').trim();
  if (def.modo === 'quadro') return `Quadro${l ? ' ' + l : ''}${inst?.quadro?.dono ? ` (${inst.quadro.dono})` : ''}`;
  if (def.soLabel) return l || def.grupo;
  return l ? `${def.grupo} – ${l}` : def.grupo;
}
/** Chave estável de um grupo entre versões (para comparar com a vistoria anterior). */
export const instKey = (def, inst) => (def.multi ? `${def.id}|${(inst.label || '').trim().toLowerCase()}` : def.id);

export const OPERADORAS = {
  vivo: { nome: 'Vivo', logo: 'assets/logos/vivo.png' },
  claro: { nome: 'Claro', logo: 'assets/logos/claro.png' },
  tim: { nome: 'TIM', logo: 'assets/logos/tim.png' },
};

export function uid() {
  return (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2, 10));
}

export function pad(n) { return String(n).padStart(2, '0'); }
export function fmtDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()}`;
}
export function fmtTime(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Foto de evidência é obrigatória quando há item "Não conf." no grupo. */
export function needsPhoto(def, inst) {
  if (inst.na) return false;
  return scoreGroup(def, inst).nok > 0 && !(inst.photos || []).length;
}

/** Nome base "ID do site - Empreendimento" seguro para arquivo/pasta. */
export function baseName(v) {
  const clean = (x) => String(x || '').replace(/[\\/:*?"<>|#%~]+/g, ' ').replace(/\s+/g, ' ').trim();
  const parts = [clean(v.header?.siteId), clean(v.header?.empreendimento)].filter(Boolean);
  return (parts.join(' - ') || 'Vistoria').slice(0, 90).trim();
}

// ---------- tipos de vistoria ----------
export const TIPOS = {
  prehof: { nome: 'Pré-HOF', arquivo: 'data/checklist.json' },
  acomp: { nome: 'Acompanhamento', arquivo: 'data/acompanhamento.json' },
};
export const tipoOf = (v) => (v?.tipo === 'acomp' ? 'acomp' : 'prehof');
export const versaoLabel = (v) => `v${v.versao || 1}.0`;
/** Rótulo do tipo para nomes de arquivo/pasta: "" (Pré-HOF) ou "Acompanhamento v2.0". */
export function typeTag(v) { return tipoOf(v) === 'acomp' ? `Acompanhamento ${versaoLabel(v)}` : ''; }

// ---------- acompanhamento (instalado x projetado) ----------
export const ETAPAS = { inst: 'Instalado', fus: 'Fusionado', conec: 'Conectado', energ: 'Energia' };
export const ETAPA_ORDER = ['inst', 'fus', 'conec', 'energ'];

export function expandAcomp(cfg) {
  return cfg.grupos.map((g) => ({
    id: g.id, grupo: g.grupo, pai: g.pai || null, tipo: 'acomp', campos: [],
    instancias: g.instancias || 'única', multi: (g.instancias || 'única') !== 'única',
    tituloLista: g.titulo_lista || g.grupo, modo: g.modo || 'itens', soLabel: !!g.so_label, labelPadrao: g.label_padrao || 'Item', rotuloAdd: g.adicionar || 'Adicionar', dica: g.dica || null,
    items: (g.itens || []).map((it) => ({
      key: slug(it.item), text: it.item, etapas: it.etapas, projDe: it.proj_de || null, dica: it.dica || null,
      metragem: !!it.metragem, preconect: !!it.preconectorizado,
    })),
  }));
}

export function num(x) {
  if (x === undefined || x === null || String(x).trim() === '') return null;
  const n = Number(String(x).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

/** Etapas válidas do item (pré-conectorizado não tem fusão). */
export const etapasOf = (it, q = {}) => (it.preconect && q.precon ? it.etapas.filter((e) => e !== 'fus') : it.etapas);

/**
 * Cálculo de um item: % de cada etapa, % do item (média das etapas) e pendente.
 * Itens com metragem: a etapa "instalado" é medida em metros (instalado/projeto) quando há metragem de projeto.
 */
export function itemCalc(v, it, q = {}) {
  const etapas = etapasOf(it, q);
  const proj = num(q.proj) ?? num(it.projDe ? v.header?.[it.projDe] : null);
  const projM = it.metragem ? num(q.projM) : null;
  const instM = it.metragem ? num(q.instM) : null;
  const vals = {}, pcts = {};
  let any = instM !== null;
  for (const e of etapas) { vals[e] = num(q[e]); if (vals[e] !== null) any = true; }
  let pct = null;
  if (proj && proj > 0 && any) {
    let sum = 0;
    for (const e of etapas) {
      const usaM = e === 'inst' && it.metragem && projM > 0 && instM !== null;
      const r = Math.min(1, usaM ? instM / projM : (vals[e] ?? 0) / proj);
      pcts[e] = Math.round(r * 100); sum += r;
    }
    pct = Math.round((sum / etapas.length) * 100);
  }
  const pend = proj != null ? Math.max(0, proj - Math.min(...etapas.map((e) => vals[e] ?? 0))) : null;
  return { etapas, proj, projM, instM, vals, pcts, pct, pend, filled: any, precon: !!(it.preconect && q.precon), usaMetros: !!(it.metragem && projM > 0 && instM !== null), semProjeto: any && !(proj > 0) };
}

export const QUADRO_ETAPAS = [['inst', 'Instalado'], ['conec', 'Conectado'], ['energ', 'Energizado']];
export function quadroCalc(inst) {
  const q = inst.quadro || {};
  const n = QUADRO_ETAPAS.filter(([k]) => q[k]).length;
  const cabos = num(q.cabos);
  return { pct: Math.round((n / QUADRO_ETAPAS.length) * 100), n, cabos, unid: q.unid === 'm' ? 'm' : 'un', filled: n > 0 || cabos !== null, dono: q.dono || '', acao: q.acao || '' };
}

export function groupCalc(v, def, inst) {
  if (def.modo === 'quadro') { const q = quadroCalc(inst); return { items: [], pct: q.pct, filled: q.filled ? 1 : 0, total: 1, quadro: q }; }
  const items = def.items.map((it) => ({ it, c: itemCalc(v, it, inst.qty?.[it.key]) }));
  const ps = items.filter((x) => x.c.pct !== null).map((x) => x.c.pct);
  return { items, pct: ps.length ? Math.round(ps.reduce((a, b) => a + b, 0) / ps.length) : null, filled: items.filter((x) => x.c.filled).length, total: items.length };
}
const mean = (a) => (a.length ? Math.round(a.reduce((x, y) => x + y, 0) / a.length) : null);
/**
 * Avanço geral = média dos grupos lógicos (HBTS, Quadros, IDF, Nanos).
 * IDF = média do resumo de IDFs + cada IDF adicionado; Quadros = média dos quadros.
 */
export function overallCalc(v, defMap) {
  const order = [...defMap.keys()];
  const groups = [];
  for (const g of v.groups || []) {
    if (g.na) continue;
    const d = defMap.get(g.groupId); if (!d) continue;
    groups.push({ g, d, ...groupCalc(v, d, g) });
  }
  const byL = new Map();
  for (const x of groups) {
    const lid = x.d.pai || x.d.id;
    if (!byL.has(lid)) byL.set(lid, { id: lid, nome: defMap.get(lid)?.grupo || x.d.grupo, list: [] });
    byL.get(lid).list.push(x);
  }
  const logical = [...byL.values()].sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id)).map((L) => ({ ...L, pct: mean(L.list.filter((x) => x.pct !== null).map((x) => x.pct)) }));
  return { groups, logical, pct: mean(logical.filter((L) => L.pct !== null).map((L) => L.pct)) };
}
/** Resumo guardado na próxima versão, para comparar a evolução. */
export function snapshotOf(v, defMap) {
  const o = overallCalc(v, defMap);
  const groups = {};
  for (const x of o.groups) {
    const items = {};
    for (const { it, c } of x.items) items[it.key] = { proj: c.proj, projM: c.projM, instM: c.instM, vals: c.vals, pct: c.pct };
    groups[instKey(x.d, x.g)] = { pct: x.pct, items, quadro: x.quadro ? { ...x.quadro, flags: { inst: !!x.g.quadro?.inst, conec: !!x.g.quadro?.conec, energ: !!x.g.quadro?.energ } } : undefined };
  }
  const logical = {}; for (const L of o.logical) logical[L.id] = L.pct;
  return { id: v.id, versao: v.versao || 1, data: v.endAt || v.startAt, pct: o.pct, groups, logical };
}
export const fmtPct = (p) => (p == null ? '—' : p + '%');
export function deltaLabel(cur, prev) {
  if (cur == null || prev == null) return '';
  const d = cur - prev;
  return d === 0 ? '= 0 p.p.' : `${d > 0 ? '▲ +' : '▼ '}${d} p.p.`;
}
