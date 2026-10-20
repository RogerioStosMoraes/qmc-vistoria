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

export function scoreGroup(def, inst) {
  const c = { ok: 0, na: 0, nok: 0, pend: 0 };
  for (const it of def.items) {
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
  if (inst.obs && inst.obs.trim()) return true;
  if ((inst.photos || []).length) return true;
  if (Object.keys(inst.answers || {}).length) return true;
  if (Object.values(inst.qty || {}).some((q) => Object.values(q || {}).some((x) => String(x ?? '').trim() !== ''))) return true;
  return Object.values(inst.values || {}).some((v) => v && String(v.v ?? '').trim() !== '');
}

export function instTitle(def, inst) {
  return inst.label ? `${def.grupo} – ${inst.label}` : def.grupo;
}

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
export const ETAPAS = { inst: 'Instalado', fus: 'Fusionado', conec: 'Conectado' };
export const ETAPA_ORDER = ['inst', 'fus', 'conec'];

export function expandAcomp(cfg) {
  return cfg.grupos.map((g) => ({
    id: g.id, grupo: g.grupo, instancias: 'única', tipo: 'acomp', campos: [],
    items: g.itens.map((it) => ({ key: slug(it.item), text: it.item, etapas: it.etapas, projDe: it.proj_de || null, dica: it.dica || null })),
  }));
}

export function num(x) {
  if (x === undefined || x === null || String(x).trim() === '') return null;
  const n = Number(String(x).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

/** Cálculo de um item: % de cada etapa (valor/projeto), % do item (média das etapas) e pendente. */
export function itemCalc(v, it, q = {}) {
  const proj = num(q.proj) ?? num(it.projDe ? v.header?.[it.projDe] : null);
  const vals = {}, pcts = {};
  let any = false;
  for (const e of it.etapas) { vals[e] = num(q[e]); if (vals[e] !== null) any = true; }
  const last = it.etapas[it.etapas.length - 1];
  let pct = null;
  if (proj && proj > 0 && any) {
    let sum = 0;
    for (const e of it.etapas) { const r = Math.min(1, (vals[e] ?? 0) / proj); pcts[e] = Math.round(r * 100); sum += r; }
    pct = Math.round((sum / it.etapas.length) * 100);
  }
  const pend = proj != null ? Math.max(0, proj - (vals[last] ?? 0)) : null;
  return { proj, vals, pcts, pct, pend, filled: any, semProjeto: any && !(proj > 0) };
}
export function groupCalc(v, def, inst) {
  const items = def.items.map((it) => ({ it, c: itemCalc(v, it, inst.qty?.[it.key]) }));
  const ps = items.filter((x) => x.c.pct !== null).map((x) => x.c.pct);
  return { items, pct: ps.length ? Math.round(ps.reduce((a, b) => a + b, 0) / ps.length) : null, filled: items.filter((x) => x.c.filled).length, total: items.length };
}
/** Avanço geral = média do avanço dos grupos (como na planilha antiga). */
export function overallCalc(v, defMap) {
  const groups = [];
  for (const g of v.groups || []) {
    if (g.na) continue;
    const d = defMap.get(g.groupId); if (!d) continue;
    groups.push({ g, d, ...groupCalc(v, d, g) });
  }
  const ps = groups.filter((x) => x.pct !== null).map((x) => x.pct);
  return { groups, pct: ps.length ? Math.round(ps.reduce((a, b) => a + b, 0) / ps.length) : null };
}
/** Resumo guardado na próxima versão, para comparar a evolução. */
export function snapshotOf(v, defMap) {
  const o = overallCalc(v, defMap);
  const groups = {};
  for (const x of o.groups) {
    const items = {};
    for (const { it, c } of x.items) items[it.key] = { proj: c.proj, vals: c.vals, pct: c.pct };
    groups[x.g.groupId] = { pct: x.pct, items };
  }
  return { id: v.id, versao: v.versao || 1, data: v.endAt || v.startAt, pct: o.pct, groups };
}
export const fmtPct = (p) => (p == null ? '—' : p + '%');
export function deltaLabel(cur, prev) {
  if (cur == null || prev == null) return '';
  const d = cur - prev;
  return d === 0 ? '= 0 p.p.' : `${d > 0 ? '▲ +' : '▼ '}${d} p.p.`;
}
