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
