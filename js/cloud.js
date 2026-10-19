// Login Microsoft (MSAL) + sincronização com o OneDrive do próprio usuário (Microsoft Graph)
import { CONFIG } from '../config.js';
import { baseName } from './model.js';
import * as db from './db.js';

let pca = null;
let account = null;
const G = 'https://graph.microsoft.com/v1.0';

export const configured = () => !!CONFIG.clientId && !/^COLE/i.test(CONFIG.clientId);
export const user = () => account;

export function redirectUri() {
  const u = new URL(location.href);
  u.hash = ''; u.search = '';
  u.pathname = u.pathname.replace(/index\.html$/, '');
  return u.href;
}

export async function initAuth() {
  if (!configured() || !window.msal) return null;
  pca = new msal.PublicClientApplication({
    auth: { clientId: CONFIG.clientId, authority: CONFIG.authority, redirectUri: redirectUri(), postLogoutRedirectUri: redirectUri(), navigateToLoginRequestUrl: false },
    cache: { cacheLocation: 'localStorage' },
  });
  await pca.initialize();
  try {
    const r = await pca.handleRedirectPromise();
    if (r && r.account) pca.setActiveAccount(r.account);
  } catch (e) { console.warn('MSAL redirect', e); throw e; }
  account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
  if (account) pca.setActiveAccount(account);
  return account;
}

export async function signIn() {
  if (!pca) {
    if (!window.msal) throw new Error('Biblioteca de login não carregou (vendor/msal-browser.min.js). Recarregue a página com Ctrl+F5.');
    await initAuth();
    if (!pca) throw new Error('Login não inicializado (confira o clientId em config.js).');
  }
  return pca.loginRedirect({ scopes: CONFIG.scopes, prompt: 'select_account' });
}
export function signOut() { return pca.logoutRedirect({ account }); }

let testToken = null;
/** Gancho usado só nos testes automatizados. */
export function _useTestToken(t) { testToken = t; account = account || { username: 'teste@exemplo.com', name: 'Teste' }; }

async function token() {
  if (testToken) return testToken;
  if (!pca || !account) throw new Error('Não conectado à conta Microsoft');
  try {
    const r = await pca.acquireTokenSilent({ scopes: CONFIG.scopes, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|invalid_grant|consent_required/.test(e.errorCode || '')) {
      await pca.acquireTokenRedirect({ scopes: CONFIG.scopes, account });
      return new Promise(() => {}); // a página será redirecionada
    }
    throw e;
  }
}

async function gfetch(path, opts = {}) {
  const t = await token();
  const r = await fetch(path.startsWith('http') ? path : G + path, { ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) } });
  if (!r.ok) {
    const txt = await r.text().catch(() => '');
    const e = new Error(`OneDrive ${r.status}: ${txt.slice(0, 160)}`);
    e.status = r.status;
    throw e;
  }
  return r;
}

const seg = (s) => encodeURIComponent(s);
const pathOf = (...parts) => '/me/drive/root:/' + [CONFIG.folder, ...parts].map(seg).join('/');

async function putFile(parts, body, contentType) {
  const size = body.size ?? body.byteLength;
  if (size < 4 * 1024 * 1024) {
    await gfetch(pathOf(...parts) + ':/content', { method: 'PUT', headers: { 'Content-Type': contentType }, body });
    return;
  }
  // arquivos grandes: sessão de upload em blocos
  const s = await (await gfetch(pathOf(...parts) + ':/createUploadSession', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }),
  })).json();
  const buf = body instanceof Blob ? new Uint8Array(await body.arrayBuffer()) : body;
  const CH = 327680 * 10;
  for (let start = 0; start < buf.length; start += CH) {
    const end = Math.min(start + CH, buf.length);
    const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${buf.length}` }, body: buf.slice(start, end) });
    if (!r.ok && r.status !== 202) throw new Error('Falha no upload: ' + r.status);
  }
}

export function remoteFolderOf(v) {
  const d = new Date(v.createdAt);
  const ymd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  if (!v.remote) v.remote = `${baseName(v)} - ${ymd} - ${v.id.slice(0, 6)}`;
  v.remoteDesired = `${baseName(v)} - ${ymd} - ${v.id.slice(0, 6)}`;
  return v.remote;
}

/** Se ID do site / empreendimento mudaram, renomeia a pasta no OneDrive. */
async function renameIfNeeded(v) {
  remoteFolderOf(v);
  const want = v.remoteDesired; delete v.remoteDesired;
  if (want && want !== v.remote) {
    try {
      await gfetch(pathOf(v.remote), { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: want }) });
      v.remote = want;
    } catch (e) {
      if (e.status === 404) v.remote = want; // ainda não existia no OneDrive
      else console.warn('Renomear pasta', e);
    }
  }
}

const photoIds = (v) => {
  const ids = [];
  if (v.cover?.id) ids.push(v.cover.id);
  for (const g of v.groups || []) for (const p of g.photos || []) ids.push(p.id);
  return ids;
};

/** Envia a vistoria (fotos novas + JSON). */
export async function push(v) {
  await renameIfNeeded(v);
  const folder = v.remote;
  for (const id of photoIds(v)) {
    const rec = await db.getPhoto(id);
    if (rec && !rec.synced && rec.blob) {
      await putFile([folder, 'fotos', id + '.jpg'], rec.blob, 'image/jpeg');
      rec.synced = true;
      await db.putPhoto(rec);
    }
  }
  for (const id of v.removedPhotos || []) {
    try { await gfetch(pathOf(folder, 'fotos', id + '.jpg'), { method: 'DELETE' }); } catch (e) { if (e.status !== 404) throw e; }
  }
  v.removedPhotos = [];
  const copy = { ...v }; delete copy.dirty;
  await putFile([folder, 'vistoria.json'], new Blob([JSON.stringify(copy)], { type: 'application/json' }), 'application/json');
}

export async function savePptx(v, bytes, name) {
  const folder = remoteFolderOf(v);
  await putFile([folder, name], new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }), 'application/octet-stream');
}

async function listFolders() {
  try {
    const r = await (await gfetch(`/me/drive/root:/${seg(CONFIG.folder)}:/children?$select=name,folder&$top=200`)).json();
    return (r.value || []).filter((x) => x.folder).map((x) => x.name);
  } catch (e) {
    if (e.status === 404) return [];
    throw e;
  }
}

async function pull(folder, remoteV) {
  remoteV.remote = folder;
  for (const id of photoIds(remoteV)) {
    const have = await db.getPhoto(id);
    if (have) continue;
    try {
      const r = await gfetch(pathOf(folder, 'fotos', id + '.jpg') + ':/content');
      const blob = await r.blob();
      await db.putPhoto({ id, vid: remoteV.id, blob, synced: true });
    } catch (e) { if (e.status !== 404) throw e; }
  }
  remoteV.dirty = false;
  await db.putVistoria(remoteV);
}

/** Sincronização completa: envia o que está pendente e baixa o que veio de outros dispositivos. */
export async function syncAll() {
  const report = { pushed: 0, pulled: 0 };
  const local = await db.listVistorias();
  for (const v of local) {
    if (v.dirty) { await push(v); v.dirty = false; await db.putVistoria(v); report.pushed++; }
  }
  const byRemote = new Map((await db.listVistorias()).map((v) => [v.remote, v]));
  for (const folder of await listFolders()) {
    let remoteV;
    try { remoteV = await (await gfetch(pathOf(folder, 'vistoria.json') + ':/content')).json(); } catch (e) { if (e.status === 404) continue; throw e; }
    const mine = byRemote.get(folder) || (await db.getVistoria(remoteV.id));
    if (!mine || (!mine.dirty && (remoteV.updatedAt || '') > (mine.updatedAt || ''))) { await pull(folder, remoteV); report.pulled++; }
  }
  return report;
}

export async function removeRemote(v) {
  if (!v.remote) return;
  try { await gfetch(pathOf(v.remote), { method: 'DELETE' }); } catch (e) { if (e.status !== 404) throw e; }
}
