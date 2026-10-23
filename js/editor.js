// Editor de imagem: setas, formas, desenho livre e texto sobre a foto (tela cheia, toque/mouse).
const COLORS = ['#E53935', '#FDD835', '#43A047', '#1E88E5', '#FFFFFF', '#111111'];
const TOOLS = [['arrow', '➚', 'Seta'], ['rect', '▭', 'Retângulo'], ['ellipse', '◯', 'Círculo'], ['pen', '✎', 'Livre'], ['text', 'T', 'Texto']];
const THICK = [['S', 1], ['M', 1.8], ['G', 3]];

async function bitmapOf(blob) {
  try { return await createImageBitmap(blob); } catch (e) { /* fallback */ }
  const i = new Image(); i.src = URL.createObjectURL(blob); await i.decode(); return i;
}

function drawShape(ctx, s) {
  ctx.strokeStyle = s.color; ctx.fillStyle = s.color; ctx.lineWidth = s.lw; ctx.lineCap = 'round'; ctx.lineJoin = 'round';
  if (s.type === 'arrow') {
    const a = Math.atan2(s.y2 - s.y1, s.x2 - s.x1), hl = Math.max(s.lw * 4.5, 14);
    const bx = s.x2 - Math.cos(a) * hl * 0.8, by = s.y2 - Math.sin(a) * hl * 0.8;
    ctx.beginPath(); ctx.moveTo(s.x1, s.y1); ctx.lineTo(bx, by); ctx.stroke();
    ctx.beginPath(); ctx.moveTo(s.x2, s.y2);
    ctx.lineTo(s.x2 - Math.cos(a - 0.45) * hl, s.y2 - Math.sin(a - 0.45) * hl);
    ctx.lineTo(s.x2 - Math.cos(a + 0.45) * hl, s.y2 - Math.sin(a + 0.45) * hl);
    ctx.closePath(); ctx.fill();
  } else if (s.type === 'rect') {
    ctx.strokeRect(Math.min(s.x1, s.x2), Math.min(s.y1, s.y2), Math.abs(s.x2 - s.x1), Math.abs(s.y2 - s.y1));
  } else if (s.type === 'ellipse') {
    ctx.beginPath(); ctx.ellipse((s.x1 + s.x2) / 2, (s.y1 + s.y2) / 2, Math.abs(s.x2 - s.x1) / 2, Math.abs(s.y2 - s.y1) / 2, 0, 0, Math.PI * 2); ctx.stroke();
  } else if (s.type === 'pen') {
    ctx.beginPath(); s.pts.forEach((p, i) => (i ? ctx.lineTo(p.x, p.y) : ctx.moveTo(p.x, p.y)));
    if (s.pts.length === 1) ctx.lineTo(s.pts[0].x + 0.1, s.pts[0].y); ctx.stroke();
  } else if (s.type === 'text') {
    ctx.font = `bold ${s.size}px "Encode Sans", system-ui, sans-serif`; ctx.textBaseline = 'top'; ctx.lineWidth = Math.max(2, s.size / 6);
    ctx.strokeStyle = s.color === '#111111' ? '#FFFFFF' : '#000000';
    s.text.split('\n').forEach((l, i) => { ctx.strokeText(l, s.x, s.y + i * s.size * 1.15); ctx.fillText(l, s.x, s.y + i * s.size * 1.15); });
  }
}

/** Abre o editor. Resolve { blob, w, h } ao salvar, ou null se cancelar. */
export async function openEditor(blob, { orig = null, askText } = {}) {
  let base = await bitmapOf(blob);
  let W = base.width || base.naturalWidth, H = base.height || base.naturalHeight;
  const st = { tool: 'arrow', color: COLORS[0], thick: 1, shapes: [], cur: null };

  const root = document.createElement('div'); root.className = 'ed';
  root.innerHTML = `
    <div class="ed-bar"><button class="btn ghost small ed-cancel">Cancelar</button><b>Editar foto</b><button class="btn primary small ed-save">Salvar</button></div>
    <div class="ed-stage"><canvas></canvas></div>
    <div class="ed-tools">
      <div class="ed-row">${TOOLS.map(([k, ic, lb]) => `<button type="button" class="edb" data-tool="${k}"><span>${ic}</span><small>${lb}</small></button>`).join('')}</div>
      <div class="ed-row">
        ${COLORS.map((c) => `<button type="button" class="edc" data-color="${c}" style="background:${c}" aria-label="Cor"></button>`).join('')}
        <span class="ed-sep"></span>
        ${THICK.map(([l, v], i) => `<button type="button" class="edb sm" data-thick="${i}">${l}</button>`).join('')}
      </div>
      <div class="ed-row">
        <button type="button" class="edb wide" data-undo>↶ Desfazer</button>
        <button type="button" class="edb wide" data-clear>Limpar</button>
        ${orig ? '<button type="button" class="edb wide" data-orig>Original</button>' : ''}
      </div>
    </div>`;
  document.body.appendChild(root); document.body.classList.add('ed-open');
  const cv = root.querySelector('canvas'); const ctx = cv.getContext('2d');
  const stage = root.querySelector('.ed-stage');

  const lw = () => THICK[st.thick][1] * Math.max(W, H) / 380;
  function draw(extra) {
    cv.width = W; cv.height = H;
    ctx.drawImage(base, 0, 0, W, H);
    for (const s of st.shapes) drawShape(ctx, s);
    if (extra) drawShape(ctx, extra);
  }
  function fitCanvas() {
    const r = stage.getBoundingClientRect(); const k = Math.min((r.width - 12) / W, (r.height - 12) / H);
    cv.style.width = Math.max(50, W * k) + 'px'; cv.style.height = Math.max(50, H * k) + 'px';
  }
  const ro = new ResizeObserver(fitCanvas); ro.observe(stage);
  const sync = () => {
    root.querySelectorAll('[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === st.tool));
    root.querySelectorAll('[data-color]').forEach((b) => b.classList.toggle('on', b.dataset.color === st.color));
    root.querySelectorAll('[data-thick]').forEach((b) => b.classList.toggle('on', +b.dataset.thick === st.thick));
  };
  draw(); fitCanvas(); sync();

  const pt = (e) => { const r = cv.getBoundingClientRect(); return { x: (e.clientX - r.left) * W / r.width, y: (e.clientY - r.top) * H / r.height }; };

  return new Promise((resolve) => {
    const done = (v) => { ro.disconnect(); root.remove(); document.body.classList.remove('ed-open'); resolve(v); };

    cv.addEventListener('pointerdown', (e) => {
      e.preventDefault(); const p = pt(e);
      if (st.tool === 'text') {
        askText?.('Texto na imagem').then((t) => { if (t) { st.shapes.push({ type: 'text', x: p.x, y: p.y, text: t, color: st.color, size: Math.round(Math.max(W, H) / 26 * THICK[st.thick][1] * 0.8 + 8) }); draw(); } });
        return;
      }
      cv.setPointerCapture(e.pointerId);
      st.cur = { type: st.tool, color: st.color, lw: lw(), x1: p.x, y1: p.y, x2: p.x, y2: p.y, pts: [p] };
    });
    cv.addEventListener('pointermove', (e) => {
      if (!st.cur) return; const p = pt(e);
      st.cur.x2 = p.x; st.cur.y2 = p.y; if (st.cur.type === 'pen') st.cur.pts.push(p);
      draw(st.cur);
    });
    const end = () => {
      const c = st.cur; if (!c) return; st.cur = null;
      const moved = Math.hypot(c.x2 - c.x1, c.y2 - c.y1) > 6 || (c.type === 'pen' && c.pts.length > 2);
      if (moved) st.shapes.push(c);
      draw();
    };
    cv.addEventListener('pointerup', end); cv.addEventListener('pointercancel', end);

    root.addEventListener('click', async (e) => {
      const b = e.target.closest('button'); if (!b) return;
      if (b.dataset.tool) { st.tool = b.dataset.tool; sync(); }
      else if (b.dataset.color) { st.color = b.dataset.color; sync(); }
      else if (b.dataset.thick !== undefined) { st.thick = +b.dataset.thick; sync(); }
      else if (b.hasAttribute('data-undo')) { st.shapes.pop(); draw(); }
      else if (b.hasAttribute('data-clear')) { if (st.shapes.length && confirm('Apagar todas as marcações?')) { st.shapes = []; draw(); } }
      else if (b.hasAttribute('data-orig')) {
        if (confirm('Voltar para a foto original (sem nenhuma edição)?')) { base = await bitmapOf(orig); W = base.width || base.naturalWidth; H = base.height || base.naturalHeight; st.shapes = []; draw(); fitCanvas(); }
      }
      else if (b.classList.contains('ed-cancel')) { if (st.shapes.length && !confirm('Descartar as marcações?')) return; done(null); }
      else if (b.classList.contains('ed-save')) {
        draw(); cv.toBlob((out) => done({ blob: out, w: W, h: H }), 'image/jpeg', 0.88);
      }
    });
  });
}
