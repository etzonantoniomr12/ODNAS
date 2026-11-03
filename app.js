'use strict';

// key de decisiones = item.key = String(número de fila), que coincide con lo que espera el backend
const state = { project: null, documents: [], review: null, index: 0, decisions: {} };
const $ = id => document.getElementById(id);
const views = ['project', 'documents', 'review', 'finish'];

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
  const lastId = localStorage.getItem('odnaLastProject');
  if (lastId) $('btn-last-project').classList.remove('hidden');
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(console.warn);
});

function bindEvents() {
  $('project-form').addEventListener('submit', createProject);
  $('btn-last-project').addEventListener('click', () => loadProject(localStorage.getItem('odnaLastProject')));
  $('btn-home').addEventListener('click', () => showView('project'));
  $('btn-back-docs').addEventListener('click', () => showView('documents'));
  $('btn-return-docs').addEventListener('click', () => showView('documents'));
  $('btn-drive').addEventListener('click', copyDriveLink);
  $('btn-retry').addEventListener('click', retryErrors);
  $('btn-prev').addEventListener('click', () => moveReview(-1));
  $('btn-next').addEventListener('click', () => moveReview(1));
  $('activity-select').addEventListener('change', saveCurrentDecision);
  $('btn-caution').addEventListener('click', toggleCaution);
  $('btn-apply').addEventListener('click', applyReview);
}

// ------------------------------------------------------------
//  API
// ------------------------------------------------------------
async function api(action, payload = {}) {
  const url = window.ODNA_CONFIG?.API_URL || '';
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec/.test(url)) {
    throw new Error('Configura la URL del Apps Script en config.js.');
  }
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload })
  });
  if (!response.ok) throw new Error(`Servidor HTTP ${response.status}`);
  const result = await response.json();
  if (!result.ok) throw new Error(result.error || 'El servidor no completó la solicitud.');
  return result.data;
}

// ------------------------------------------------------------
//  PROYECTO
// ------------------------------------------------------------
async function createProject(event) {
  event.preventDefault();
  setLoading(true, 'Creando proyecto', 'Localizando los Excel en la carpeta origen…');
  try {
    // FIX: envía sourceUrl + destinationUrl (no name)
    const project = await api('createProject', {
      sourceUrl:      $('source-url').value.trim(),
      destinationUrl: $('destination-url').value.trim()
    });
    state.project = project;
    localStorage.setItem('odnaLastProject', project.id);
    await processQueue();
    await loadProject(project.id);
  } catch (error) { toast(error.message, true); }
  finally { setLoading(false); }
}

async function processQueue() {
  let pending = state.project.pending || 0;
  while (pending > 0) {
    $('loader-text').textContent = `Procesando documento ${state.project.total - pending + 1} de ${state.project.total}…`;
    const result = await api('processNext', { projectId: state.project.id });
    state.project = result.project;
    pending = result.project.pending;
  }
}

async function loadProject(projectId) {
  if (!projectId) return;
  setLoading(true, 'Abriendo proyecto', 'Consultando documentos…');
  try {
    const data = await api('getProject', { projectId });
    state.project = data.project;
    state.documents = data.documents || [];
    renderDocuments();
    showView('documents');
  } catch (error) { toast(error.message, true); }
  finally { setLoading(false); }
}

async function retryErrors() {
  setLoading(true, 'Reintentando', 'Volviendo a procesar documentos con error…');
  try {
    await api('retryErrors', { projectId: state.project.id });
    await processQueue();
    await loadProject(state.project.id);
  } catch (error) { toast(error.message, true); }
  finally { setLoading(false); }
}

// ------------------------------------------------------------
//  LISTA DE DOCUMENTOS
// ------------------------------------------------------------
function renderDocuments() {
  $('project-title').textContent = state.project.name;
  const completed = state.documents.filter(d => d.status === 'DONE').length;
  const hasErrors = state.documents.some(d => d.status === 'ERROR');
  $('project-progress').textContent = `${completed}/${state.documents.length}`;
  $('btn-retry').classList.toggle('hidden', !hasErrors);

  $('document-list').innerHTML = state.documents.map(doc => `
    <article class="document-card">
      <div class="document-title"><span class="doc-icon">S</span><div><b>${escapeHtml(doc.name)}</b><small>${statusText(doc)}</small></div></div>
      <div class="document-actions">
        <button class="review-btn" data-review="${doc.outputId || ''}" ${doc.status !== 'DONE' ? 'disabled' : ''}>REVISAR ÍTEMS</button>
        <button class="pdf-btn" data-pdf="${doc.outputId || ''}" ${doc.status !== 'DONE' || !doc.applied ? 'disabled' : ''}>PDF</button>
      </div>
    </article>`).join('') || '<div class="card form-card">No existen documentos procesados.</div>';

  document.querySelectorAll('[data-review]').forEach(b => b.addEventListener('click', () => openReview(b.dataset.review)));
  document.querySelectorAll('[data-pdf]').forEach(b => b.addEventListener('click', () => generatePdf(b.dataset.pdf, b)));
}

function statusText(doc) {
  if (doc.status === 'DONE') return doc.applied ? 'Actividad generada' : 'Listo para revisar';
  if (doc.status === 'ERROR') return `Error: ${doc.error || 'no procesado'}`;
  return 'Pendiente';
}

// ------------------------------------------------------------
//  REVISIÓN DE ÍTEMS
// ------------------------------------------------------------
async function openReview(sheetId) {
  setLoading(true, 'Cargando ítems', 'Leyendo información y fotografías…');
  try {
    state.review = await api('getReview', { projectId: state.project.id, sheetId });
    state.index = 0;
    state.decisions = state.review.decisions || {};
    // FIX: activities llega como [{row, text}] → extraemos .text para el select
    fillActivities(state.review.activities || []);
    renderReview();
    showView('review');
  } catch (error) { toast(error.message, true); }
  finally { setLoading(false); }
}

function fillActivities(activities) {
  const separators = new Set([12, 22, 25, 28, 31, 34, 37, 41, 50]);
  const select = $('activity-select');
  select.innerHTML = '<option value="">Seleccione una actividad</option>';
  // FIX: cada activity es { row, text } — usamos .text como valor y etiqueta
  activities.forEach((activity, i) => {
    const text = typeof activity === 'object' ? activity.text : activity;
    const option = document.createElement('option');
    option.value = text;
    option.textContent = text;
    select.append(option);
    if (separators.has(i + 2)) {
      const divider = document.createElement('option');
      divider.disabled = true;
      divider.textContent = '────────────';
      select.append(divider);
    }
  });
}

function renderReview() {
  const item = state.review.items[state.index];
  // FIX: clave = item.key = String(fila), que coincide con el backend
  const decision = state.decisions[item.key] || { activity: '', caution: false };

  $('review-document').textContent = state.review.name;
  $('review-counter').textContent = `Ítem ${state.index + 1} de ${state.review.items.length}`;
  $('review-status').textContent = decision.caution ? 'Cautelar' : decision.activity ? 'Asignado' : 'Pendiente';
  $('item-data').innerHTML =
    field('NÚMERO DE ÍTEM', item.item) +
    field('PLAZO', item.deadline) +
    field('UBICACIÓN', `${item.startKm} – ${item.endKm}`, true) +
    field('VENCIMIENTO', item.dueDate, true) +
    field('DEFECTO', item.defect, true);

  const photoBox = $('photo-box');
  photoBox.innerHTML = item.photoUrl
    ? `<img src="${escapeAttr(item.photoUrl)}" alt="Fotografía del ítem ${escapeAttr(item.item)}">`
    : '<div class="photo-placeholder">Fotografía pendiente de extracción</div>';

  $('activity-select').value = decision.activity || '';
  $('activity-select').disabled = !!decision.caution;
  $('btn-caution').classList.toggle('active', !!decision.caution);
  $('btn-prev').disabled = state.index === 0;
  $('btn-next').textContent = state.index === state.review.items.length - 1 ? 'FINALIZAR' : 'SIGUIENTE';
}

function field(label, value, wide = false) {
  return `<div class="item-field ${wide ? 'wide' : ''}"><small>${label}</small><b>${escapeHtml(value ?? '')}</b></div>`;
}

function saveCurrentDecision() {
  const item = state.review.items[state.index];
  // FIX: clave = item.key
  const current = state.decisions[item.key] || {};
  state.decisions[item.key] = { activity: $('activity-select').value, caution: !!current.caution };
  renderReview();
}

function toggleCaution() {
  const item = state.review.items[state.index];
  // FIX: clave = item.key
  const current = state.decisions[item.key] || {};
  state.decisions[item.key] = { activity: '', caution: !current.caution };
  renderReview();
}

function moveReview(delta) {
  saveCurrentDecision();
  if (delta > 0 && !validDecision()) return toast('Selecciona una actividad o MEDIDA CAUTELAR.', true);
  const next = state.index + delta;
  if (next >= state.review.items.length) return showView('finish');
  if (next < 0) return;
  state.index = next;
  renderReview();
}

function validDecision() {
  const item = state.review.items[state.index];
  // FIX: clave = item.key
  const d = state.decisions[item.key] || {};
  return !!d.activity || !!d.caution;
}

async function applyReview() {
  // FIX: verifica usando item.key
  const missing = state.review.items.filter(i => {
    const d = state.decisions[i.key] || {};
    return !d.activity && !d.caution;
  });
  if (missing.length) return toast(`Faltan ${missing.length} ítem(s) por revisar.`, true);

  setLoading(true, 'Generando actividad ofertada', 'Aplicando textos, colores y alturas…');
  try {
    await api('applyReview', {
      projectId: state.project.id,
      sheetId:   state.review.sheetId,
      decisions: state.decisions   // ya tiene claves por número de fila
    });
    toast('Sheet actualizado correctamente.');
    await loadProject(state.project.id);
  } catch (error) { toast(error.message, true); }
  finally { setLoading(false); }
}

// ------------------------------------------------------------
//  PDF Y DRIVE
// ------------------------------------------------------------
async function generatePdf(sheetId, button) {
  button.disabled = true;
  setLoading(true, 'Generando PDF', 'Exportando todas las hojas visibles…');
  try {
    await api('generatePdf', { projectId: state.project.id, sheetId });
    toast('PDF generado en la carpeta del proyecto.');
  } catch (error) { toast(error.message, true); }
  finally { button.disabled = false; setLoading(false); }
}

async function copyDriveLink() {
  try { await navigator.clipboard.writeText(state.project.folderUrl); toast('Enlace de Drive copiado.'); }
  catch { window.open(state.project.folderUrl, '_blank', 'noopener'); }
}

// ------------------------------------------------------------
//  UTILIDADES
// ------------------------------------------------------------
function showView(name) {
  views.forEach(v => $(`view-${v}`).classList.toggle('active', v === name));
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
function setLoading(on, title, text) {
  $('loader').classList.toggle('hidden', !on);
  if (title) $('loader-title').textContent = title;
  if (text)  $('loader-text').textContent  = text;
}
function toast(message, error = false) {
  const el = $('toast');
  el.textContent = message;
  el.className = `toast${error ? ' error' : ''}`;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => el.classList.add('hidden'), 3600);
}
function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;' }[c]));
}
function escapeAttr(value) { return escapeHtml(value); }
