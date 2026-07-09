/**
 * script.js — Lógica principal
 * Modelo Corporativo de Comunicación y Escalamiento
 *
 * Incluye persistencia automática en localStorage.
 * Sin frameworks. Sin dependencias externas.
 */

/* ============================================================
   CONSTANTES
   ============================================================ */
const STORAGE_KEY = 'MCE_draft_v1';

// Registro de aplicación en Azure AD — ver docs/GUIA_AZURE_AD.md
const AZURE_CLIENT_ID = "3371848c-ad5c-47b0-adea-a3d4b2abf908";
const AZURE_TENANT_ID = "eef48c79-e718-4ebd-b3c5-1c11368b1759";

// Ubicación fija del archivo Excel en SharePoint (ya resuelta, no hace falta cambiarla)
const EXCEL_DRIVE_ID = "b!oTWwotgO7UyfHzusUjAfA9GdzHlt3xpEoLT_SC9gC4zu8XqjUpHURLRHivHRu1no";
const EXCEL_ITEM_ID = "01DT6T2IXVMIIY5XX44FGIKWCBFKNE7GA5";

let msalInstance = null;
function getMsalInstance() {
  if (msalInstance) return msalInstance;
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: AZURE_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${AZURE_TENANT_ID}`,
      redirectUri: window.location.href.split('#')[0].split('?')[0],
    },
  });
  return msalInstance;
}

let lastTokenScopes = [];
let lastTokenAccount = "";

async function getGraphToken() {
  const instance = getMsalInstance();
  const scopes = ["Files.ReadWrite", "Sites.ReadWrite.All"];
  const accounts = instance.getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await instance.acquireTokenSilent({ scopes, account: accounts[0] });
      lastTokenScopes = result.scopes;
      lastTokenAccount = result.account ? result.account.username : "";
      return result.accessToken;
    } catch (e) {
      // el token silencioso falló (expiró, revocado, etc.) — se pide login de nuevo abajo
    }
  }
  const result = await instance.loginPopup({ scopes, prompt: "consent" });
  lastTokenScopes = result.scopes;
  lastTokenAccount = result.account ? result.account.username : "";
  return result.accessToken;
}

async function addRowsToTable(token, tableName, rows) {
  const url = `https://graph.microsoft.com/v1.0/drives/${EXCEL_DRIVE_ID}/items/${EXCEL_ITEM_ID}/workbook/tables/${tableName}/rows/add`;
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values: rows }),
  });
  if (!resp.ok) {
    const detail = await resp.text();
    throw new Error(`HTTP ${resp.status} al escribir en la tabla ${tableName}: ${detail}`);
  }
}

/* ============================================================
   ESTADO
   ============================================================ */
const state = {
  nombre: null,
  proceso: null,
  secciones: [],
  seccionActual: 0,
  respuestas: {},
  submitted: false,
};

/* ============================================================
   INIT
   ============================================================ */
document.addEventListener('DOMContentLoaded', () => {
  // Fecha de hoy
  const dateEl = document.getElementById('todayDate');
  if (dateEl) {
    dateEl.textContent = new Date().toLocaleDateString('es-CO', {
      day: '2-digit', month: 'long', year: 'numeric'
    });
  }

  renderProcessGrid();
  checkDraft();
});

/* ============================================================
   LOCALSTORAGE — AUTOSAVE
   ============================================================ */
function saveDraft() {
  if (state.submitted) return;
  try {
    const draft = {
      nombre: state.nombre,
      proceso: state.proceso,
      secciones: state.secciones,
      seccionActual: state.seccionActual,
      respuestas: state.respuestas,
      savedAt: new Date().toISOString(),
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(draft));
    updateAutosaveMsg(draft.savedAt);
  } catch (e) {
    // localStorage no disponible (modo privado muy restrictivo) — continúa sin guardar
  }
}

function loadDraft() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function clearDraftStorage() {
  try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* noop */ }
}

function updateAutosaveMsg(isoDate) {
  const el = document.getElementById('autosaveMsg');
  if (!el) return;
  const d = new Date(isoDate);
  el.textContent = `Guardado automáticamente a las ${d.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' })}`;
}

function showAutosaveBar(show) {
  const bar = document.getElementById('autosaveBar');
  if (bar) bar.classList.toggle('hidden', !show);
}

/* ============================================================
   DETECCIÓN Y RECUPERACIÓN DE BORRADOR
   ============================================================ */
function checkDraft() {
  const draft = loadDraft();
  if (!draft || !draft.proceso) return;

  const banner = document.getElementById('draftBanner');
  if (!banner) return;

  // Mostrar info del borrador
  const saved = new Date(draft.savedAt);
  const timeStr = saved.toLocaleDateString('es-CO', { day: '2-digit', month: 'short' })
    + ' a las ' + saved.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });

  const strong = banner.querySelector('strong');
  if (strong) strong.textContent = `Borrador de "${draft.proceso}" guardado el ${timeStr}`;

  banner.classList.remove('hidden');
}

function resumeDraft() {
  const draft = loadDraft();
  if (!draft) return;

  // Restaurar estado completo
  state.nombre = draft.nombre;
  state.proceso = draft.proceso;
  state.secciones = draft.secciones;
  state.seccionActual = draft.seccionActual;
  state.respuestas = draft.respuestas;
  document.getElementById('nombreInput').value = state.nombre || '';

  document.getElementById('draftBanner').classList.add('hidden');
  renderSection();
  goTo('screen-form');
  showAutosaveBar(true);
}

function clearDraft() {
  clearDraftStorage();
  document.getElementById('draftBanner').classList.add('hidden');
  showToast('Borrador eliminado. Comienza de nuevo.');
}

/* ============================================================
   NAVEGACIÓN
   ============================================================ */
function goTo(screenId) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(screenId).classList.add('active');
  window.scrollTo({ top: 0, behavior: 'smooth' });
  updateProgress(screenId);
}

function updateProgress(screenId) {
  const map = {
    'screen-welcome': 0,
    'screen-select': 15,
    'screen-form': 50,
    'screen-summary': 100,
  };
  document.getElementById('progressBar').style.width = (map[screenId] ?? 50) + '%';
}

/* ============================================================
   PANTALLA 2 — SELECCIÓN DE PROCESO
   ============================================================ */
function renderProcessGrid() {
  const grid = document.getElementById('processGrid');
  if (!grid) return;
  grid.innerHTML = '';

  PROCESOS.forEach(proc => {
    const sec = SECCIONES[proc];
    const card = document.createElement('div');
    card.className = 'process-card';
    card.dataset.proc = proc;
    card.innerHTML = `
      <div class="process-icon">${sec.icono}</div>
      <div>
        <div class="process-name">${proc}</div>
        <div class="process-hint">${sec.situaciones.length} situaciones</div>
      </div>
      <div class="checkmark">✓</div>
    `;
    card.addEventListener('click', () => selectProcess(proc));
    grid.appendChild(card);
  });
}

function onNombreInput(value) {
  state.nombre = value;
  updateStartFormButton();
}

function updateStartFormButton() {
  const nombreOk = !!(state.nombre && state.nombre.trim().length > 0);
  document.getElementById('btnStartForm').disabled = !(nombreOk && state.proceso);
}

function selectProcess(proc) {
  state.proceso = proc;
  document.querySelectorAll('.process-card').forEach(c => {
    c.classList.toggle('selected', c.dataset.proc === proc);
  });
  updateStartFormButton();
}

/* ============================================================
   INICIO DEL FORMULARIO
   ============================================================ */
function startForm() {
  if (!state.proceso) return;

  state.nombre = (state.nombre || '').trim();
  state.secciones = [state.proceso, TRANSVERSAL_KEY];
  state.seccionActual = 0;
  state.submitted = false;

  // Inicializar respuestas vacías
  state.respuestas = {};
  state.secciones.forEach(key => {
    const sec = SECCIONES[key];
    state.respuestas[key] = {
      situaciones: sec.situaciones.map(sit => ({
        situacion: sit,
        frecuencia: null,
        impacto: null,
        observaciones: ''
      })),
      preguntasAbiertas: sec.preguntasAbiertas.map(q => ({
        pregunta: q,
        respuesta: ''
      }))
    };
  });

  saveDraft();
  showAutosaveBar(true);
  renderSection();
  goTo('screen-form');
}

/* ============================================================
   PANTALLA 3 — RENDERIZAR SECCIÓN
   ============================================================ */
function renderSection() {
  const key = state.secciones[state.seccionActual];
  const sec = SECCIONES[key];
  const total = state.secciones.length;
  const current = state.seccionActual + 1;

  // Progress
  const pct = 15 + Math.round((current / (total + 1)) * 80);
  document.getElementById('progressBar').style.width = pct + '%';

  // Step indicator
  document.getElementById('formNav').innerHTML = `
    <div class="step-indicator">
      <span class="step-badge">Sección ${current} de ${total}</span>
      <span class="step-separator"></span>
      <span class="step-title">Proceso: <strong>${state.proceso}</strong></span>
    </div>
  `;

  // Table rows
  const resps = state.respuestas[key].situaciones;
  let rows = '';
  resps.forEach((r, idx) => { rows += buildRow(key, idx, r); });

  // Open questions
  let openQs = '';
  state.respuestas[key].preguntasAbiertas.forEach((qa, qIdx) => {
    openQs += `
      <div class="open-q-item">
        <label class="open-q-label" for="oq_${escKey(key)}_${qIdx}">${qa.pregunta}</label>
        <textarea class="open-q-input" id="oq_${escKey(key)}_${qIdx}"
          placeholder="Escribe tu respuesta aquí (opcional)…"
          oninput="saveOpenQ('${escKey(key)}', ${qIdx}, this.value)">${escHtml(qa.respuesta)}</textarea>
      </div>
    `;
  });

  document.getElementById('formContent').innerHTML = `
    <div class="section-block">
      <div class="section-header">
        <span class="icon">${sec.icono}</span>
        <div class="section-header-text">
          <h2>${sec.titulo}</h2>
          <p>${sec.situaciones.length} situaciones · Selecciona frecuencia e impacto para cada una</p>
        </div>
      </div>
      <div class="situations-table-wrap">
        <table class="situations-table">
          <thead>
            <tr>
              <th class="td-num">#</th>
              <th>Situación</th>
              <th>Frecuencia <span style="color:#d93025">*</span></th>
              <th>Impacto <span style="color:#d93025">*</span></th>
              <th>Observaciones</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <div class="open-questions">
        <div class="open-q-title">💬 Preguntas abiertas</div>
        ${openQs}
      </div>
    </div>
  `;

  // Buttons
  const isLast = state.seccionActual === total - 1;
  document.getElementById('formButtons').innerHTML = `
    <button class="btn btn-ghost" onclick="prevSection()">← Regresar</button>
    ${isLast
      ? `<button class="btn btn-primary" onclick="submitForm()">Enviar formulario ✔</button>`
      : `<button class="btn btn-primary" onclick="nextSection()">Continuar →</button>`
    }
  `;

  hideBanner();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function buildRow(key, idx, r) {
  const ek = escKey(key);
  const freqs = ['Alta', 'Media', 'Baja', 'No aplica'];
  const impacts = ['Alto', 'Medio', 'Bajo'];
  const isNA = r.frecuencia === 'No aplica';

  const freqRadios = freqs.map(f => `
    <label class="radio-label">
      <input type="radio" name="freq_${ek}_${idx}" value="${f}"
        ${r.frecuencia === f ? 'checked' : ''}
        onchange="saveFreq('${ek}', ${idx}, '${f}')">
      ${f}
    </label>
  `).join('');

  const impRadios = impacts.map(imp => `
    <label class="radio-label">
      <input type="radio" name="imp_${ek}_${idx}" value="${imp}"
        ${r.impacto === imp ? 'checked' : ''}
        onchange="saveImp('${ek}', ${idx}, '${imp}')">
      ${imp}
    </label>
  `).join('');

  return `
    <tr id="row_${ek}_${idx}" class="${isNA ? 'na-row' : ''}">
      <td class="td-num">${idx + 1}</td>
      <td class="td-situation">
        ${escHtml(r.situacion)}
        <div class="error-msg" id="err_${ek}_${idx}">Selecciona frecuencia e impacto</div>
      </td>
      <td><div class="radio-group">${freqRadios}</div></td>
      <td>
        <div class="radio-group" id="impGroup_${ek}_${idx}"
          style="${isNA ? 'opacity:.4;pointer-events:none' : ''}">
          ${impRadios}
        </div>
      </td>
      <td>
        <textarea class="obs-input" placeholder="Opcional…"
          oninput="saveObs('${ek}', ${idx}, this.value)">${escHtml(r.observaciones)}</textarea>
      </td>
    </tr>
  `;
}

/* ============================================================
   SAVE HANDLERS (con autosave tras cada cambio)
   ============================================================ */

// Mapea el ek (key escapada) de vuelta a la clave real del estado
function realKey(ek) {
  return state.secciones.find(k => escKey(k) === ek) || ek;
}

function saveFreq(ek, idx, value) {
  const key = realKey(ek);
  state.respuestas[key].situaciones[idx].frecuencia = value;

  const row = document.getElementById(`row_${ek}_${idx}`);
  const impGroup = document.getElementById(`impGroup_${ek}_${idx}`);

  if (value === 'No aplica') {
    row.classList.add('na-row');
    impGroup.style.opacity = '.4';
    impGroup.style.pointerEvents = 'none';
    state.respuestas[key].situaciones[idx].impacto = 'No aplica';
    document.querySelectorAll(`input[name="imp_${ek}_${idx}"]`).forEach(r => r.checked = false);
  } else {
    row.classList.remove('na-row');
    impGroup.style.opacity = '';
    impGroup.style.pointerEvents = '';
    if (state.respuestas[key].situaciones[idx].impacto === 'No aplica') {
      state.respuestas[key].situaciones[idx].impacto = null;
    }
  }

  clearRowError(ek, idx);
  saveDraft(); // ← autosave
}

function saveImp(ek, idx, value) {
  const key = realKey(ek);
  state.respuestas[key].situaciones[idx].impacto = value;
  clearRowError(ek, idx);
  saveDraft(); // ← autosave
}

function saveObs(ek, idx, value) {
  const key = realKey(ek);
  state.respuestas[key].situaciones[idx].observaciones = value.trim();
  saveDraft(); // ← autosave
}

function saveOpenQ(ek, qIdx, value) {
  const key = realKey(ek);
  state.respuestas[key].preguntasAbiertas[qIdx].respuesta = value.trim();
  saveDraft(); // ← autosave
}

/* ============================================================
   VALIDACIÓN
   ============================================================ */
function validateCurrentSection() {
  const key = state.secciones[state.seccionActual];
  const ek = escKey(key);
  const resps = state.respuestas[key].situaciones;
  let valid = true, errors = 0;

  resps.forEach((r, idx) => {
    const needsImp = r.frecuencia !== 'No aplica';
    const bad = !r.frecuencia || (needsImp && !r.impacto);
    setRowError(ek, idx, bad);
    if (bad) { valid = false; errors++; }
  });

  if (!valid) {
    showBanner(`Faltan ${errors} situación(es) por completar. Selecciona Frecuencia e Impacto.`);
  }
  return valid;
}

function setRowError(ek, idx, hasError) {
  const row = document.getElementById(`row_${ek}_${idx}`);
  if (row) row.classList.toggle('row-error', hasError);
}

function clearRowError(ek, idx) {
  const row = document.getElementById(`row_${ek}_${idx}`);
  if (row) row.classList.remove('row-error');
}

function showBanner(msg) {
  const banner = document.getElementById('validationBanner');
  document.getElementById('validationMsg').textContent = msg;
  banner.classList.remove('hidden');
  banner.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function hideBanner() {
  document.getElementById('validationBanner').classList.add('hidden');
}

/* ============================================================
   NAVEGACIÓN ENTRE SECCIONES
   ============================================================ */
function nextSection() {
  if (!validateCurrentSection()) return;
  state.seccionActual++;
  saveDraft();
  renderSection();
}

function prevSection() {
  if (state.seccionActual > 0) {
    state.seccionActual--;
    saveDraft();
    renderSection();
  } else {
    goTo('screen-select');
  }
}

/* ============================================================
   ENVÍO
   ============================================================ */
async function submitForm() {
  if (!validateCurrentSection()) return;
  state.submitted = true;
  clearDraftStorage(); // borrador ya no es necesario
  showAutosaveBar(false);
  renderSummary();
  goTo('screen-summary');
  document.getElementById('progressBar').style.width = '100%';
  await sendToExcel(buildExportData());
}

/* ============================================================
   PANTALLA 4 — RESUMEN
   ============================================================ */
function renderSummary() {
  document.getElementById('summarySubtitle').textContent =
    `Proceso: ${state.proceso} · ${new Date().toLocaleDateString('es-CO', {
      day: '2-digit', month: 'long', year: 'numeric'
    })}`;

  // Stats
  let totalSit = 0, totalAlta = 0, totalAlto = 0, totalNA = 0;
  state.secciones.forEach(key => {
    state.respuestas[key].situaciones.forEach(r => {
      totalSit++;
      if (r.frecuencia === 'Alta')     totalAlta++;
      if (r.impacto === 'Alto')        totalAlto++;
      if (r.frecuencia === 'No aplica') totalNA++;
    });
  });

  document.getElementById('summaryStats').innerHTML = `
    <div class="stat-card">
      <div class="stat-num">${totalSit}</div>
      <div class="stat-label">Situaciones evaluadas</div>
    </div>
    <div class="stat-card">
      <div class="stat-num" style="color:#b91c1c">${totalAlta}</div>
      <div class="stat-label">Frecuencia Alta</div>
    </div>
    <div class="stat-card">
      <div class="stat-num" style="color:#b91c1c">${totalAlto}</div>
      <div class="stat-label">Impacto Alto</div>
    </div>
    <div class="stat-card">
      <div class="stat-num" style="color:#6b6b6b">${totalNA}</div>
      <div class="stat-label">No aplican</div>
    </div>
  `;

  let html = '';
  state.secciones.forEach(key => {
    const sec = SECCIONES[key];
    const resps = state.respuestas[key];

    const rows = resps.situaciones.map((r, i) => {
      const fTag = tagClass(r.frecuencia);
      const iTag = tagClass(r.impacto);
      return `
        <tr>
          <td style="color:#888;font-size:12px;width:32px">${i + 1}</td>
          <td>${escHtml(r.situacion)}</td>
          <td><span class="tag ${fTag}">${r.frecuencia || '—'}</span></td>
          <td><span class="tag ${iTag}">${r.impacto || '—'}</span></td>
          <td style="font-size:13px;color:#555">${r.observaciones || '<span style="color:#bbb">—</span>'}</td>
        </tr>
      `;
    }).join('');

    const qaHtml = resps.preguntasAbiertas.map(qa => `
      <div class="summary-qa-item">
        <div class="summary-qa-q">💬 ${escHtml(qa.pregunta)}</div>
        <div class="summary-qa-a">${qa.respuesta
          ? `"${escHtml(qa.respuesta)}"`
          : '<span class="summary-qa-empty">Sin respuesta</span>'
        }</div>
      </div>
    `).join('');

    html += `
      <div class="summary-section">
        <div class="summary-section-title">${sec.icono} ${sec.titulo}</div>
        <table class="summary-table">
          <thead>
            <tr>
              <th>#</th><th>Situación</th><th>Frecuencia</th><th>Impacto</th><th>Observaciones</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
        <div class="summary-qa">${qaHtml}</div>
      </div>
    `;
  });

  document.getElementById('summaryContent').innerHTML = html;
}

/* ============================================================
   EXPORTAR
   ============================================================ */
function buildExportData() {
  const respuestas = [];
  const preguntasAbiertas = [];

  state.secciones.forEach(key => {
    const data = state.respuestas[key];
    data.situaciones.forEach(r => {
      respuestas.push({
        seccion: SECCIONES[key].titulo,
        situacion: r.situacion,
        frecuencia: r.frecuencia || '',
        impacto: r.impacto || '',
        observaciones: r.observaciones || ''
      });
    });
    data.preguntasAbiertas.forEach(qa => {
      preguntasAbiertas.push({
        seccion: SECCIONES[key].titulo,
        pregunta: qa.pregunta,
        respuesta: qa.respuesta || ''
      });
    });
  });

  return {
    nombre: state.nombre,
    proceso: state.proceso,
    fecha: new Date().toISOString(),
    respuestas,
    preguntasAbiertas
  };
}

function showSubmitStatus(kind, message) {
  const el = document.getElementById('submitStatus');
  el.className = `submit-status ${kind}`;
  el.textContent = message;
}

async function sendToExcel(exportData) {
  if (!AZURE_CLIENT_ID || AZURE_CLIENT_ID.includes('PON_AQUI') || !AZURE_TENANT_ID || AZURE_TENANT_ID.includes('PON_AQUI')) {
    showSubmitStatus('warning', '⚠️ Envío automático no configurado todavía. Descarga tu JSON o CSV abajo y compártelo con el equipo de Excelencia Corporativa.');
    return;
  }
  try {
    const token = await getGraphToken();

    const respuestasRows = exportData.respuestas.map(r => [
      exportData.nombre, exportData.proceso, exportData.fecha, r.seccion, r.situacion, r.frecuencia, r.impacto, r.observaciones
    ]);
    const preguntasRows = exportData.preguntasAbiertas.map(p => [
      exportData.nombre, exportData.proceso, exportData.fecha, p.seccion, p.pregunta, p.respuesta
    ]);

    await addRowsToTable(token, "Respuestas", respuestasRows);
    await addRowsToTable(token, "PreguntasAbiertas", preguntasRows);

    showSubmitStatus('success', '✅ Tus respuestas se guardaron automáticamente en Excel.');
  } catch (e) {
    const diag = ' [diag] cuenta: ' + lastTokenAccount + ' | scopes del token: ' + JSON.stringify(lastTokenScopes) + ' | drive: ' + EXCEL_DRIVE_ID + ' | item: ' + EXCEL_ITEM_ID;
    showSubmitStatus('warning', '⚠️ No se pudo enviar automáticamente (' + e.message + diag + '). Descarga tu JSON o CSV abajo y compártelo con el equipo de Excelencia Corporativa.');
  }
}

function downloadJSON() {
  const data = buildExportData();
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  triggerDownload(blob, `MCE_${slugify(state.proceso)}_${dateStamp()}.json`);
  showToast('JSON descargado ✓');
}

function downloadCSV() {
  const data = buildExportData();
  const BOM = '﻿'; // UTF-8 BOM para Excel

  const esc = v => {
    const s = String(v ?? '');
    return (s.includes(',') || s.includes('"') || s.includes('\n'))
      ? `"${s.replace(/"/g, '""')}"`
      : s;
  };

  let csv = BOM;
  csv += 'Proceso,Sección,Situación,Frecuencia,Impacto,Observaciones\n';
  data.respuestas.forEach(r => {
    csv += [data.proceso, r.seccion, r.situacion, r.frecuencia, r.impacto, r.observaciones]
      .map(esc).join(',') + '\n';
  });

  csv += '\n\nProceso,Sección,Pregunta,Respuesta\n';
  data.preguntasAbiertas.forEach(q => {
    csv += [data.proceso, q.seccion, q.pregunta, q.respuesta]
      .map(esc).join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  triggerDownload(blob, `MCE_${slugify(state.proceso)}_${dateStamp()}.csv`);
  showToast('CSV descargado ✓');
}

/* ============================================================
   REINICIAR
   ============================================================ */
function restartForm() {
  clearDraftStorage();
  state.nombre = null;
  state.proceso = null;
  state.secciones = [];
  state.seccionActual = 0;
  state.respuestas = {};
  state.submitted = false;
  document.querySelectorAll('.process-card').forEach(c => c.classList.remove('selected'));
  document.getElementById('nombreInput').value = '';
  document.getElementById('btnStartForm').disabled = true;
  document.getElementById('progressBar').style.width = '0%';
  showAutosaveBar(false);
  document.getElementById('draftBanner').classList.add('hidden');
  goTo('screen-welcome');
}

/* ============================================================
   UTILIDADES
   ============================================================ */
function escKey(k) {
  return k.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
}

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function tagClass(val) {
  if (!val) return 'tag-na';
  const map = {
    'Alta': 'tag-alta', 'Media': 'tag-media', 'Baja': 'tag-baja', 'No aplica': 'tag-noaplica',
    'Alto': 'tag-alto', 'Medio': 'tag-medio', 'Bajo': 'tag-bajo',
  };
  return map[val] || 'tag-na';
}

function slugify(str) {
  return str.normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
}

function dateStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

let toastTimer;
function showToast(msg) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
}
