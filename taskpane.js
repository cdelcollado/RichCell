/* global Office, Excel, Quill */

'use strict';

// ─── State ────────────────────────────────────────────────────────────────────
let quill = null;
let isLoadingFromCell = false;  // Guard: avoid re-triggering on programmatic edits
let statusTimer = null;

// ─── Office Initialization ────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initEditor();
    bindButtons();
    registerSelectionHandler();
    loadCurrentCellContent();
  } else {
    showStatus('Aquest complement requereix Microsoft Excel.', 'error');
  }
});

// ─── Quill Editor Setup ───────────────────────────────────────────────────────
function initEditor() {
  quill = new Quill('#editor', {
    modules: {
      toolbar: {
        container: '#toolbar'
      }
    },
    placeholder: 'Escriu la descripció del producte aquí...',
    theme: 'snow'
  });

  // Update HTML preview on every content change
  quill.on('text-change', updateHTMLPreview);
}

// ─── Button Bindings ──────────────────────────────────────────────────────────
function bindButtons() {
  document.getElementById('btn-send').addEventListener('click', sendToExcel);
  document.getElementById('btn-clear').addEventListener('click', clearEditor);
}

// ─── Excel Selection Handler ──────────────────────────────────────────────────

/**
 * Registers a persistent handler so that every time the user clicks
 * a different cell the editor is updated automatically.
 */
function registerSelectionHandler() {
  Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(onSelectionChanged);
    await context.sync();
  }).catch(handleError);
}

/**
 * Called by Excel whenever the selection changes.
 * We open a NEW Excel.run context here (the event arg context is read-only).
 */
async function onSelectionChanged(/* event */) {
  if (isLoadingFromCell) return;
  await loadCurrentCellContent();
}

// ─── Load Cell Content Into Editor ───────────────────────────────────────────

/**
 * Reads the currently selected cell and loads its value into Quill
 * if it contains HTML (or plain text).
 */
async function loadCurrentCellContent() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'cellCount']);
      await context.sync();

      updateCellAddressDisplay(range.address);

      if (range.cellCount !== 1) {
        showStatus('Selecciona una única cel·la per editar-ne el contingut.', 'warning');
        return;
      }

      const rawValue = range.values[0][0];

      isLoadingFromCell = true;

      try {
        if (rawValue === null || rawValue === undefined || rawValue === '') {
          // Empty cell → clear editor silently
          quill.setText('');
        } else {
          const strValue = String(rawValue);

          if (looksLikeHTML(strValue)) {
            quill.root.innerHTML = strValue;
            showStatus('HTML carregat des de la cel·la.', 'success');
          } else {
            // Plain text: set as a paragraph
            quill.setText(strValue);
            showStatus('Text carregat des de la cel·la.', 'info');
          }
        }
      } finally {
        isLoadingFromCell = false;
        updateHTMLPreview();
      }
    });
  } catch (error) {
    isLoadingFromCell = false;
    handleError(error);
  }
}

// ─── Send to Excel ────────────────────────────────────────────────────────────

/**
 * Takes the current editor content, sanitises it and writes it to
 * the active cell as an HTML string.
 */
async function sendToExcel() {
  const htmlContent = buildCleanHTML();

  if (!htmlContent) {
    showStatus("L'editor està buit. Afegeix contingut primer.", 'warning');
    return;
  }

  const btn = document.getElementById('btn-send');

  try {
    btn.disabled = true;
    btn.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
           fill="none" stroke="currentColor" stroke-width="2.5">
        <circle cx="12" cy="12" r="10"/>
        <polyline points="12 6 12 12 16 14"/>
      </svg>
      Enviant…`;

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('cellCount');
      await context.sync();

      if (range.cellCount !== 1) {
        showStatus('Selecciona una única cel·la on escriure el contingut.', 'warning');
        return;
      }

      // Write as plain string (HTML stored as cell text)
      range.values = [[htmlContent]];

      // Auto-fit row height so the content is visible
      range.format.autofitRows();

      await context.sync();
      showStatus('HTML enviat correctament a Excel!', 'success');
    });
  } catch (error) {
    handleError(error);
  } finally {
    btn.disabled = false;
    btn.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24"
           fill="none" stroke="currentColor" stroke-width="2.5"
           stroke-linecap="round" stroke-linejoin="round">
        <path d="M22 2L11 13"/>
        <path d="M22 2L15 22 11 13 2 9l20-7z"/>
      </svg>
      Enviar a Excel`;
  }
}

// ─── Clear Editor ─────────────────────────────────────────────────────────────
function clearEditor() {
  quill.setText('');
  updateHTMLPreview();
  showStatus('Editor netejat.', 'info');
}

// ─── HTML Utilities ───────────────────────────────────────────────────────────

/**
 * Returns sanitised HTML from the editor, removing Quill's internal
 * class/style attributes that are not needed in the stored value.
 */
function buildCleanHTML() {
  let html = quill.root.innerHTML.trim();

  // Quill's empty state
  if (html === '<p><br></p>' || html === '<p></p>' || html === '') return '';

  // Remove Quill-specific attributes that add no semantic value
  html = html
    .replace(/\sclass="[^"]*"/g, '')
    .replace(/\sstyle="[^"]*"/g, '')
    .replace(/\sdata-[^=]+="[^"]*"/g, '');

  // Strip all event handler attributes (on* = potential XSS vector)
  html = html.replace(/\s+on\w+="[^"]*"/gi, '');

  // Strip javascript: URIs from href/src attributes
  html = html.replace(/(href|src)="javascript:[^"]*"/gi, '');

  // Collapse multiple whitespace sequences inside tags
  html = html.replace(/\s{2,}/g, ' ').trim();

  return html;
}

/**
 * Simple heuristic: does the string contain at least one HTML tag?
 */
function looksLikeHTML(str) {
  return /<[a-z][\s\S]*>/i.test(str);
}

// ─── HTML Preview ─────────────────────────────────────────────────────────────
function updateHTMLPreview() {
  const preview = document.getElementById('html-preview');
  if (!preview) return;
  const html = buildCleanHTML();
  preview.textContent = html || '(buit)';
}

// ─── Cell Address Display ─────────────────────────────────────────────────────
function updateCellAddressDisplay(fullAddress) {
  const el = document.getElementById('cell-address');
  if (!el) return;
  // Strip sheet name (e.g. "Sheet1!B4" → "B4")
  const short = fullAddress.includes('!') ? fullAddress.split('!')[1] : fullAddress;
  el.textContent = short;
}

// ─── Status Bar ───────────────────────────────────────────────────────────────

/**
 * @param {string} message
 * @param {'success'|'error'|'warning'|'info'} type
 */
function showStatus(message, type) {
  const el = document.getElementById('status-message');
  if (!el) return;

  el.textContent = message;
  el.className = `status-${type}`;

  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => {
    el.textContent = '';
    el.className = '';
  }, 4000);
}

// ─── Error Handler ────────────────────────────────────────────────────────────
function handleError(error) {
  console.error('[RichCell]', error);

  let msg;
  if (error && error.code) {
    // OfficeExtension.Error
    msg = `Error d'Office (${error.code}): ${error.message}`;
  } else if (error && error.message) {
    msg = `Error: ${error.message}`;
  } else {
    msg = 'S\'ha produït un error inesperat.';
  }

  showStatus(msg, 'error');
}
