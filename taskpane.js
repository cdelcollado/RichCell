/* global Office, Excel, Quill */

'use strict';

// ─── State ────────────────────────────────────────────────────────────────────

/** @type {Quill|null} The Quill editor instance. */
let quill = null;

/**
 * Guard flag that prevents the selection-change handler from re-triggering
 * while the editor is being updated programmatically.
 * @type {boolean}
 */
let isLoadingFromCell = false;

/** @type {number|null} Timer ID for auto-clearing the status bar message. */
let statusTimer = null;

/**
 * Stores the last HTML value loaded from the active cell, enabling the
 * "Undo" button to restore the cell's original content after an edit.
 * Null when no cell has been loaded yet.
 * @type {string|null}
 */
let lastCellSnapshot = null;

// ─── Office Initialization ────────────────────────────────────────────────────

/**
 * Entry point: called by Office.js once the host application is ready.
 * Initialises the editor, wires up all UI controls, and loads the
 * currently selected cell's content.
 */
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

/**
 * Initialises the Quill WYSIWYG editor with the snow theme and attaches
 * it to the #editor element. Registers H1/H2 headers and link support.
 * Updates the HTML preview and unsaved indicator on every keystroke.
 */
function initEditor() {
  quill = new Quill('#editor', {
    modules: {
      toolbar: {
        container: '#toolbar',
        handlers: {
          // Default link handler prompts for URL — no override needed
        }
      }
    },
    formats: ['bold', 'italic', 'list', 'link', 'header'],
    placeholder: 'Escriu el contingut aquí...',
    theme: 'snow'
  });

  quill.on('text-change', () => {
    updateHTMLPreview();
    updateRenderPreview();
    updateUnsavedIndicator();
  });
}

// ─── Button Bindings ──────────────────────────────────────────────────────────

/**
 * Attaches click handlers to all action buttons and registers the
 * Ctrl+Enter keyboard shortcut for sending content to Excel.
 */
function bindButtons() {
  document.getElementById('btn-send').addEventListener('click', sendToExcel);
  document.getElementById('btn-undo').addEventListener('click', undoToCell);
  document.getElementById('btn-clear').addEventListener('click', clearEditor);

  // Ctrl+Enter shortcut: send to Excel without reaching for the mouse
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'Enter') {
      e.preventDefault();
      sendToExcel();
    }
  });
}

// ─── Excel Selection Handler ──────────────────────────────────────────────────

/**
 * Registers a persistent workbook-level handler so that each time the user
 * selects a different cell the editor content is updated automatically.
 */
function registerSelectionHandler() {
  Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(onSelectionChanged);
    await context.sync();
  }).catch(handleError);
}

/**
 * Called by Excel whenever the active selection changes.
 * Opens a new Excel.run context because the event argument context is read-only.
 */
async function onSelectionChanged(/* event */) {
  if (isLoadingFromCell) return;
  await loadCurrentCellContent();
}

// ─── Load Cell Content Into Editor ───────────────────────────────────────────

/**
 * Reads the value of the currently selected cell and loads it into the
 * Quill editor. If the value looks like HTML it is set as innerHTML;
 * plain text is set via setText(). Also snapshots the loaded value so
 * that the Undo button can restore it later.
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
          quill.setText('');
          setSnapshot(null);
        } else {
          const strValue = String(rawValue);

          if (looksLikeHTML(strValue)) {
            quill.root.innerHTML = strValue;
            showStatus('HTML carregat des de la cel·la.', 'success');
          } else {
            quill.setText(strValue);
            showStatus('Text carregat des de la cel·la.', 'info');
          }

          setSnapshot(strValue);
        }
      } finally {
        isLoadingFromCell = false;
        updateHTMLPreview();
        updateRenderPreview();
        setUnsavedIndicator(false);
      }
    });
  } catch (error) {
    isLoadingFromCell = false;
    handleError(error);
  }
}

// ─── Send to Excel ────────────────────────────────────────────────────────────

/**
 * Sanitises the current editor content and writes it to the active cell
 * as a plain HTML string. Auto-fits the row height after writing.
 * Disables the send button while the async operation is in progress.
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

      setSnapshot(htmlContent);
      setUnsavedIndicator(false);
      showStatus('HTML enviat correctament a Excel! (Ctrl+Enter)', 'success');
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

// ─── Undo ─────────────────────────────────────────────────────────────────────

/**
 * Restores the editor to the snapshot value that was loaded from the cell
 * the last time it was read. Does nothing if no snapshot exists.
 */
function undoToCell() {
  if (lastCellSnapshot === null) return;

  isLoadingFromCell = true;
  try {
    if (looksLikeHTML(lastCellSnapshot)) {
      quill.root.innerHTML = lastCellSnapshot;
    } else {
      quill.setText(lastCellSnapshot);
    }
  } finally {
    isLoadingFromCell = false;
    updateHTMLPreview();
    updateRenderPreview();
    setUnsavedIndicator(false);
  }

  showStatus('Contingut restaurat.', 'info');
}

/**
 * Saves a snapshot of the current cell value and updates the Undo button state.
 * Pass null to indicate there is nothing to undo (e.g. empty cell).
 * @param {string|null} value - The cell value to snapshot, or null to clear.
 */
function setSnapshot(value) {
  lastCellSnapshot = value;
  const btn = document.getElementById('btn-undo');
  if (btn) btn.disabled = (value === null);
}

// ─── Clear Editor ─────────────────────────────────────────────────────────────

/**
 * Clears all content from the editor without modifying the Excel cell.
 * Does not affect the undo snapshot.
 */
function clearEditor() {
  quill.setText('');
  updateHTMLPreview();
  updateRenderPreview();
  showStatus('Editor netejat.', 'info');
}

// ─── HTML Utilities ───────────────────────────────────────────────────────────

/**
 * Returns sanitised HTML from the editor, stripping attributes that are
 * either Quill-internal (class, style, data-*) or security-sensitive
 * (on* event handlers, javascript: URIs).
 * Returns an empty string when the editor is blank.
 * @returns {string} Clean HTML string, or '' if the editor is empty.
 */
function buildCleanHTML() {
  let html = quill.root.innerHTML.trim();

  // Quill's empty state representations
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
 * Returns true if the given string contains at least one HTML tag.
 * Used as a heuristic to decide whether to load a cell value as HTML
 * or as plain text.
 * @param {string} str - The string to test.
 * @returns {boolean}
 */
function looksLikeHTML(str) {
  return /<[a-z][\s\S]*>/i.test(str);
}

// ─── HTML Preview ─────────────────────────────────────────────────────────────

/**
 * Updates the collapsible raw HTML preview panel with the current
 * sanitised editor content.
 */
function updateHTMLPreview() {
  const preview = document.getElementById('html-preview');
  if (!preview) return;
  const html = buildCleanHTML();
  preview.textContent = html || '(buit)';
}

// ─── Rendered Preview ─────────────────────────────────────────────────────────

/**
 * Updates the collapsible rendered preview panel by setting its innerHTML
 * to the sanitised editor content. This shows how the HTML will look
 * when rendered by a browser or e-commerce platform.
 */
function updateRenderPreview() {
  const container = document.getElementById('render-preview');
  if (!container) return;
  const html = buildCleanHTML();
  container.innerHTML = html || '<em style="color:#a19f9d">(buit)</em>';
}

// ─── Unsaved Indicator ────────────────────────────────────────────────────────

/**
 * Compares the current editor content with the last saved cell snapshot
 * and shows or hides the unsaved-changes indicator dot accordingly.
 */
function updateUnsavedIndicator() {
  if (isLoadingFromCell) return;
  const currentHTML = buildCleanHTML();
  const isDirty = lastCellSnapshot !== null
    ? currentHTML !== lastCellSnapshot
    : currentHTML !== '';
  setUnsavedIndicator(isDirty);
}

/**
 * Directly sets the visibility of the unsaved-changes indicator dot.
 * @param {boolean} dirty - True to show the indicator, false to hide it.
 */
function setUnsavedIndicator(dirty) {
  const el = document.getElementById('unsaved-indicator');
  if (!el) return;
  el.hidden = !dirty;
}

// ─── Cell Address Display ─────────────────────────────────────────────────────

/**
 * Updates the cell address badge in the header.
 * Strips the sheet name prefix (e.g. "Sheet1!B4" → "B4").
 * @param {string} fullAddress - The full cell address returned by the Excel API.
 */
function updateCellAddressDisplay(fullAddress) {
  const el = document.getElementById('cell-address');
  if (!el) return;
  const short = fullAddress.includes('!') ? fullAddress.split('!')[1] : fullAddress;
  el.textContent = short;
}

// ─── Status Bar ───────────────────────────────────────────────────────────────

/**
 * Displays a temporary message in the status bar and clears it after 4 seconds.
 * @param {string} message - The message to display.
 * @param {'success'|'error'|'warning'|'info'} type - Controls the colour of the badge.
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

/**
 * Logs an error to the browser console and shows a human-readable message
 * in the status bar. Handles both OfficeExtension errors (with .code) and
 * standard JavaScript errors (with .message).
 * @param {Error|Office.Error} error - The caught error object.
 */
function handleError(error) {
  console.error('[RichCell]', error);

  let msg;
  if (error && error.code) {
    msg = `Error d'Office (${error.code}): ${error.message}`;
  } else if (error && error.message) {
    msg = `Error: ${error.message}`;
  } else {
    msg = 'S\'ha produït un error inesperat.';
  }

  showStatus(msg, 'error');
}
