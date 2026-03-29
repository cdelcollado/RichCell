'use strict';

// ─── Translation Strings ──────────────────────────────────────────────────────

const TRANSLATIONS = {

  ca: {
    // Header
    activeCellLabel: 'Cel·la activa:',
    unsavedTitle:    'Hi ha canvis no guardats a Excel',
    // Toolbar tooltips
    headerTooltip:  'Capçalera',
    boldTooltip:    'Negreta (Ctrl+B)',
    italicTooltip:  'Cursiva (Ctrl+I)',
    bulletTooltip:  'Llista de punts',
    orderedTooltip: 'Llista numerada',
    linkTooltip:    'Inserir enllaç',
    cleanTooltip:   'Eliminar tot el format',
    // Preview sections
    renderPreviewTitle: 'Vista prèvia renderitzada',
    htmlPreviewTitle:   'Vista prèvia HTML',
    // Action buttons
    sendButton:       'Enviar a Excel',
    sendButtonTitle:  'Escriu el contingut HTML a la cel·la seleccionada (Ctrl+Enter)',
    sendingButton:    'Enviant\u2026',
    undoButton:       'Desfer',
    undoButtonTitle:  'Recupera el contingut original de la cel·la',
    clearButton:      'Netejar',
    clearButtonTitle: "Netejar l'editor",
    // Editor
    placeholder: 'Escriu el contingut aquí...',
    // Status messages
    requiresExcel:          'Aquest complement requereix Microsoft Excel.',
    selectSingleCellToLoad: 'Selecciona una única cel·la per editar-ne el contingut.',
    htmlLoaded:             'HTML carregat des de la cel·la.',
    textLoaded:             'Text carregat des de la cel·la.',
    editorEmpty:            "L'editor està buit. Afegeix contingut primer.",
    selectSingleCellToSend: 'Selecciona una única cel·la on escriure el contingut.',
    sentOk:                 'HTML enviat correctament a Excel! (Ctrl+Enter)',
    restored:               'Contingut restaurat.',
    cleared:                'Editor netejat.',
    empty:                  '(buit)',
    // Error messages
    officeError:     (code, msg) => `Error d'Office (${code}): ${msg}`,
    genericError:    (msg) => `Error: ${msg}`,
    unexpectedError: "S'ha produït un error inesperat.",
  },

  en: {
    activeCellLabel: 'Active cell:',
    unsavedTitle:    'Unsaved changes in Excel',
    headerTooltip:  'Heading',
    boldTooltip:    'Bold (Ctrl+B)',
    italicTooltip:  'Italic (Ctrl+I)',
    bulletTooltip:  'Bullet list',
    orderedTooltip: 'Numbered list',
    linkTooltip:    'Insert link',
    cleanTooltip:   'Remove all formatting',
    renderPreviewTitle: 'Rendered preview',
    htmlPreviewTitle:   'HTML preview',
    sendButton:       'Send to Excel',
    sendButtonTitle:  'Write HTML content to the selected cell (Ctrl+Enter)',
    sendingButton:    'Sending\u2026',
    undoButton:       'Undo',
    undoButtonTitle:  'Restore the original cell content',
    clearButton:      'Clear',
    clearButtonTitle: 'Clear the editor',
    placeholder: 'Write content here...',
    requiresExcel:          'This add-in requires Microsoft Excel.',
    selectSingleCellToLoad: 'Select a single cell to edit its content.',
    htmlLoaded:             'HTML loaded from cell.',
    textLoaded:             'Text loaded from cell.',
    editorEmpty:            'The editor is empty. Add content first.',
    selectSingleCellToSend: 'Select a single cell to write the content to.',
    sentOk:                 'HTML sent to Excel successfully! (Ctrl+Enter)',
    restored:               'Content restored.',
    cleared:                'Editor cleared.',
    empty:                  '(empty)',
    officeError:     (code, msg) => `Office error (${code}): ${msg}`,
    genericError:    (msg) => `Error: ${msg}`,
    unexpectedError: 'An unexpected error occurred.',
  },

  es: {
    activeCellLabel: 'Celda activa:',
    unsavedTitle:    'Hay cambios sin guardar en Excel',
    headerTooltip:  'Encabezado',
    boldTooltip:    'Negrita (Ctrl+B)',
    italicTooltip:  'Cursiva (Ctrl+I)',
    bulletTooltip:  'Lista de viñetas',
    orderedTooltip: 'Lista numerada',
    linkTooltip:    'Insertar enlace',
    cleanTooltip:   'Eliminar todo el formato',
    renderPreviewTitle: 'Vista previa renderizada',
    htmlPreviewTitle:   'Vista previa HTML',
    sendButton:       'Enviar a Excel',
    sendButtonTitle:  'Escribe el contenido HTML en la celda seleccionada (Ctrl+Enter)',
    sendingButton:    'Enviando\u2026',
    undoButton:       'Deshacer',
    undoButtonTitle:  'Restaura el contenido original de la celda',
    clearButton:      'Limpiar',
    clearButtonTitle: 'Limpiar el editor',
    placeholder: 'Escribe el contenido aquí...',
    requiresExcel:          'Este complemento requiere Microsoft Excel.',
    selectSingleCellToLoad: 'Selecciona una única celda para editar su contenido.',
    htmlLoaded:             'HTML cargado desde la celda.',
    textLoaded:             'Texto cargado desde la celda.',
    editorEmpty:            'El editor está vacío. Añade contenido primero.',
    selectSingleCellToSend: 'Selecciona una única celda donde escribir el contenido.',
    sentOk:                 '¡HTML enviado correctamente a Excel! (Ctrl+Enter)',
    restored:               'Contenido restaurado.',
    cleared:                'Editor limpiado.',
    empty:                  '(vacío)',
    officeError:     (code, msg) => `Error de Office (${code}): ${msg}`,
    genericError:    (msg) => `Error: ${msg}`,
    unexpectedError: 'Se ha producido un error inesperado.',
  },

};

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Returns the translation object that best matches the given Office display
 * language tag (BCP-47, e.g. "en-US", "es-ES", "ca-ES").
 * Supported languages: Catalan (ca), English (en), Spanish (es).
 * Falls back to English for any unsupported locale.
 * @param {string} displayLanguage - Value from Office.context.displayLanguage.
 * @returns {object} Flat translation map for the resolved language.
 */
function getTranslations(displayLanguage) {
  const lang = (displayLanguage || 'en-US').split('-')[0].toLowerCase();
  return TRANSLATIONS[lang] || TRANSLATIONS.en;
}
