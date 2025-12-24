const ATTR_TABLE_JSON = 'tbljson';
if (typeof window !== 'undefined') {
  if (window.__epDataTablesLoaded) {
    console.debug('[ep_data_tables] Duplicate client_hooks.js load suppressed');
    return;
  }
  window.__epDataTablesLoaded = true;
}
const ATTR_CELL       = 'td';
const ATTR_CLASS_PREFIX = 'tbljson-';
const log             = (...m) => console.debug('[ep_data_tables:client_hooks]', ...m);
const DELIMITER       = '\u241F';
const HIDDEN_DELIM    = DELIMITER;
const INPUTTYPE_REPLACEMENT_TYPES = new Set(['insertReplacementText', 'insertFromComposition']);

const rand = () => Math.random().toString(36).slice(2, 8);

const enc = s => btoa(s).replace(/\+/g, '-').replace(/\//g, '_');
const dec = s => {
    const str = s.replace(/-/g, '+').replace(/_/g, '/');
    try {
        if (typeof atob === 'function') {
            return atob(str); 
        } else if (typeof Buffer === 'function') {
            return Buffer.from(str, 'base64').toString('utf8');
        } else {
            console.error('[ep_data_tables] Base64 decoding function (atob or Buffer) not found.');
            return null;
        }
    } catch (e) {
        console.error('[ep_data_tables] Error decoding base64 string:', s, e);
        return null;
    }
};

let lastClickedCellInfo = null;
let isResizing = false;
let resizeStartX = 0;
let resizeCurrentX = 0;
let resizeTargetTable = null;
let resizeTargetColumn = -1;
let resizeOriginalWidths = [];
let resizeTableMetadata = null;
let resizeLineNum = -1;
let resizeOverlay = null;
let suppressNextBeforeInputInsertTextOnce = false;
let isAndroidChromeComposition = false;
let handledCurrentComposition = false;
let suppressBeforeInputInsertTextDuringComposition = false;
let EP_DT_EDITOR_INFO = null;
const __epDT_postWriteScheduled = new Set();
// Composition tracking for corruption recovery
let __epDT_compositionOriginalLine = { tblId: null, lineNum: null, timestamp: 0 };
let __epDT_columnOperationInProgress = false;
let __epDT_compositionActive = false;
let __epDT_lastCompositionEndTime = 0;
const __epDT_COMPOSITION_COOLDOWN_MS = 500;

const isInCompositionCooldown = () => {
  if (__epDT_compositionActive) return true;
  if (__epDT_lastCompositionEndTime > 0) {
    const elapsed = Date.now() - __epDT_lastCompositionEndTime;
    if (elapsed < __epDT_COMPOSITION_COOLDOWN_MS) return true;
  }
  return false;
};

// DOM desync detection (browser extensions like Grammarly can corrupt state)
let __epDT_domDesynced = false;
let __epDT_desyncErrorCount = 0;
const __epDT_MAX_DESYNC_ERRORS = 3;

// Styling cache: preserves character-level attrs when emitting canonical text
const __epDT_pendingStylingForLine = new Map();
const __epDT_STYLING_CACHE_TTL_MS = 5000; // Expire after 5 seconds

const cacheStylingForLine = (lineId, styling, cells) => {
  if (!lineId || !styling || styling.length === 0) return;
  const now = Date.now();
  for (const [k, v] of __epDT_pendingStylingForLine.entries()) {
    if (now - v.timestamp > __epDT_STYLING_CACHE_TTL_MS) __epDT_pendingStylingForLine.delete(k);
  }
  __epDT_pendingStylingForLine.set(lineId, { styling, cells: cells ? cells.slice() : [], timestamp: now });
  console.debug('[ep_data_tables:stylingCache] cached', { lineId: lineId.slice(0, 20), stylingCount: styling.length, cellCount: cells?.length || 0 });
};

const getCachedStyling = (lineId) => {
  if (!lineId) return null;
  const cached = __epDT_pendingStylingForLine.get(lineId);
  if (!cached) return null;
  if (Date.now() - cached.timestamp > __epDT_STYLING_CACHE_TTL_MS) {
    __epDT_pendingStylingForLine.delete(lineId);
    return null;
  }
  return cached;
};

const clearCachedStyling = (lineId) => {
  if (lineId) __epDT_pendingStylingForLine.delete(lineId);
};

const isDestructiveOperationSafe = (operation) => {
  if (__epDT_domDesynced) {
    console.warn(`[ep_data_tables:SAFE_MODE] Blocking destructive operation: ${operation} (DOM desynced by extension)`);
    return false;
  }
  return true;
};

// Handle keyToNodeMap errors from browser extensions (Grammarly etc.)
const handleDomDesyncError = (error, source) => {
  const errorStr = String(error?.message || error || '');
  const stackStr = String(error?.stack || '');
  if (errorStr.includes('entry') || stackStr.includes('atKey') || stackStr.includes('keyToNodeMap')) {
    __epDT_desyncErrorCount++;
    console.error(`[ep_data_tables:DOM_DESYNC] keyToNodeMap error (${__epDT_desyncErrorCount}/${__epDT_MAX_DESYNC_ERRORS})`, { source, error: errorStr.slice(0, 200) });
    if (__epDT_desyncErrorCount >= __epDT_MAX_DESYNC_ERRORS) {
      __epDT_domDesynced = true;
      console.error('[ep_data_tables:DOM_DESYNC] Entering SAFE MODE');
      try {
        if (typeof $.gritter !== 'undefined' && $.gritter.add) {
          $.gritter.add({ title: 'Table Editor: Safe Mode', text: 'A browser extension is interfering. Some operations disabled.', sticky: true, class_name: 'gritter-warning' });
        }
      } catch (_) {}
    }
    return true;
  }
  return false;
};

// Global error handler to detect keyToNodeMap errors
if (typeof window !== 'undefined' && !window.__epDT_errorHandlerInstalled) {
  window.__epDT_errorHandlerInstalled = true;
  const originalOnError = window.onerror;
  window.onerror = function(message, source, lineno, colno, error) {
    handleDomDesyncError(error || message, 'window.onerror');
    if (originalOnError) return originalOnError.call(this, message, source, lineno, colno, error);
    return false;
  };
  window.addEventListener('unhandledrejection', (event) => handleDomDesyncError(event.reason, 'unhandledrejection'));
}

// Validate line exists, has valid lineNode, and node is in DOM (prevents keyToNodeMap errors)
const isLineSafeToModify = (rep, lineNum, logPrefix = '') => {
  try {
    if (!rep || !rep.lines) { if (logPrefix) console.debug(logPrefix, 'rep invalid'); return false; }
    const lineCount = rep.lines.length();
    if (lineNum < 0 || lineNum >= lineCount) { if (logPrefix) console.debug(logPrefix, 'line out of bounds', { lineNum, lineCount }); return false; }
    const lineEntry = rep.lines.atIndex(lineNum);
    if (!lineEntry) { if (logPrefix) console.debug(logPrefix, 'no line entry', { lineNum }); return false; }
    if (!lineEntry.lineNode) { if (logPrefix) console.debug(logPrefix, 'lineNode missing', { lineNum }); return false; }
    if (!lineEntry.lineNode.parentNode) { if (logPrefix) console.debug(logPrefix, 'lineNode orphaned', { lineNum }); return false; }
    if (typeof lineEntry.lineNode.isConnected === 'boolean' && !lineEntry.lineNode.isConnected) {
      if (logPrefix) console.debug(logPrefix, 'lineNode not connected', { lineNum }); return false;
    }
    
    return true;
  } catch (err) {
    if (logPrefix) console.debug(logPrefix, 'validation error', err?.message);
    return false;
  }
};

function isAndroidUA() {
  const ua = (navigator.userAgent || '').toLowerCase();
  const isAndroid = ua.includes('android');
  // Treat Chrome OS (Chromebooks) similarly because touch-screen Chromebooks exhibit the same
  // duplicate beforeinput / composition quirks we patch for Android.
  const isChromeOS = ua.includes('cros'); // "CrOS" token present in Chromebook UAs
  const isIOS = ua.includes('iphone') || ua.includes('ipad') || ua.includes('ipod') || ua.includes('crios');
  return (isAndroid || isChromeOS) && !isIOS;
}

function isIOSUA() {
  const ua = (navigator.userAgent || '').toLowerCase();
  return ua.includes('iphone') || ua.includes('ipad') || ua.includes('ipod') || ua.includes('ios') || ua.includes('crios') || ua.includes('fxios') || ua.includes('edgios');
}

// Normalise soft whitespace (NBSP, newline, carriage return, tab) to ASCII space and collapse runs.
const normalizeSoftWhitespace = (str) => (
  (str || '')
    .replace(/[\u00A0\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
);

const extractCellPlainText = (td) => {
  if (!td) return '';
  let text = '';
  try {
    text = td.textContent || '';
  } catch (_) {
    text = '';
  }
  text = (text || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  // Ensure each cell has at least a space so delimiters remain navigable.
  if (text.length === 0) text = ' ';
  return text;
};

/**
 * @param {HTMLElement} element 
 * @returns {HTMLElement|null} 
 */
function findTbljsonElement(element) {
  if (!element) return null;
  if (element.classList) {
    for (const cls of element.classList) {
      if (cls.startsWith(ATTR_CLASS_PREFIX)) {
        return element;
      }
    }
  }
  if (element.children) {
    for (const child of element.children) {
      const found = findTbljsonElement(child);
      if (found) return found;
    }
  }
  return null;
}

/**
 * Gets the table metadata for a given line, falling back to a DOM search if the
 * line attribute is not present (e.g., for block-styled lines).
 * @param {number} lineNum - The line number.
 * @param {object} editorInfo - The editor instance.
 * @param {object} docManager - The document attribute manager.
 * @returns {object|null} - The parsed metadata object or null.
 */
function getTableLineMetadata(lineNum, editorInfo, docManager) {
  const funcName = 'getTableLineMetadata';
  try {
    const attribs = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
    if (attribs) {
      try {
      const metadata = JSON.parse(attribs);
        if (metadata && metadata.tblId) {
          return metadata;
        }
      } catch (e) {
      }
    }

    const rep = editorInfo.ace_getRep();

    const lineEntry = rep.lines.atIndex(lineNum);
    const lineNode = lineEntry?.lineNode;

    if (!lineNode) {
      return null;
    }

    const tbljsonElement = findTbljsonElement(lineNode);
    if (tbljsonElement) {
      for (const className of tbljsonElement.classList) {
        if (className.startsWith(ATTR_CLASS_PREFIX)) {
          const encodedData = className.substring(ATTR_CLASS_PREFIX.length);
          try {
            const decodedString = atob(encodedData);
            const metadata = JSON.parse(decodedString);
            return metadata;
          } catch (e) {
            console.error(`${funcName}: Failed to decode/parse tbljson class on line ${lineNum}:`, e);
            return null;
          }
        }
      }
    }

    return null;
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error getting metadata for line ${lineNum}:`, e);
    return null;
  }
}

/**
 * Navigate to the next cell in the table (Tab key behavior)
 * @param {number} currentLineNum - Current line number
 * @param {number} currentCellIndex - Current cell index (0-based)
 * @param {object} tableMetadata - Table metadata object
 * @param {boolean} shiftKey - Whether Shift was held (for reverse navigation)
 * @param {object} editorInfo - Editor instance
 * @param {object} docManager - Document attribute manager
 * @returns {boolean} - Success of navigation
 */
function navigateToNextCell(currentLineNum, currentCellIndex, tableMetadata, shiftKey, editorInfo, docManager) {
  const funcName = 'navigateToNextCell';

  try {
  let targetRow = tableMetadata.row;
    let targetCol = currentCellIndex;

    if (shiftKey) {
      targetCol--;
  if (targetCol < 0) {
    targetRow--;
    targetCol = tableMetadata.cols - 1;
      }
    } else {
      targetCol++;
      if (targetCol >= tableMetadata.cols) {
    targetRow++;
    targetCol = 0;
      }
  }


  const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);
    if (targetLineNum === -1) {
      return false;
    }

  return navigateToCell(targetLineNum, targetCol, editorInfo, docManager);

  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during navigation:`, e);
    return false;
  }
}

/**
 * Navigate to the cell below in the same column (Enter key behavior)
 * @param {number} currentLineNum - Current line number
 * @param {number} currentCellIndex - Current cell index (0-based)
 * @param {object} tableMetadata - Table metadata object
 * @param {object} editorInfo - Editor instance
 * @param {object} docManager - Document attribute manager
 * @returns {boolean} - Success of navigation
 */
function navigateToCellBelow(currentLineNum, currentCellIndex, tableMetadata, editorInfo, docManager) {
  const funcName = 'navigateToCellBelow';

  try {
  const targetRow = tableMetadata.row + 1;
    const targetCol = currentCellIndex;


  const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);

  if (targetLineNum !== -1) {
      return navigateToCell(targetLineNum, targetCol, editorInfo, docManager);
    } else {
  const rep = editorInfo.ace_getRep();
      const lineTextLength = rep.lines.atIndex(currentLineNum).text.length;
      const endOfLinePos = [currentLineNum, lineTextLength];

      editorInfo.ace_performSelectionChange(endOfLinePos, endOfLinePos, false);
  editorInfo.ace_performDocumentReplaceRange(endOfLinePos, endOfLinePos, '\n');

      editorInfo.ace_updateBrowserSelectionFromRep();
  editorInfo.ace_focus();

      const editor = editorInfo.editor;
      if (editor) editor.ep_data_tables_last_clicked = null;

      return true;
    }
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during navigation:`, e);
    return false;
  }
}

/**
 * Find the line number for a specific table row
 * @param {string} tblId - Table ID
 * @param {number} targetRow - Target row index
 * @param {object} editorInfo - Editor instance
 * @param {object} docManager - Document attribute manager
 * @returns {number} - Line number (-1 if not found)
 */
function findLineForTableRow(tblId, targetRow, editorInfo, docManager) {
  const funcName = 'findLineForTableRow';

  try {
  const rep = editorInfo.ace_getRep();
    if (!rep || !rep.lines) {
      return -1;
    }

  const totalLines = rep.lines.length();
    for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
      try {
        let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

        if (!lineAttrString) {
          const lineEntry = rep.lines.atIndex(lineIndex);
          if (lineEntry && lineEntry.lineNode) {
            const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
            if (tableInDOM) {
              const domTblId = tableInDOM.getAttribute('data-tblId');
              const domRow = tableInDOM.getAttribute('data-row');
              if (domTblId === tblId && domRow !== null && parseInt(domRow, 10) === targetRow) {
                return lineIndex;
              }
            }
          }
        }

        if (lineAttrString) {
          const lineMetadata = JSON.parse(lineAttrString);
          if (lineMetadata.tblId === tblId && lineMetadata.row === targetRow) {
            return lineIndex;
          }
        }
      } catch (e) {
        continue;
      }
    }

    return -1;
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error searching for line:`, e);
  return -1;
  }
}

/**
 * Navigate to a specific cell and position caret at the end of its text
 * @param {number} targetLineNum - Target line number
 * @param {number} targetCellIndex - Target cell index (0-based)
 * @param {object} editorInfo - Editor instance
 * @param {object} docManager - Document attribute manager
 * @returns {boolean} - Success of navigation
 */
function navigateToCell(targetLineNum, targetCellIndex, editorInfo, docManager) {
  const funcName = 'navigateToCell';
  let targetPos;

  try {
    const rep = editorInfo.ace_getRep();
    if (!rep || !rep.lines) {
      return false;
    }

    const lineEntry = rep.lines.atIndex(targetLineNum);
    if (!lineEntry) {
      return false;
    }

    const lineText = lineEntry.text || '';
    const cells = lineText.split(DELIMITER);

    if (targetCellIndex >= cells.length) {
      return false;
    }

    let targetCol = 0;
    for (let i = 0; i < targetCellIndex; i++) {
      targetCol += (cells[i]?.length ?? 0) + DELIMITER.length;
    }
    const targetCellContent = cells[targetCellIndex] || '';
    targetCol += targetCellContent.length;

    const clampedTargetCol = Math.min(targetCol, lineText.length);
    targetPos = [targetLineNum, clampedTargetCol];

    try {
      const editor = editorInfo.ep_data_tables_editor;
    const tableMetadata = getTableLineMetadata(targetLineNum, editorInfo, docManager);

    if (editor && tableMetadata) {
      editor.ep_data_tables_last_clicked = {
        lineNum: targetLineNum,
        tblId: tableMetadata.tblId,
        cellIndex: targetCellIndex,
        relativePos: targetCellContent.length,
      };
      } else {
      }
    } catch (e) {
    }

    try {
    editorInfo.ace_performSelectionChange(targetPos, targetPos, false);

    editorInfo.ace_updateBrowserSelectionFromRep();

    editorInfo.ace_focus();

    } catch(e) {
      console.error(`[ep_data_tables] ${funcName}: Error during direct navigation update:`, e);
      return false;
    }

  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during cell navigation:`, e);
    return false;
  }

  return true;
}

// collectContentPre: Process tbljson-* classes and reapply table metadata during DOM collection
exports.collectContentPre = (hook, ctx) => {
  const state = ctx.state;
  const cc = ctx.cc;
  const classes = ctx.cls ? ctx.cls.split(' ') : [];
  for (const cls of classes) {
    if (cls.startsWith('tbljson-')) {
      const encodedMetadata = cls.substring(8);
      try {
        const decodedMetadata = dec(encodedMetadata);
        if (decodedMetadata) {
          cc.doAttrib(state, `${ATTR_TABLE_JSON}::${decodedMetadata}`);
        }
      } catch (e) {
        console.error(`[ep_data_tables] collectContentPre: Error decoding tbljson class:`, e);
      }
      break; // Only process first tbljson class found
    }
  }
};

exports.aceAttribsToClasses = (hook, ctx) => {
  const funcName = 'aceAttribsToClasses';
  if (ctx.key === ATTR_TABLE_JSON) {
    const rawJsonValue = ctx.value;

    let parsedMetadataForLog = '[JSON Parse Error]';
    try {
        parsedMetadataForLog = JSON.parse(rawJsonValue);
    } catch(e) {
    }

    const className = `tbljson-${enc(rawJsonValue)}`;
    return [className];
  }
  if (ctx.key === ATTR_CELL) {
    //// log(`${funcName}: Processing ATTR_CELL: ${ctx.value}`); // Optional: Uncomment if needed
    return [`tblCell-${ctx.value}`];
  }
  //// log(`${funcName}: Processing other key: ${ctx.key}`); // Optional: Uncomment if needed
  return [];
};



function escapeHtml(text = '') {
  const strText = String(text);
  var map = {
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  };
  return strText.replace(/[&<>"'']/g, function(m) { return map[m]; });
}
function buildTableFromDelimitedHTML(metadata, innerHTMLSegments) {
  const funcName = 'buildTableFromDelimitedHTML';

  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined') {
    console.error(`[ep_data_tables] ${funcName}: Invalid or missing metadata. Aborting.`);
    return '<table class="dataTable dataTable-error" writingsuggestions="false" autocorrect="off" autocapitalize="off" spellcheck="false"><tbody><tr><td>Error: Missing table metadata</td></tr></tbody></table>';
  }

  const numCols = innerHTMLSegments.length;
  const columnWidths = metadata.columnWidths || Array(numCols).fill(100 / numCols);

  while (columnWidths.length < numCols) {
    columnWidths.push(100 / numCols);
  }
  if (columnWidths.length > numCols) {
    columnWidths.splice(numCols);
  }

  const tdStyle = `padding: 5px 7px; word-wrap:break-word; vertical-align: top; border: 1px solid #000; position: relative;`;

  let encodedTbljsonClass = '';
  try {
    encodedTbljsonClass = `tbljson-${enc(JSON.stringify(metadata))}`;
  } catch (_) { encodedTbljsonClass = ''; }

  const cellsHtml = innerHTMLSegments.map((segment, index) => {
    const textOnly = (segment || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
    let modifiedSegment = segment || '';
    const containsImage = /\bimage-placeholder\b/.test(modifiedSegment);
    const isEmpty = (!segment || textOnly === '') && !containsImage;
    if (isEmpty) {
      const cellClass = encodedTbljsonClass ? `${encodedTbljsonClass} tblCell-${index}` : `tblCell-${index}`;
      modifiedSegment = `<span class="${cellClass}">&nbsp;</span>`;
    }
    if (index > 0) {
      const delimSpan = `<span class="ep-data_tables-delim" contenteditable="false">${HIDDEN_DELIM}</span>`;
      modifiedSegment = modifiedSegment.replace(/^(<span[^>]*>)/i, `$1${delimSpan}`);
      if (!/^<span[^>]*>/i.test(modifiedSegment)) modifiedSegment = `${delimSpan}${modifiedSegment}`;
    }

    const caretAnchorSpan = '<span class="ep-data_tables-caret-anchor" contenteditable="false"></span>';
    const anchorInjected = modifiedSegment.replace(/<\/span>\s*$/i, `${caretAnchorSpan}</span>`);
    modifiedSegment = (anchorInjected !== modifiedSegment)
      ? anchorInjected
      : (isEmpty
          ? `<span class="${encodedTbljsonClass ? `${encodedTbljsonClass} ` : ''}tblCell-${index}">${modifiedSegment}${caretAnchorSpan}</span>`
          : `${modifiedSegment}${caretAnchorSpan}`);

    try {
      const requiredCellClass = `tblCell-${index}`;
      const leadingDelimMatch = modifiedSegment.match(/^\s*<span[^>]*\bep-data_tables-delim\b[^>]*>[\s\S]*?<\/span>\s*/i);
      const head = leadingDelimMatch ? leadingDelimMatch[0] : '';
      const tail = leadingDelimMatch ? modifiedSegment.slice(head.length) : modifiedSegment;

      const openSpanMatch = tail.match(/^\s*<span([^>]*)>/i);
      if (!openSpanMatch) {
        const baseClasses = `${encodedTbljsonClass ? `${encodedTbljsonClass} ` : ''}${requiredCellClass}`;
        modifiedSegment = `${head}<span class="${baseClasses}">${tail}</span>`;
      } else {
        const fullOpen = openSpanMatch[0];
        const attrs = openSpanMatch[1] || '';
        const classMatch = /\bclass\s*=\s*"([^"]*)"/i.exec(attrs);
        let classList = classMatch ? classMatch[1].split(/\s+/).filter(Boolean) : [];
        classList = classList.filter(c => !/^tblCell-\d+$/.test(c));
        classList.push(requiredCellClass);
        const unique = Array.from(new Set(classList));
        const newClassAttr = ` class="${unique.join(' ')}"`;
        const attrsWithoutClass = classMatch ? attrs.replace(/\s*class\s*=\s*"[^"]*"/i, '') : attrs;
        const rebuiltOpen = `<span${newClassAttr}${attrsWithoutClass}>`;
        const afterOpen = tail.slice(fullOpen.length);
        const cleanedTail = afterOpen.replace(/(<span[^>]*class=")([^"]*)(")/ig, (m, p1, classes, p3) => {
          const filtered = classes.split(/\s+/).filter(c => c && !/^tblCell-\d+$/.test(c)).join(' ');
          return p1 + filtered + p3;
        });
        modifiedSegment = head + rebuiltOpen + cleanedTail;
      }
    } catch (_) { /* ignore normalization errors */ }

    const widthPercent = columnWidths[index] || (100 / numCols);
    const cellStyle = `${tdStyle} width: ${widthPercent}%;`;

    const isLastColumn = index === innerHTMLSegments.length - 1;
    const resizeHandle = !isLastColumn ? 
      `<div class="ep-data_tables-resize-handle" data-column="${index}" style="position: absolute; top: 0; right: -2px; width: 4px; height: 100%; cursor: col-resize; background: transparent; z-index: 10;"></div>` : '';

    const tdContent = `<td style="${cellStyle}" data-column="${index}" draggable="false" autocorrect="off" autocapitalize="off" spellcheck="false">${modifiedSegment}${resizeHandle}</td>`;
    return tdContent;
  }).join('');

  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';

  const tableHtml = `<table class="dataTable${firstRowClass}" writingsuggestions="false" autocorrect="off" autocapitalize="off" spellcheck="false" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse: collapse; table-layout: fixed;" draggable="false"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
  return tableHtml;
}

// acePostWriteDomLineHTML: Render table from delimiter-separated text
exports.acePostWriteDomLineHTML = function (hook_name, args, cb) {
  const funcName = 'acePostWriteDomLineHTML';
  const node = args?.node;
  const nodeId = node?.id;
  
  let lineNum = -1;
  try {
    if (node && nodeId) {
      lineNum = _getLineNumberOfElement(node);
      console.debug('[ep_data_tables:acePostWriteDomLineHTML] resolve-dom-index', { nodeId, lineNum });
    }
  } catch (e) {
    console.error('[ep_data_tables:acePostWriteDomLineHTML] resolve-exception', e);
  }
  const logPrefix = '[ep_data_tables:acePostWriteDomLineHTML]';

  if (!node || !nodeId) {
    console.error(`[ep_data_tables] ${funcName}: Received invalid node or node without ID.`);
    return cb();
  }

  let rowMetadata = null;
  let encodedJsonString = null;

  function findTbljsonClass(element) {
    if (element.classList) {
      for (const cls of element.classList) {
        if (cls.startsWith('tbljson-')) return cls.substring(8);
      }
    }
    if (element.children) {
      for (let i = 0; i < element.children.length; i++) {
        const found = findTbljsonClass(element.children[i]);
        if (found) return found;
      }
    }
    return null;
  }

  encodedJsonString = findTbljsonClass(node);

  if (!encodedJsonString) {
    return cb(); 
  }

  const existingTable = node.querySelector('table.dataTable[data-tblId]');
  if (existingTable) {
    return cb();
  }

  try {
    const decoded = dec(encodedJsonString);
      if (!decoded) throw new Error('Decoded string is null or empty.');
    rowMetadata = JSON.parse(decoded);

      if (!rowMetadata || typeof rowMetadata.tblId === 'undefined' || typeof rowMetadata.row === 'undefined' || typeof rowMetadata.cols !== 'number') {
          throw new Error('Invalid or incomplete metadata (missing tblId, row, or cols).');
      }

  } catch(e) { 
      console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Failed to decode/parse/validate tbljson.`, encodedJsonString, e);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Invalid table metadata attribute found.</div>';
    return cb();
  }

  const delimitedTextFromLine = node.innerHTML;

  const delimiterCount = (delimitedTextFromLine || '').split(DELIMITER).length - 1;

  let pos = -1;
  const delimiterPositions = [];
  while ((pos = delimitedTextFromLine.indexOf(DELIMITER, pos + 1)) !== -1) {
    delimiterPositions.push(pos);
  }

  const spanDelimRegex = /<span class="ep-data_tables-delim"[^>]*>[\s\S]*?<\/span>/ig;
  const sanitizedHTMLForSplit = (delimitedTextFromLine || '')
    .replace(spanDelimRegex, DELIMITER)
    .replace(/<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig, '')
    .replace(/\r?\n/g, ' ')
    .replace(/<br\s*\/?>/gi, ' ')
  //   .replace(/\u00A0/gu, ' ')
  //  .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ');
  const htmlSegments = sanitizedHTMLForSplit.split(DELIMITER);

  // Warn if segment contains multiple tblCell markers (possible corruption)
  for (let i = 0; i < htmlSegments.length; i++) {
    try {
      const uniqueCells = Array.from(new Set((htmlSegments[i] || '').match(/\btblCell-(\d+)\b/g) || []));
      if (uniqueCells.length > 1) console.warn('[ep_data_tables][diag] segment has multiple tblCell markers', { segIndex: i, uniqueCells });
    } catch (_) {}
  }

  let finalHtmlSegments = htmlSegments;
  let skipMismatchReturn = false;

  if (htmlSegments.length !== rowMetadata.cols) {
    // GRACEFUL DEGRADATION: Log the mismatch but render the table with actual segments
    // This prevents the cascade of: repair → re-render → mismatch → repair → crash
    // The table will display with its actual content (may differ from metadata column count)
    // Content and styling are preserved - user can edit normally
    // If user edits this row, handleDesktopCommitInput will normalize it at that time
    console.warn('[ep_data_tables][diag] Segment/column mismatch (rendering as-is)', { 
      nodeId, lineNum, 
      actualSegs: htmlSegments.length, 
      expectedCols: rowMetadata.cols, 
      tblId: rowMetadata.tblId, 
      row: rowMetadata.row 
    });
    
    // Always continue to render - don't return early, don't schedule repair
    skipMismatchReturn = true;
  }

  try {
    // De-dupe guard: if another ace-line already has a table for this tblId/row, skip rewrite here.
    try {
      const doc = node.ownerDocument;
      if (doc && rowMetadata && nodeId) {
        const selector = `div.ace-line:not(#${nodeId}) table.dataTable[data-tblId="${rowMetadata.tblId}"][data-row="${rowMetadata.row}"]`;
        const dupe = doc.querySelector(selector);
        if (dupe) {
          console.warn(`${logPrefix} NodeID#${nodeId}: Duplicate table detected elsewhere (tblId=${rowMetadata.tblId}, row=${rowMetadata.row}). Skipping rewrite to prevent duplication.`);
          return cb();
        } else {
          console.debug(`${logPrefix} NodeID#${nodeId}: no-duplicate-found`, {
            selector, tblId: rowMetadata.tblId, row: rowMetadata.row,
          });
        }
      }
    } catch (dupeErr) {
      console.error(`${logPrefix} NodeID#${nodeId}: duplicate-detection-error`, dupeErr);
    }
    const newTableHTML = buildTableFromDelimitedHTML(rowMetadata, finalHtmlSegments);

      const tbljsonElement = findTbljsonElement(node);

      if (tbljsonElement && tbljsonElement.parentElement && tbljsonElement.parentElement !== node) {
        const parentTag = tbljsonElement.parentElement.tagName.toLowerCase();
    const blockElements = ['center', 'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'right', 'left', 'ul', 'ol', 'li', 'code'];

        if (blockElements.includes(parentTag)) {
          tbljsonElement.parentElement.innerHTML = newTableHTML;
    } else {
      node.innerHTML = newTableHTML;
    }
      } else {
      node.innerHTML = newTableHTML;
      }

  } catch (renderError) {
      console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Error building/rendering table.`, renderError);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Failed to render table structure.</div>';
      return cb();
  }


  return cb();
};

function _getLineNumberOfElement(element) {
    let currentElement = element;
    let count = 0;
    while (currentElement = currentElement.previousElementSibling) {
        count++;
    }
    return count;
}
exports.aceKeyEvent = (h, ctx) => {
  const funcName = 'aceKeyEvent';
  const evt = ctx.evt;
  const rep = ctx.rep;
  const editorInfo = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

  const startLogTime = Date.now();
  const logPrefix = '[ep_data_tables:aceKeyEvent]';

  if (!rep || !rep.selStart || !editorInfo || !evt || !docManager) {
    return false;
  }

  const reportedLineNum = rep.selStart[0];
  const reportedCol = rep.selStart[1]; 

  let tableMetadata = null;
  let lineAttrString = null;
  try {
    lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);

    if (typeof docManager.getAttributesOnLine === 'function') {
      try {
        const allAttribs = docManager.getAttributesOnLine(reportedLineNum);
      } catch(e) {
      }
    }

    if (!lineAttrString) {
      try {
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(reportedLineNum);
        if (lineEntry && lineEntry.lineNode) {
          const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
          if (tableInDOM) {
            const domTblId = tableInDOM.getAttribute('data-tblId');
            const domRow = tableInDOM.getAttribute('data-row');
            const domCells = tableInDOM.querySelectorAll('td');
            if (domTblId && domRow !== null && domCells.length > 0) {
              const reconstructedMetadata = {
                tblId: domTblId,
                row: parseInt(domRow, 10),
                cols: domCells.length
              };
              lineAttrString = JSON.stringify(reconstructedMetadata);
            }
          }
        }
      } catch(e) {
      }
    }

    if (lineAttrString) {
        tableMetadata = JSON.parse(lineAttrString);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
             tableMetadata = null;
        }
    } else {
    }
  } catch(e) {
    console.error(`${logPrefix} Error checking/parsing line attribute for line ${reportedLineNum}.`, e);
    tableMetadata = null;
  }

  const editor = editorInfo.editor;
  const lastClick = editor?.ep_data_tables_last_clicked;

  let currentLineNum = -1;
  let targetCellIndex = -1;
  let relativeCaretPos = -1;
  let precedingCellsOffset = 0; 
  let cellStartCol = 0; 
  let lineText = '';
  let cellTexts = [];
  let metadataForTargetLine = null;
  let trustedLastClick = false;

  if (lastClick) {
      let storedLineAttrString = null;
      let storedLineMetadata = null;
      try {
          storedLineAttrString = docManager.getAttributeOnLine(lastClick.lineNum, ATTR_TABLE_JSON);

          if (storedLineAttrString) {
            storedLineMetadata = JSON.parse(storedLineAttrString);
          }

          if (storedLineMetadata && typeof storedLineMetadata.cols === 'number' && storedLineMetadata.tblId === lastClick.tblId) {
              trustedLastClick = true;
              currentLineNum = lastClick.lineNum; 
              targetCellIndex = lastClick.cellIndex;
              metadataForTargetLine = storedLineMetadata; 
              lineAttrString = storedLineAttrString;

              lineText = rep.lines.atIndex(currentLineNum)?.text || '';
              cellTexts = lineText.split(DELIMITER);

              if (cellTexts.length !== metadataForTargetLine.cols) {
              }

              cellStartCol = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                  cellStartCol += (cellTexts[i]?.length ?? 0) + DELIMITER.length;
              }
              precedingCellsOffset = cellStartCol;

              if (typeof lastClick.relativePos === 'number' && lastClick.relativePos >= 0) {
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(lastClick.relativePos, currentCellTextLength));
  } else {
                  relativeCaretPos = reportedCol - cellStartCol;
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(relativeCaretPos, currentCellTextLength)); 
              }
          } else {
              if (editor) editor.ep_data_tables_last_clicked = null;
          }
      } catch (e) {
           console.error(`${logPrefix} Error validating stored click info for line ${lastClick.lineNum}.`, e);
           if (editor) editor.ep_data_tables_last_clicked = null;
      }
  }

  if (!trustedLastClick) {
      try {
          lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
          if (lineAttrString) tableMetadata = JSON.parse(lineAttrString);
          if (!tableMetadata || typeof tableMetadata.cols !== 'number') tableMetadata = null;

          if (!lineAttrString) {
            try {
              const rep = editorInfo.ace_getRep();
              const lineEntry = rep.lines.atIndex(reportedLineNum);
              if (lineEntry && lineEntry.lineNode) {
                const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
                if (tableInDOM) {
                  const domTblId = tableInDOM.getAttribute('data-tblId');
                  const domRow = tableInDOM.getAttribute('data-row');
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domTblId && domRow !== null && domCells.length > 0) {
                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length
                    };
                    lineAttrString = JSON.stringify(reconstructedMetadata);
                    tableMetadata = reconstructedMetadata;
                  }
                }
              }
            } catch(e) {
            }
          }
      } catch(e) { tableMetadata = null; }

      if (!tableMetadata) {
           return false;
      }

      currentLineNum = reportedLineNum;
      metadataForTargetLine = tableMetadata;

      lineText = rep.lines.atIndex(currentLineNum)?.text || '';
      cellTexts = lineText.split(DELIMITER);

      if (cellTexts.length !== metadataForTargetLine.cols) {
      }

      let currentOffset = 0;
      let foundIndex = -1;
      for (let i = 0; i < cellTexts.length; i++) {
          const cellLength = cellTexts[i]?.length ?? 0;
          const cellEndCol = currentOffset + cellLength;
          if (reportedCol >= currentOffset && reportedCol <= cellEndCol) {
              foundIndex = i;
              relativeCaretPos = reportedCol - currentOffset;
              cellStartCol = currentOffset;
              precedingCellsOffset = cellStartCol;
              break; 
          }
          if (i < cellTexts.length - 1 && reportedCol === cellEndCol + DELIMITER.length) {
              foundIndex = i + 1;
              relativeCaretPos = 0; 
              cellStartCol = currentOffset + cellLength + DELIMITER.length;
              precedingCellsOffset = cellStartCol;
              break;
          }
          currentOffset += cellLength + DELIMITER.length;
      }

      if (foundIndex === -1) {
          if (reportedCol === lineText.length && cellTexts.length > 0) {
                foundIndex = cellTexts.length - 1;
                cellStartCol = 0; 
                for (let i = 0; i < foundIndex; i++) { cellStartCol += (cellTexts[i]?.length ?? 0) + DELIMITER.length; }
                precedingCellsOffset = cellStartCol;
                relativeCaretPos = cellTexts[foundIndex]?.length ?? 0; 
          } else {
            return false; 
          }
      }
      targetCellIndex = foundIndex;
  }

  if (currentLineNum < 0 || targetCellIndex < 0 || !metadataForTargetLine || targetCellIndex >= metadataForTargetLine.cols) {
    if (editor) editor.ep_data_tables_last_clicked = null;
    return false;
  }


  const selStartActual = rep.selStart;
  const selEndActual = rep.selEnd;
  const hasSelection = selStartActual[0] !== selEndActual[0] || selStartActual[1] !== selEndActual[1];

  if (hasSelection) {

    if (selStartActual[0] !== currentLineNum || selEndActual[0] !== currentLineNum) {
      evt.preventDefault();
      return true; 
    }

    let selectionStartColInLine = selStartActual[1];
    let selectionEndColInLine = selEndActual[1];

    const currentCellFullText = cellTexts[targetCellIndex] || '';
    const cellContentStartColInLine = cellStartCol;
    const cellContentEndColInLine = cellStartCol + currentCellFullText.length;

    /* If the user selected the whole cell plus delimiter characters,
     * clamp the selection to just the cell content.                        */
    const hasTrailingDelim =
      targetCellIndex < metadataForTargetLine.cols - 1 &&
      selectionEndColInLine === cellContentEndColInLine + DELIMITER.length;

    const hasLeadingDelim =
      targetCellIndex > 0 &&
      selectionStartColInLine === cellContentStartColInLine - DELIMITER.length;

    console.log(`[ep_data_tables:highlight-deletion] Selection analysis:`, {
      targetCellIndex,
      totalCols: metadataForTargetLine.cols,
      selectionStartCol: selectionStartColInLine,
      selectionEndCol: selectionEndColInLine,
      cellContentStartCol: cellContentStartColInLine,
      cellContentEndCol: cellContentEndColInLine,
      delimiterLength: DELIMITER.length,
      expectedTrailingDelimiterPos: cellContentEndColInLine + DELIMITER.length,
      expectedLeadingDelimiterPos: cellContentStartColInLine - DELIMITER.length,
      hasTrailingDelim,
      hasLeadingDelim,
      cellText: currentCellFullText
    });

    if (hasLeadingDelim) {
      console.log(`[ep_data_tables:highlight-deletion] CLAMPING selection start from ${selectionStartColInLine} to ${cellContentStartColInLine}`);
      selectionStartColInLine = cellContentStartColInLine;
    }

    if (hasTrailingDelim) {
      console.log(`[ep_data_tables:highlight-deletion] CLAMPING selection end from ${selectionEndColInLine} to ${cellContentEndColInLine}`);
      selectionEndColInLine = cellContentEndColInLine;
    }


    const isSelectionEntirelyWithinCell =
      selectionStartColInLine >= cellContentStartColInLine &&
      selectionEndColInLine <= cellContentEndColInLine;


    if (isSelectionEntirelyWithinCell) {

      if (evt.type !== 'keydown') return false;

      if (evt.ctrlKey || evt.metaKey || evt.altKey) return false;

    }

    const isCurrentKeyDelete = evt.key === 'Delete' || evt.keyCode === 46;
    const isCurrentKeyBackspace = evt.key === 'Backspace' || evt.keyCode === 8;
    const isCurrentKeyTyping = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;


    if (isSelectionEntirelyWithinCell && (isCurrentKeyDelete || isCurrentKeyBackspace || isCurrentKeyTyping)) {

      if (evt.type !== 'keydown') {
        return false; 
      }
      evt.preventDefault();

      const rangeStart = [currentLineNum, selectionStartColInLine];
      const rangeEnd = [currentLineNum, selectionEndColInLine];
      let replacementText = '';
      let newAbsoluteCaretCol = selectionStartColInLine;
      const repBeforeEdit = editorInfo.ace_getRep();

      if (isCurrentKeyTyping) {
        replacementText = evt.key;
        newAbsoluteCaretCol = selectionStartColInLine + replacementText.length;
      } else {
        const isWholeCell = selectionStartColInLine <= cellContentStartColInLine && selectionEndColInLine >= cellContentEndColInLine;
        if (isWholeCell) {
          replacementText = ' ';
          newAbsoluteCaretCol = selectionStartColInLine + 1;
        }
      }

      try {
        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, replacementText);

        if (replacementText.length > 0) {
          const attrStart = [currentLineNum, selectionStartColInLine];
          const attrEnd   = [currentLineNum, selectionStartColInLine + replacementText.length];
          console.log(`[ep_data_tables:highlight-deletion] Applying cell attribute to replacement text "${replacementText}" at range [${attrStart[0]},${attrStart[1]}] to [${attrEnd[0]},${attrEnd[1]}]`);
          editorInfo.ace_performDocumentApplyAttributesToRange(
            attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
          );
        }
        const repAfterReplace = editorInfo.ace_getRep();


        const applyHelper = editorInfo.ep_data_tables_applyMeta;
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) {
          const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;
          applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
        } else {
          console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
          const currentRepFallback = editorInfo.ace_getRep();
          if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
            applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
          } else {
            console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute even with fallback rep.`);
          }
        }

        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterSelectionChange = editorInfo.ace_getRep();

    editorInfo.ace_fastIncorp(1);
        const repAfterFastIncorp = editorInfo.ace_getRep();

        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterReassert = editorInfo.ace_getRep();

        const newRelativePos = newAbsoluteCaretCol - cellStartCol;
        if (editor) {
            editor.ep_data_tables_last_clicked = {
                lineNum: currentLineNum,
                tblId: metadataForTargetLine.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos
            };
        } else {
        }

      return true;
      } catch (error) {
        console.error('[ep_data_tables] Error processing highlight modification:', error);
        return true;
      }
    }
  }

  const isCutKey = (evt.ctrlKey || evt.metaKey) && (evt.key === 'x' || evt.key === 'X' || evt.keyCode === 88);
  if (isCutKey && hasSelection) {
    return false;
  } else if (isCutKey && !hasSelection) {
    return false;
  }

  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete' || evt.keyCode === 46;
  const isBackspaceKey = evt.key === 'Backspace' || evt.keyCode === 8;
  const isNavigationKey = [33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode);
  const isTabKey = evt.key === 'Tab';
  const isEnterKey = evt.key === 'Enter';

  const currentCellTextLengthEarly = cellTexts[targetCellIndex]?.length ?? 0;

  if (evt.type === 'keydown' && !evt.ctrlKey && !evt.metaKey && !evt.altKey) {
    if (evt.keyCode === 39 && relativeCaretPos >= currentCellTextLengthEarly && targetCellIndex < metadataForTargetLine.cols - 1) {
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, false, editorInfo, docManager);
      return true;
    }

    if (evt.keyCode === 37 && relativeCaretPos === 0 && targetCellIndex > 0) {
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, true, editorInfo, docManager);
      return true;
    }
  }


  if (isNavigationKey && !isTabKey) {
      if (editor) editor.ep_data_tables_last_clicked = null;
      return false;
  }

  if (isTabKey) { 
    evt.preventDefault();

     if (evt.type !== 'keydown') {
    return true;
  }

     const success = navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, evt.shiftKey, editorInfo, docManager);
     if (!success) {
     }
     return true;
  }

  if (isEnterKey) {
    evt.preventDefault();

      if (evt.type !== 'keydown') {
    return true;
  }

      const success = navigateToCellBelow(currentLineNum, targetCellIndex, metadataForTargetLine, editorInfo, docManager);
      if (!success) {
      }
      return true; 
  }

      const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex > 0) {
    evt.preventDefault();
          return true;
      }
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex === 0) {
        evt.preventDefault();
        return true;
      }
  if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex < metadataForTargetLine.cols - 1) {
          evt.preventDefault();
          return true;
      }
      if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex === metadataForTargetLine.cols - 1) {
        evt.preventDefault();
        return true;
      }

  const isInternalBackspace = isBackspaceKey && relativeCaretPos > 0;
  const isInternalDelete = isDeleteKey && relativeCaretPos < currentCellTextLength;

  if ((isInternalBackspace && relativeCaretPos === 1 && targetCellIndex > 0) ||
      (isInternalDelete && relativeCaretPos === 0 && targetCellIndex > 0)) {
    evt.preventDefault();
    return true;
  }

  if (isTypingKey || isInternalBackspace || isInternalDelete) {
    if (isTypingKey && relativeCaretPos === 0 && targetCellIndex > 0) {
      const safePosAbs = cellStartCol + 1;
      editorInfo.ace_performSelectionChange([currentLineNum, safePosAbs], [currentLineNum, safePosAbs], false);
      editorInfo.ace_updateBrowserSelectionFromRep();
      relativeCaretPos = 1;
    }
    const currentCol = cellStartCol + relativeCaretPos;

    if (evt.type !== 'keydown') {
        return false; 
    }

    evt.preventDefault();

    let newAbsoluteCaretCol = -1;
    let repBeforeEdit = null;

    try {
        repBeforeEdit = editorInfo.ace_getRep();

    if (isTypingKey) {
            const insertPos = [currentLineNum, currentCol];
            editorInfo.ace_performDocumentReplaceRange(insertPos, insertPos, evt.key);
            newAbsoluteCaretCol = currentCol + 1;

        } else if (isInternalBackspace) {
            const delRangeStart = [currentLineNum, currentCol - 1];
            const delRangeEnd = [currentLineNum, currentCol];
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol - 1;

        } else if (isInternalDelete) {
            const delRangeStart = [currentLineNum, currentCol];
            const delRangeEnd = [currentLineNum, currentCol + 1];
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol;
        }
        const repAfterReplace = editorInfo.ace_getRep();




        const applyHelper = editorInfo.ep_data_tables_applyMeta; 
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) { 
             const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;


             applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
                } else {
             console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
             const currentRepFallback = editorInfo.ace_getRep();
             if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
                 applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
            } else {
                  console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute even with fallback rep.`);
             }
        }

        if (newAbsoluteCaretCol >= 0) {
             const newCaretPos = [currentLineNum, newAbsoluteCaretCol];
             try {
                editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                const repAfterSelectionChange = editorInfo.ace_getRep();

                editorInfo.ace_fastIncorp(1); 
                const repAfterFastIncorp = editorInfo.ace_getRep();

                const targetCaretPosForReassert = [currentLineNum, newAbsoluteCaretCol];
                editorInfo.ace_performSelectionChange(targetCaretPosForReassert, targetCaretPosForReassert, false);
                const repAfterReassert = editorInfo.ace_getRep();

                const newRelativePos = newAbsoluteCaretCol - cellStartCol;
                editor.ep_data_tables_last_clicked = {
                    lineNum: currentLineNum, 
                    tblId: metadataForTargetLine.tblId,
                    cellIndex: targetCellIndex,
                    relativePos: newRelativePos
                };


            } catch (selError) {
                 console.error(`${logPrefix} -> ERROR setting selection immediately:`, selError);
             }
        } else {
            }

        } catch (error) {
            console.error('[ep_data_tables] Error processing key event update:', error);
    return true;
  }

    const endLogTime = Date.now();
    return true;

  }


  const endLogTimeFinal = Date.now();
  return false;
};
exports.aceInitialized = (h, ctx) => {
  const logPrefix = '[ep_data_tables:aceInitialized]';
  const ed = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;
  try {
    EP_DT_EDITOR_INFO = ed;
    console.debug('[ep_data_tables:aceInitialized] stored editorInfo for late-stage hooks');
  } catch (_) {
    console.error('[ep_data_tables:aceInitialized] failed to store editorInfo');
  }

  try {
    if (typeof window !== 'undefined') {
      window.__epDataTablesReady = true;
      const guard = document.getElementById('ep-data-tables-guard');
      if (guard && guard.parentNode) guard.parentNode.removeChild(guard);
    }
  } catch (_) {}

  ed.ep_data_tables_applyMeta = applyTableLineMetadataAttribute;

  ed.ep_data_tables_docManager = docManager;

  ed.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = '[ep_data_tables:aceInitialized:callWithAceForListeners]';

    if (!ace || !ace.editor) {
      console.error(`${callWithAceLogPrefix} ERROR: ace or ace.editor is not available. Cannot attach listeners.`);
      return;
    }
    const editor = ace.editor;

    ed.ep_data_tables_editor = editor;

    // Retry logic for iframe access to handle timing/race conditions
    const tryGetIframeBody = (attempt = 0) => {
      if (attempt > 0) {
        console.log(`${callWithAceLogPrefix} Retry attempt ${attempt}/5 to access iframe body`);
      }
      
      const $iframeOuter = $('iframe[name="ace_outer"]');
      if ($iframeOuter.length === 0) {
        if (attempt < 5) {
          setTimeout(() => tryGetIframeBody(attempt + 1), 100);
          return;
        }
        console.error(`${callWithAceLogPrefix} ERROR: Could not find outer iframe (ace_outer) after ${attempt} attempts.`);
        return;
      }

      const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
      if ($iframeInner.length === 0) {
        if (attempt < 5) {
          setTimeout(() => tryGetIframeBody(attempt + 1), 100);
          return;
        }
        console.error(`${callWithAceLogPrefix} ERROR: Could not find inner iframe (ace_inner) after ${attempt} attempts.`);
        return;
      }

      const innerDocBody = $iframeInner.contents().find('body');
      if (innerDocBody.length === 0) {
        if (attempt < 5) {
          setTimeout(() => tryGetIframeBody(attempt + 1), 100);
          return;
        }
        console.error(`${callWithAceLogPrefix} ERROR: Could not find body element in inner iframe after ${attempt} attempts.`);
        return;
      }

      const $inner = $(innerDocBody[0]);
      if (attempt > 0) {
        console.log(`${callWithAceLogPrefix} Successfully found iframe body on attempt ${attempt + 1}`);
      }
     
      // SUCCESS - Now attach all listeners and set attributes
      attachListeners($inner, $iframeOuter, $iframeInner, innerDocBody);
    };

    const attachListeners = ($inner, $iframeOuter, $iframeInner, innerDocBody) => {
      try {
      const platformLabel = () => (isAndroidUA() ? 'android' : (isIOSUA() ? 'ios' : 'desktop'));
      const sanitizeCellContent = (cellText = '', preserveZeroWidth = false) => {
        let sanitized = normalizeSoftWhitespace(cellText || '')
          .replace(new RegExp(DELIMITER, 'g'), ' ');
        // Images use ZWS (U+200B) as placeholder - if we strip it, images are lost
        // Auto-detect: if the content contains ZWS, preserve it (likely image content)
        const hasZWS = /[\u200B\u200C\u200D\uFEFF]/.test(sanitized);
        if (!preserveZeroWidth && !hasZWS) {
          sanitized = sanitized.replace(/[\u200B\u200C\u200D\uFEFF]/g, '');
        }
        if (!sanitized) sanitized = ' ';
        return sanitized;
      };
      const logCompositionEvent = (tag, evt, extras = {}) => {
        try {
          const nativeEvt = evt && (evt.originalEvent || evt);
          const rep = ed.ace_getRep && ed.ace_getRep();
          const sel = rep && rep.selStart ? { line: rep.selStart[0], col: rep.selStart[1] } : null;
          let tblMeta = null;
          if (sel && typeof sel.line === 'number' && docManager && typeof docManager.getAttributeOnLine === 'function') {
            try {
              const lineAttr = docManager.getAttributeOnLine(sel.line, ATTR_TABLE_JSON);
              if (lineAttr) {
                tblMeta = JSON.parse(lineAttr);
              }
            } catch (metaErr) {
              tblMeta = { error: metaErr?.message };
            }
          }
          console.debug(`[ep_data_tables:compositionTrace] ${tag}`, {
            inputType: nativeEvt && nativeEvt.inputType,
            data: typeof nativeEvt?.data === 'string' ? nativeEvt.data : null,
            isComposing: !!(nativeEvt && nativeEvt.isComposing),
            eventType: nativeEvt && nativeEvt.type,
            selection: sel,
            tableMeta: tblMeta,
            platform: platformLabel(),
            ...extras,
          });
        } catch (logErr) {
          console.error('[ep_data_tables:compositionTrace] log error', logErr);
        }
      };

      let suppressNextInputCommit = false;
      // One-shot guard to skip the first desktop beforeinput commit after a compositionend-driven commit.
      let suppressNextBeforeinputCommitOnce = false;
      let desktopComposition = { active: false, start: null, end: null, lineNum: null, cellIndex: -1, snapshot: null, snapshotMeta: null };

      const getTableMetadataForLine = (lineNum) => {
        let lineAttrString = docManager && docManager.getAttributeOnLine
          ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON)
          : null;
        let tableMetadata = null;
        if (lineAttrString) {
          try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
        }
        if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
        return (tableMetadata && typeof tableMetadata.cols === 'number') ? tableMetadata : null;
      };

      const collectSanitizedCells = (lineEntry, tableMetadata, contextTag = 'collect') => {
        const sanitizedCells = new Array(tableMetadata.cols).fill(' ');
        try {
          logCompositionEvent(`${contextTag}-collect-start`, null, {
            lineEntryText: lineEntry?.text,
            metadataCols: tableMetadata.cols,
            lineNodePresent: !!lineEntry?.lineNode,
          });
        } catch (_) {}
        if (!lineEntry) return sanitizedCells;
        const lineText = lineEntry.text || '';
        const textCells = lineText.split(DELIMITER);
        if (textCells.length === tableMetadata.cols) {
          for (let i = 0; i < textCells.length; i++) {
            sanitizedCells[i] = sanitizeCellContent(textCells[i] || '');
          }
          logCompositionEvent(`${contextTag}-collect-result`, null, {
            source: 'text',
            sanitizedCells,
          });
          return sanitizedCells;
        }
        try {
          const tableNode = lineEntry.lineNode
            ? lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${tableMetadata.tblId}"]`)
            : null;
          if (tableNode) {
            const rowElement = tableNode.querySelector('tbody > tr');
            if (rowElement) {
              const domCells = Array.from(rowElement.children || []);
              if (domCells.length === tableMetadata.cols) {
                for (let i = 0; i < domCells.length; i++) {
                  sanitizedCells[i] = sanitizeCellContent(domCells[i].innerText || '');
                }
                logCompositionEvent(`${contextTag}-collect-result`, null, {
                  source: 'dom',
                  sanitizedCells,
                });
                return sanitizedCells;
              }
            }
          }
        } catch (_) { /* ignore DOM errors */ }
        for (let i = 0; i < Math.min(textCells.length, tableMetadata.cols); i++) {
          sanitizedCells[i] = sanitizeCellContent(textCells[i] || '');
        }
        logCompositionEvent(`${contextTag}-collect-result-fallback`, null, {
          source: 'fallback',
          sanitizedCells,
        });
        return sanitizedCells;
      };

      // Compute target cell index and base offset using RAW (unsanitized) line text.
      const computeTargetCellIndexFromRaw = (lineEntry, selectionCol) => {
        if (!lineEntry || typeof selectionCol !== 'number') return { index: -1, baseOffset: 0, cellLen: 0 };
        const text = lineEntry.text || '';
        const rawCells = text.split(DELIMITER);
        let offset = 0;
        for (let i = 0; i < rawCells.length; i++) {
          const len = rawCells[i]?.length ?? 0;
          const end = offset + len;
          if (selectionCol >= offset && selectionCol <= end) {
            return { index: i, baseOffset: offset, cellLen: len };
          }
          offset += len + DELIMITER.length;
        }
        return { index: -1, baseOffset: 0, cellLen: 0 };
      };

      const getDomCellTargetFromSelection = () => {
        try {
          const innerDoc = ($inner && $inner.length) ? $inner[0].ownerDocument : null;
          const sel = innerDoc && typeof innerDoc.getSelection === 'function' ? innerDoc.getSelection() : null;
          const anchor = sel && sel.anchorNode;
          if (!anchor) return null;
          const asElement = (anchor.nodeType === 1) ? anchor : (anchor.parentElement || null);
          if (!asElement || typeof asElement.closest !== 'function') return null;
          const cellSpan = asElement.closest('span[class*="tblCell-"]');
          let idx = -1;
          if (cellSpan && cellSpan.className) {
            const m = cellSpan.className.match(/\btblCell-(\d+)\b/);
            if (m) idx = parseInt(m[1], 10);
          }
          const table = asElement.closest('table.dataTable[data-tblId], table.dataTable[data-tblid]');
          const tblId = table ? (table.getAttribute('data-tblId') || table.getAttribute('data-tblid')) : null;
          const lineDiv = asElement.closest('div.ace-line');
          let lineNum = null;
          try {
            const rep = ed.ace_getRep && ed.ace_getRep();
            if (lineDiv && rep && rep.lines && typeof rep.lines.indexOfKey === 'function') {
              lineNum = rep.lines.indexOfKey(lineDiv.id);
            }
          } catch (_) {}
          return { idx, tblId, lineNum };
        } catch (_) { return null; }
      };

      // Find a line number by matching tblId across the document, using attribs or DOM as fallback.
      const findLineNumByTblId = (tblId) => {
        try {
          if (!tblId) return -1;
          const rep = ed.ace_getRep && ed.ace_getRep();
          if (!rep || !rep.lines) return -1;
          const total = rep.lines.length();
          for (let ln = 0; ln < total; ln++) {
            try {
              let s = docManager && docManager.getAttributeOnLine ? docManager.getAttributeOnLine(ln, ATTR_TABLE_JSON) : null;
              if (s) {
                try {
                  const meta = JSON.parse(s);
                  if (meta && meta.tblId === tblId) return ln;
                } catch (_) {}
              } else {
                const entry = rep.lines.atIndex(ln);
                const lineNode = entry && entry.lineNode;
                if (lineNode && typeof lineNode.querySelector === 'function') {
                  const table = lineNode.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
                  const domTblId = table && (table.getAttribute('data-tblId') || table.getAttribute('data-tblid'));
                  if (domTblId === tblId) return ln;
                }
              }
            } catch (_) { /* continue */ }
          }
        } catch (_) {}
        return -1;
      };

      const computeTargetCellIndexFromSelection = (selectionCol, sanitizedCells) => {
        if (typeof selectionCol !== 'number') return -1;
        let offset = 0;
        for (let i = 0; i < sanitizedCells.length; i++) {
          const cellLen = sanitizedCells[i]?.length ?? 0;
          const cellEnd = offset + cellLen;
          if (selectionCol >= offset && selectionCol <= cellEnd) return i;
          offset = cellEnd + DELIMITER.length;
        }
        return -1;
      };

      // Helper to extract ALL character-level styling from a line's DOM table
      // Returns array of { cellIdx, relStart, len, text, attrs }
      const extractStylingFromLineDOM = (lineNode) => {
        const extractedStyling = [];
        try {
          if (!lineNode) return extractedStyling;
          const tableEl = lineNode.querySelector('table.dataTable');
          if (!tableEl) return extractedStyling;
          
          const tds = tableEl.querySelectorAll('td');
          let totalSpansProcessed = 0;
          let totalSpansWithAttrs = 0;
          tds.forEach((td, cellIdx) => {
            const spans = td.querySelectorAll('span:not(.ep-data_tables-delim):not(.ep-data_tables-caret-anchor):not(.image-inner):not(.image-resize-handle)');
            let relPos = 0; // Position relative to cell start
            spans.forEach((span) => {
              totalSpansProcessed++;
              const parentSpan = span.parentElement?.closest?.('span[class*="image:"]');
              if (parentSpan && parentSpan !== span) {
                return;
              }
              
              // Detect if this is an image-related span (comprehensive check)
              // Includes: inline-image, image-placeholder, image:URL, image-height:, image-width:, imageCssAspectRatio:, image-id-
              const isImageSpan = span.classList && (
                span.classList.contains('inline-image') ||
                span.classList.contains('image-placeholder') ||
                span.classList.contains('character') ||
                Array.from(span.classList).some(c => 
                  c.startsWith('image:') || 
                  c.startsWith('image-height:') || 
                  c.startsWith('image-width:') ||
                  c.startsWith('imageCssAspectRatio:') ||
                  c.startsWith('image-id-') ||
                  c.startsWith('image-float-')
                )
              );
              
              // For images, get direct text only
              let text;
              if (isImageSpan) {
                text = Array.from(span.childNodes)
                  .filter(n => n.nodeType === 3) // TEXT_NODE
                  .map(n => n.nodeValue || '')
                  .join('')
                  .replace(/\u00A0/g, ' ');
              } else {
                text = (span.textContent || '').replace(/\u00A0/g, ' '); // Normalize nbsp
              }
              
              const textLen = text.length;
              if (textLen === 0) return;
              
              // Extract ALL classes except table-related ones
              const stylingAttrs = [];
              if (span.classList) {
                for (const cls of span.classList) {
                  if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) continue;
                  if (cls === 'ace-line' || cls.startsWith('ep-data_tables-')) continue;
                  
                  // Parse class name to attribute key-value pair
                  // Etherpad uses: attributeName:value (colon) or attributeName-value (hyphen)
                  if (cls.includes(':')) {
                    const colonIdx = cls.indexOf(':');
                    stylingAttrs.push([cls.substring(0, colonIdx), cls.substring(colonIdx + 1)]);
                  } else if (cls.includes('-')) {
                    const knownBooleanClasses = [
                      'inline-image', 'image-placeholder', 'image-inner', 'image-resize-handle'
                    ];
                    if (knownBooleanClasses.includes(cls)) {
                      stylingAttrs.push([cls, 'true']);
                    } else {
                      const knownPrefixes = ['image-id-', 'image-float-', 'font-family-'];
                      let matched = false;
                      for (const prefix of knownPrefixes) {
                        if (cls.startsWith(prefix)) {
                          const attrName = prefix.slice(0, -1);
                          const attrValue = cls.substring(prefix.length);
                          stylingAttrs.push([attrName, attrValue]);
                          matched = true;
                          console.debug('[ep_data_tables:extractStyling] prefix-match', { cls, prefix, attrName, attrValue });
                          break;
                        }
                      }
                      if (!matched) {
                        const dashIdx = cls.indexOf('-');
                        const attrName = cls.substring(0, dashIdx);
                        const attrValue = cls.substring(dashIdx + 1);
                        stylingAttrs.push([attrName, attrValue]);
                        if (attrValue.includes('-') && cls.startsWith('image')) {
                          console.warn('[ep_data_tables:extractStyling] POSSIBLE MISPARSED IMAGE CLASS', { cls, attrName, attrValue });
                        }
                      }
                    }
                  } else {
                    stylingAttrs.push([cls, 'true']);
                  }
                }
              }
              
              if (stylingAttrs.length > 0) {
                totalSpansWithAttrs++;
                let domCellLen = td.textContent?.replace(/\u00A0/g, ' ').length || 0;
                if (cellIdx > 0 && domCellLen > 0) domCellLen -= 1;
                const isImage = isImageSpan || stylingAttrs.some(([k]) => k === 'image' || k === 'inline');
                extractedStyling.push({ cellIdx, relStart: relPos, len: textLen, text, attrs: stylingAttrs, domCellLen, isImage });
              }
              relPos += textLen;
            });
          });
          const summary = extractedStyling.map(s => ({ cell: s.cellIdx, text: s.text.slice(0, 12).replace(/[\u200B]/g, '⁰'), attrs: s.attrs.map(a => a[0]).join(','), isImage: s.isImage }));
          console.debug('[ep_data_tables:extractStyling] extracted', {
            count: extractedStyling.length,
            spansProcessed: totalSpansProcessed,
            spansWithAttrs: totalSpansWithAttrs,
            byCells: [...new Set(extractedStyling.map(s => s.cellIdx))].sort(),
            summary: summary.slice(0, 8), // First 8 for debugging
          });
        } catch (_) {}
        return extractedStyling;
      };
      
      // Helper to re-apply extracted styling to a line after replacement
      // Uses TEXT-BASED MATCHING to find where styled content is in the new cells
      const reapplyStylingToLine = (aceInstance, lineNum, cells, extractedStyling) => {
        if (!extractedStyling || extractedStyling.length === 0) return;
        let appliedCount = 0;
        const appliedDetails = [];
        
        const imagePositionsPerCell = {};
        if (cells && cells.length > 0) {
          const cellLens = cells.map((c, i) => `c${i}:${c?.length || 0}`).join(' ');
          const fullLen = cells.reduce((sum, c) => sum + (c?.length || 0), 0) + (cells.length - 1);
          console.debug('[ep_data_tables:reapplyStyling] cells', { lineNum, cellLens, fullLen, delimiters: cells.length - 1 });
        }
        try {
          for (const style of extractedStyling) {
            const originalCellIdx = style.cellIdx;
            const styledText = style.text;
            const originalRelStart = style.relStart;
            
            if (!styledText || styledText.length === 0) continue;
            
            const originalCellContent = cells[originalCellIdx] || '';
            let foundCellIdx = -1;
            let foundPos = -1;
            let matchType = 'none';
            
            // Text-based matching with occurrence heuristic for duplicates
            const findWithOccurrenceHeuristic = (content, text, wasNearEnd) => {
              const first = content.indexOf(text);
              if (first === -1) return { pos: -1, type: 'not-found' };
              
              const second = content.indexOf(text, first + 1);
              if (second === -1) return { pos: first, type: 'unique' };
              if (wasNearEnd) {
                let last = first;
                let next = second;
                while (next !== -1) {
                  last = next;
                  next = content.indexOf(text, last + 1);
                }
                return { pos: last, type: 'last' };
              } else {
                return { pos: first, type: 'first' };
              }
            };
            
            const domCellLen = style.domCellLen || originalCellContent.length;
            const distFromEnd = domCellLen - (originalRelStart + styledText.length);
            const wasNearEnd = distFromEnd <= 2;
            
            // Short text: use occurrence-aware matching
            if (styledText.length < 4 && originalCellContent.length > 0) {
              const result = findWithOccurrenceHeuristic(originalCellContent, styledText, wasNearEnd);
              if (result.pos !== -1) {
                foundCellIdx = originalCellIdx;
                foundPos = result.pos;
                matchType = result.type === 'unique' ? 'unique-short' 
                          : result.type === 'last' ? 'last-occurrence' 
                          : 'first-occurrence';
              }
              
              if (foundCellIdx === -1) {
                const lowerResult = findWithOccurrenceHeuristic(
                  originalCellContent.toLowerCase(), 
                  styledText.toLowerCase(), 
                  wasNearEnd
                );
                if (lowerResult.pos !== -1) {
                  foundCellIdx = originalCellIdx;
                  foundPos = lowerResult.pos;
                  matchType = 'case-insensitive-' + lowerResult.type;
                }
              }
            }
            
            // Long text: simple indexOf
            if (foundCellIdx === -1 && styledText.length >= 4 && originalCellContent) {
              foundPos = originalCellContent.indexOf(styledText);
              if (foundPos !== -1) { foundCellIdx = originalCellIdx; matchType = 'exact'; }
              if (foundCellIdx === -1) {
                const lowerCell = originalCellContent.toLowerCase();
                const lowerStyled = styledText.toLowerCase();
                const lowerPos = lowerCell.indexOf(lowerStyled);
                if (lowerPos !== -1) {
                  foundCellIdx = originalCellIdx;
                  foundPos = lowerPos;
                  matchType = 'case-insensitive';
                }
              }
            }
            
            // Search other cells if not found
            if (foundCellIdx === -1) {
              for (let c = 0; c < cells.length; c++) {
                if (c === originalCellIdx) continue;
                const cellContent = cells[c] || '';
                foundPos = cellContent.indexOf(styledText);
                if (foundPos !== -1) {
                  foundCellIdx = c;
                  matchType = 'other-cell';
                  break;
                }
              }
            }
            
            if (foundCellIdx === -1 || foundPos === -1) {
              // Image styles: reuse position for zero-width characters
              const isZeroWidthText = /^[\u200B\u200C\u200D\uFEFF]+$/.test(styledText);
              if (style.isImage || isZeroWidthText) {
                if (imagePositionsPerCell[originalCellIdx] !== undefined) {
                  foundCellIdx = originalCellIdx; foundPos = imagePositionsPerCell[originalCellIdx]; matchType = 'image-reuse';
                } else {
                  const zwsMatch = originalCellContent.match(/[\u200B\u200C\u200D\uFEFF]/);
                  if (zwsMatch) {
                    foundCellIdx = originalCellIdx; foundPos = originalCellContent.indexOf(zwsMatch[0]); matchType = 'image-zws-fallback';
                    imagePositionsPerCell[originalCellIdx] = foundPos;
                  }
                }
              }
              if (foundCellIdx === -1 || foundPos === -1) {
                console.debug('[ep_data_tables:reapplyStyling] skipped', {
                  originalCellIdx, text: styledText.slice(0, 10), reason: 'not-found',
                  isImage: style.isImage, isZeroWidth: isZeroWidthText,
                });
                continue;
              }
            }
            
            const foundCellContent = cells[foundCellIdx] || '';
            if (foundPos + styledText.length > foundCellContent.length) {
              console.debug('[ep_data_tables:reapplyStyling] skipped', {
                foundCellIdx, text: styledText.slice(0, 10), reason: 'out-of-bounds',
                foundPos, textLen: styledText.length, cellLen: foundCellContent.length,
              });
              continue;
            }
            
            let absStart = 0;
            for (let c = 0; c < foundCellIdx; c++) {
              absStart += (cells[c]?.length || 0) + DELIMITER.length;
            }
            absStart += foundPos;
            
            const absEnd = absStart + styledText.length;
            
            let totalLineLen = 0;
            for (let c = 0; c < cells.length; c++) {
              totalLineLen += (cells[c]?.length || 0);
              if (c < cells.length - 1) totalLineLen += DELIMITER.length;
            }
            
            if (absEnd > totalLineLen) {
              console.debug('[ep_data_tables:reapplyStyling] skipped', {
                foundCellIdx, text: styledText.slice(0, 10), reason: 'exceeds-line-bounds',
                absStart, absEnd, totalLineLen,
              });
              continue;
            }
            
            if (style.attrs.length > 0) {
              const isZeroWidthText = /^[\u200B\u200C\u200D\uFEFF]+$/.test(styledText);
              if ((style.isImage || isZeroWidthText) && imagePositionsPerCell[foundCellIdx] === undefined) imagePositionsPerCell[foundCellIdx] = foundPos;
              const priorCellsDebug = []; let priorSum = 0;
              for (let c = 0; c < foundCellIdx; c++) { priorCellsDebug.push(`c${c}:${cells[c]?.length || 0}`); priorSum += (cells[c]?.length || 0) + DELIMITER.length; }
              console.debug('[ep_data_tables:reapplyStyling] applying', { cell: foundCellIdx, foundPos, priorSum, absStart, absEnd, text: styledText.slice(0, 10), match: matchType, priorCells: priorCellsDebug.join('+'), attrs: style.attrs.map(a => `${a[0]}:${String(a[1]).slice(0,10)}`) });
              aceInstance.ace_performDocumentApplyAttributesToRange(
                [lineNum, absStart],
                [lineNum, absEnd],
                style.attrs
              );
              appliedCount++;
              appliedDetails.push({
                cell: foundCellIdx,
                pos: `${absStart}-${absEnd}`,
                match: matchType,
              });
            }
          }
          if (appliedCount > 0) {
            console.debug('[ep_data_tables:reapplyStyling] applied', {
              lineNum, 
              applied: appliedCount, 
              total: extractedStyling.length,
              details: appliedDetails,
            });
            
            // Force immediate incorporation after applying attributes
            try {
              if (typeof aceInstance.ace_fastIncorp === 'function') {
                aceInstance.ace_fastIncorp(100);
              }
            } catch (_) {}
          }
        } catch (err) {
          console.debug('[ep_data_tables:reapplyStyling] error (non-fatal)', err?.message);
        }
      };
      
      // Expose styling functions for collectContentLineText
      ed.ep_data_tables_extractStyling = extractStylingFromLineDOM;
      
      /**
       * Restore cached styling for a line after collection has finished.
       * Called via setTimeout from collectContentLineText when canonical is emitted.
       * @param {string} cacheKey - The cache key (lineId like magicdomidXXX, or fallback like tblId:row)
       * @param {string} [tblId] - Optional table ID for fallback line lookup
       * @param {number} [row] - Optional row number for fallback line lookup
       */
      ed.ep_data_tables_restoreCachedStyling = (cacheKey, tblId, row) => {
        const cached = getCachedStyling(cacheKey);
        if (!cached || !cached.styling || cached.styling.length === 0) {
          console.debug('[ep_data_tables:restoreCachedStyling] no cached styling', { cacheKey: cacheKey?.slice?.(0, 20) || cacheKey });
          return;
        }
        
        try {
          ed.ace_callWithAce((aceInstance) => {
            try {
              const rep = aceInstance.ace_getRep();
              if (!rep || !rep.lines) {
                console.debug('[ep_data_tables:restoreCachedStyling] no rep');
                return;
              }
              
              let lineIndex = -1;
              
              // Try to find line by lineId first (if cacheKey looks like a magicdomid)
              if (cacheKey && cacheKey.startsWith('magicdomid')) {
                lineIndex = rep.lines.indexOfKey(cacheKey);
              }
              
              // Fallback: find line by tblId + row (search through all lines)
              if ((lineIndex < 0 || lineIndex >= rep.lines.length()) && tblId != null && row != null) {
                const totalLines = rep.lines.length();
                for (let li = 0; li < totalLines; li++) {
                  try {
                    const lineAttr = docManager?.getAttributeOnLine?.(li, ATTR_TABLE_JSON);
                    if (lineAttr) {
                      const meta = JSON.parse(lineAttr);
                      if (meta.tblId === tblId && meta.row === row) {
                        lineIndex = li;
                        console.debug('[ep_data_tables:restoreCachedStyling] found via tblId:row', { tblId, row, lineIndex });
                        break;
                      }
                    }
                  } catch (_) {}
                }
              }
              
              if (lineIndex < 0 || lineIndex >= rep.lines.length()) {
                console.debug('[ep_data_tables:restoreCachedStyling] line not found', { 
                  cacheKey: cacheKey?.slice?.(0, 20) || cacheKey, 
                  tblId, 
                  row,
                  lineIndex 
                });
                return;
              }
              
              const lineEntry = rep.lines.atIndex(lineIndex);
              const lineText = lineEntry?.text || '';
              const cells = lineText.split(DELIMITER);
              
              console.debug('[ep_data_tables:restoreCachedStyling] restoring', {
                cacheKey: cacheKey?.slice?.(0, 20) || cacheKey,
                lineIndex,
                stylingCount: cached.styling.length,
                cellCount: cells.length,
              });
              
              reapplyStylingToLine(aceInstance, lineIndex, cells, cached.styling);
              
              clearCachedStyling(cacheKey);
              
              console.debug('[ep_data_tables:restoreCachedStyling] restored successfully');
            } catch (innerErr) {
              console.debug('[ep_data_tables:restoreCachedStyling] inner error', innerErr?.message);
            }
          }, 'ep_data_tables:restoreCachedStyling', true);
        } catch (outerErr) {
          console.debug('[ep_data_tables:restoreCachedStyling] outer error', outerErr?.message);
        }
      };

      const applyCanonicalCellsToLine = ({
        lineNum,
        sanitizedCells,
        tableMetadata,
        targetCellIndex,
        evt,
        logTag,
        suppressInputCommit = false,
        preCapturedStyling = null, // Optional: pre-captured styling from before collection
      }) => {
        const canonicalLine = sanitizedCells.join(DELIMITER);
        logCompositionEvent(`${logTag}-apply-start`, evt, {
          lineNum,
          targetCellIndex,
          canonicalLine,
          sanitizedCells,
        });
        ed.ace_callWithAce((aceInstance) => {
          try {
            const repBefore = aceInstance.ace_getRep();
            const linesLength = repBefore.lines.length();
            const lineEntryBefore = repBefore.lines.atIndex(lineNum);
            logCompositionEvent(`${logTag}-ace-before`, evt, {
              linesLength,
              lineNum,
              lineEntryBeforeText: lineEntryBefore?.text,
              lineEntryBeforeKey: lineEntryBefore?.key,
              selStart: repBefore.selStart,
              selEnd: repBefore.selEnd,
            });
            if (!lineEntryBefore) {
              logCompositionEvent(`${logTag}-ace-missing-line`, evt, {
                linesLength,
                lineNum,
              });
              return;
            }
            
            // Validate line is safe to modify (prevents keyToNodeMap errors from IME/Grammarly)
            if (!isLineSafeToModify(repBefore, lineNum, `[ep_data_tables:${logTag}]`)) {
              logCompositionEvent(`${logTag}-ace-unsafe-line`, evt, { lineNum });
              return;
            }

            let extractedStyling = preCapturedStyling;
            if (!extractedStyling || extractedStyling.length === 0) {
              extractedStyling = extractStylingFromLineDOM(lineEntryBefore.lineNode);
            }
            if (extractedStyling && extractedStyling.length > 0) {
              console.debug('[ep_data_tables:applyCanonicalCellsToLine] using styling', {
                lineNum, count: extractedStyling.length,
                source: preCapturedStyling ? 'pre-captured' : 'lineEntry',
                sample: extractedStyling.slice(0, 2).map(s => ({ cell: s.cellIdx, attrs: s.attrs })),
              });
            }

            const existingLength = lineEntryBefore.text.length;
            try {
              aceInstance.ace_performDocumentReplaceRange([lineNum, 0], [lineNum, existingLength], canonicalLine);
            } catch (replaceErr) {
              logCompositionEvent(`${logTag}-ace-replace-error`, evt, { error: replaceErr?.stack || replaceErr });
              throw replaceErr;
            }
            logCompositionEvent(`${logTag}-ace-after-replace`, evt, {
              canonicalLine,
              existingLength,
            });

            let offset = 0;
            for (let i = 0; i < sanitizedCells.length; i++) {
              const cellText = sanitizedCells[i] || '';
              if (cellText.length > 0) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [lineNum, offset],
                  [lineNum, offset + cellText.length],
                  [[ATTR_CELL, String(i)]],
                );
              }
              offset += cellText.length;
              if (i < sanitizedCells.length - 1) offset += DELIMITER.length;
            }
            
            // Re-apply extracted styling after cell attributes
            reapplyStylingToLine(aceInstance, lineNum, sanitizedCells, extractedStyling);

            const repAfter = aceInstance.ace_getRep();
            ed.ep_data_tables_applyMeta(lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, repAfter, ed, null, docManager);

            let caretOffset = 0;
            for (let i = 0; i < targetCellIndex; i++) {
              caretOffset += (sanitizedCells[i]?.length ?? 0) + DELIMITER.length;
            }
            const sanitizedEndCol = caretOffset + (sanitizedCells[targetCellIndex]?.length ?? 0);
            const newCaretPos = [lineNum, sanitizedEndCol];
            try {
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              logCompositionEvent(`${logTag}-ace-selection-change`, evt, { newCaretPos });
            } catch (selErr) {
              logCompositionEvent(`${logTag}-ace-selection-error`, evt, { error: selErr?.stack || selErr });
            }

            const editor = ed.ep_data_tables_editor;
            if (editor) {
              editor.ep_data_tables_last_clicked = {
                lineNum,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: Math.max(0, (sanitizedCells[targetCellIndex]?.length ?? 0)),
              };
            }

            logCompositionEvent(logTag, evt, {
              lineNum,
              targetCellIndex,
              canonicalLine,
            });
          } catch (aceErr) {
            const repState = (() => {
              try {
                const repDebug = aceInstance.ace_getRep();
                const repDebugLinesLength = repDebug?.lines?.length && repDebug.lines.length();
                const lineEntryDebug = (repDebug?.lines?.atIndex && lineNum < repDebugLinesLength)
                  ? repDebug.lines.atIndex(lineNum)
                  : null;
                return {
                  linesLength: repDebugLinesLength,
                  selStart: repDebug?.selStart,
                  selEnd: repDebug?.selEnd,
                  lineText: lineEntryDebug?.text,
                  lineKey: lineEntryDebug?.key,
                  docTextSample: repDebug && typeof repDebug.alines?.[lineNum] === 'string'
                    ? repDebug.alines[lineNum]
                    : null,
                };
              } catch (_) { return null; }
            })();
            logCompositionEvent(`${logTag}-ace-error`, evt, {
              error: aceErr?.stack || aceErr?.message || aceErr,
              errorName: aceErr?.name,
              lineNum,
              canonicalLine,
              sanitizedCells,
              repState,
            });
            logCompositionEvent(`${logTag}-ace-error`, evt, {
              additionalContext: 'preErrorState',
              lineNum,
              sanitizedCells,
              beforeDOM: lineEntryBefore?.lineNode?.outerHTML,
            });
            try {
              logCompositionEvent(`${logTag}-ace-error-operations`, evt, {
                docLine0: aceInstance?.document?._doc?._lines?.[lineNum]?.text || null,
                hasPerformDocumentReplaceRange: typeof aceInstance?.ace_performDocumentReplaceRange === 'function',
                hasPerformSelectionChange: typeof aceInstance?.ace_performSelectionChange === 'function',
              });
            } catch (_) {
              // ignore introspection errors
            }
          }
        }, `${logTag}-ace`);

        if (suppressInputCommit) suppressNextInputCommit = true;
      };

      // Minimal in-cell replacement that preserves attributes outside the edited span.
      const applyCellEditMinimal = ({
        lineNum,
        tableMetadata,
        cellIndex,
        relStart,
        relEnd,
        insertText,
        evt,
        logTag = 'cell-edit-minimal',
        aceInstance: providedAceInstance = null,
      }) => {
        const safeSan = (s) => sanitizeCellContent(s || '');
        const run = (aceInstance) => {
          try {
            // Validate metadata and indices
            if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
              logCompositionEvent(`${logTag}-invalid-meta`, evt, { lineNum, meta: tableMetadata });
              return;
            }
            if (typeof cellIndex !== 'number' || cellIndex < 0 || cellIndex >= tableMetadata.cols) {
              logCompositionEvent(`${logTag}-invalid-cell-index`, evt, { lineNum, cellIndex });
              return;
            }
            const repBefore = aceInstance.ace_getRep();
            const lineEntryBefore = repBefore.lines.atIndex(lineNum);
            if (!lineEntryBefore) {
              logCompositionEvent(`${logTag}-missing-line`, evt, { lineNum });
              return;
            }
            const currentText = lineEntryBefore.text || '';
            const lineLen = currentText.length;
            const cells = currentText.split(DELIMITER);
            // Compute base offset up to the target cell
            let baseOffset = 0;
            for (let i = 0; i < cellIndex; i++) {
              baseOffset += (cells[i]?.length ?? 0) + DELIMITER.length;
            }
            const cellText = cells[cellIndex] || '';
            const cellLen = cellText.length;
            const cellAbsStart = baseOffset;
            const cellAbsEnd = baseOffset + cellLen;
            const sAbs = baseOffset + Math.max(0, Math.min(relStart, cellLen));
            const eAbs = baseOffset + Math.max(0, Math.min(relEnd, cellLen));
            // Strictly clamp to the cell boundaries so we never touch delimiters or adjacent structures
            const sFinal = Math.max(cellAbsStart, Math.min(cellAbsEnd, Math.min(sAbs, eAbs)));
            const eFinal = Math.max(cellAbsStart, Math.min(cellAbsEnd, Math.max(sAbs, eAbs)));
            const toInsert = safeSan(insertText);

            if (eFinal < sFinal) {
              logCompositionEvent(`${logTag}-invalid-range`, evt, { sFinal, eFinal });
              return;
            }
            if (!toInsert && sFinal === eFinal) {
              // No-op edit
              return;
            }

            logCompositionEvent(`${logTag}-apply-start`, evt, {
              lineNum, cellIndex, relStart, relEnd, start: sFinal, end: eFinal, insertLen: toInsert.length,
            });

            try {
              aceInstance.ace_performDocumentReplaceRange([lineNum, sFinal], [lineNum, eFinal], toInsert);
            } catch (replaceErr) {
              logCompositionEvent(`${logTag}-replace-error`, evt, { error: replaceErr?.stack || replaceErr });
              return;
            }

            const repAfter = aceInstance.ace_getRep();
            const insertedEnd = sFinal + toInsert.length;
            if (toInsert.length > 0) {
              try {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [lineNum, sFinal],
                  [lineNum, insertedEnd],
                  [[ATTR_CELL, String(cellIndex)]],
                );
              } catch (attrErr) {
                logCompositionEvent(`${logTag}-attr-error`, evt, { error: attrErr?.stack || attrErr });
              }
            }

            try {
              ed.ep_data_tables_applyMeta(
                lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, repAfter, ed, null, docManager,
              );
            } catch (metaErr) {
              logCompositionEvent(`${logTag}-meta-error`, evt, { error: metaErr?.stack || metaErr });
            }

            // Verify delimiter integrity; if mismatch, rebuild canonical line very cautiously
            try {
              const lineEntryAfter = repAfter.lines.atIndex(lineNum);
              const rebuilt = (lineEntryAfter?.text || '').split(DELIMITER);
              if (rebuilt.length !== tableMetadata.cols) {
                const sanitizedCells = new Array(tableMetadata.cols).fill(' ');
                for (let i = 0; i < Math.min(rebuilt.length, tableMetadata.cols); i++) {
                  sanitizedCells[i] = safeSan(rebuilt[i] || ' ');
                }
                const canonicalLine = sanitizedCells.join(DELIMITER);
                const existingLength2 = (lineEntryAfter?.text || '').length;
                try {
                  aceInstance.ace_performDocumentReplaceRange([lineNum, 0], [lineNum, existingLength2], canonicalLine);
                } catch (repl2Err) {
                  logCompositionEvent(`${logTag}-fallback-replace-error`, evt, { error: repl2Err?.stack || repl2Err });
                  return;
                }
                // Re-apply per-cell attributes only for non-empty spans
                let off = 0;
                for (let i = 0; i < sanitizedCells.length; i++) {
                  const t = sanitizedCells[i] || '';
                  if (t.length > 0) {
                    try {
                      aceInstance.ace_performDocumentApplyAttributesToRange(
                        [lineNum, off],
                        [lineNum, off + t.length],
                        [[ATTR_CELL, String(i)]],
                      );
                    } catch (attr2Err) {
                      logCompositionEvent(`${logTag}-fallback-attr-error`, evt, { error: attr2Err?.stack || attr2Err });
                    }
                  }
                  off += t.length;
                  if (i < sanitizedCells.length - 1) off += DELIMITER.length;
                }
                try {
                  const repFix = aceInstance.ace_getRep();
                  ed.ep_data_tables_applyMeta(lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, repFix, ed, null, docManager);
                } catch (meta2Err) {
                  logCompositionEvent(`${logTag}-fallback-meta-error`, evt, { error: meta2Err?.stack || meta2Err });
                }
              }
            } catch (rebuildErr) {
              logCompositionEvent(`${logTag}-rebuild-error`, evt, { error: rebuildErr?.stack || rebuildErr });
            }

            try {
              aceInstance.ace_performSelectionChange([lineNum, insertedEnd], [lineNum, insertedEnd], false);
            } catch (selErr) {
              logCompositionEvent(`${logTag}-selection-error`, evt, { error: selErr?.stack || selErr });
            }
            logCompositionEvent(`${logTag}`, evt, {
              lineNum, cellIndex, start: sFinal, end: insertedEnd,
            });
            try { aceInstance.ace_fastIncorp(5); } catch (incErr) {
              logCompositionEvent(`${logTag}-fastincorp-error`, evt, { error: incErr?.stack || incErr });
            }
          } catch (err) {
            logCompositionEvent(`${logTag}-error`, evt, { error: err?.stack || err });
          }
        };
        if (providedAceInstance) {
          run(providedAceInstance);
        } else {
          ed.ace_callWithAce((ace) => run(ace), `${logTag}-ace`);
        }
      };

      const handleDesktopCommitInput = (evt) => {
        const nativeEvt = evt && (evt.originalEvent || evt);
        if (evt && !evt._epDataTablesHandled) evt._epDataTablesHandled = true;
        if (nativeEvt && !nativeEvt._epDataTablesHandled) nativeEvt._epDataTablesHandled = true;

        // Capture styling from DOM before rep access
        let preCapturedStyling = null;
        let preCapturedLineNum = -1;
        try {
          const innerDoc = ed.ep_data_tables_innerDoc || 
                          (typeof $inner !== 'undefined' && $inner[0]?.ownerDocument) || 
                          document;
          const sel = innerDoc.defaultView?.getSelection?.() || window.getSelection?.();
          if (sel && sel.anchorNode) {
            const aceLine = sel.anchorNode.nodeType === 1 
              ? sel.anchorNode.closest?.('div.ace-line')
              : sel.anchorNode.parentElement?.closest?.('div.ace-line');
            if (aceLine) {
              const allAceLines = innerDoc.querySelectorAll('div.ace-line');
              for (let i = 0; i < allAceLines.length; i++) {
                if (allAceLines[i] === aceLine) {
                  preCapturedLineNum = i;
                  preCapturedStyling = extractStylingFromLineDOM(aceLine);
                  break;
                }
              }
            }
          }
        } catch (preCapErr) {}

        const rep = ed.ace_getRep();
        if (!rep || !rep.selStart) return;
        let lineNum = rep.selStart[0];

        let tableMetadata = getTableMetadataForLine(lineNum);
        
        // Try pre-captured line if table metadata not found on current line
        if (!tableMetadata && preCapturedLineNum >= 0 && preCapturedLineNum !== lineNum) {
          tableMetadata = getTableMetadataForLine(preCapturedLineNum);
          if (tableMetadata) lineNum = preCapturedLineNum;
        }
        
        if (!tableMetadata) return;

        const lineEntry = rep.lines.atIndex(lineNum);
        let sanitizedCells = collectSanitizedCells(lineEntry, tableMetadata, 'input-commit');
        const currentLineText = lineEntry?.text || '';
        let canonicalLine = sanitizedCells.join(DELIMITER);
        
        const usablePreCapturedStyling = (preCapturedStyling && preCapturedStyling.length > 0)
          ? preCapturedStyling
          : null;

        const resolveTargetCellIndex = (cells) => {
          let idx = computeTargetCellIndexFromSelection(rep.selStart[1], cells);
          if (idx === -1) {
            const editor = ed.ep_data_tables_editor;
            const lastClick = editor?.ep_data_tables_last_clicked;
            if (lastClick && lastClick.tblId === tableMetadata.tblId) {
              idx = lastClick.cellIndex;
            }
          }
          if (idx === -1) idx = Math.min(tableMetadata.cols - 1, 0);
          return idx;
        };

        const forceCanonicalRewrite = (cells, reasonTag, extras = {}) => {
          const targetCellIndex = resolveTargetCellIndex(cells);
          logCompositionEvent(reasonTag, evt, {
            lineNum,
            targetCellIndex,
            cells,
            canonical: cells.join(DELIMITER),
            ...extras,
          });
          applyCanonicalCellsToLine({
            lineNum,
            sanitizedCells: cells,
            tableMetadata,
            targetCellIndex,
            evt,
            logTag: `${reasonTag}-apply`,
            suppressInputCommit: true,
            preCapturedStyling: usablePreCapturedStyling, // Pass pre-captured styling
          });
        };

        if (sanitizedCells.length !== tableMetadata.cols) {
          const domCells = collectSanitizedCells(lineEntry, tableMetadata, 'input-commit-dom');
          forceCanonicalRewrite(domCells, 'input-commit-canonicalize-cells', {
            reason: 'cell-count-mismatch-text',
            metadataCols: tableMetadata.cols,
            sanitizedCells,
            currentLineText,
          });
          return;
        }

        if (canonicalLine === currentLineText) {
          const domCells = collectSanitizedCells(lineEntry, tableMetadata, 'input-commit-dom');
          const domCanonical = domCells.join(DELIMITER);
          if (domCanonical !== canonicalLine) {
            forceCanonicalRewrite(domCells, 'input-commit-force-dom', {
              reason: 'dom-canonical-differs',
              canonicalLine,
              domCanonical,
            });
            return;
          }
          
          // Even when text matches, styling may have been lost during collection
          try {
            if (usablePreCapturedStyling && usablePreCapturedStyling.length > 0) {
              // Get fresh line text inside ace_callWithAce
              ed.ace_callWithAce((aceInstance) => {
                const freshRep = aceInstance.ace_getRep();
                
                // Use preCapturedLineNum for styling to avoid stale lineNum
                const stylingLineNum = (preCapturedLineNum >= 0 && preCapturedLineNum < freshRep.lines.length())
                  ? preCapturedLineNum
                  : lineNum;
                
                const freshLineEntry = freshRep.lines.atIndex(stylingLineNum);
                const freshLineText = freshLineEntry?.text || '';
                const actualCells = freshLineText.split(DELIMITER);
                
                reapplyStylingToLine(aceInstance, stylingLineNum, actualCells, usablePreCapturedStyling);
                
                // Clean up blank lines after this table row
                try {
                  const repAfter = aceInstance.ace_getRep();
                  if (repAfter && repAfter.lines) {
                    const totalLines = repAfter.lines.length();
                    const blankLinesToRemove = [];
                    
                    for (let li = stylingLineNum + 1; li < totalLines && li < stylingLineNum + 10; li++) {
                      const checkEntry = repAfter.lines.atIndex(li);
                      const checkText = checkEntry?.text || '';
                      
                      const isTableRow = docManager && typeof docManager.getAttributeOnLine === 'function'
                        && docManager.getAttributeOnLine(li, ATTR_TABLE_JSON);
                      if (isTableRow) break;
                      
                      if (!checkText.trim() || checkText === '\n') {
                        blankLinesToRemove.push(li);
                      } else {
                        break;
                      }
                    }
                    
                    if (blankLinesToRemove.length > 0) {
                      blankLinesToRemove.sort((a, b) => b - a).forEach((blankLineNum) => {
                        try {
                          const checkRep = aceInstance.ace_getRep();
                          if (checkRep && blankLineNum < checkRep.lines.length()) {
                            aceInstance.ace_performDocumentReplaceRange([blankLineNum, 0], [blankLineNum + 1, 0], '');
                          }
                        } catch (_) {}
                      });
                    }
                  }
                } catch (_) {}
              }, 'ep_data_tables:input-commit-restore-styling', true);
              logCompositionEvent('input-commit-restored-styling', evt, { 
                lineNum, 
                styleCount: usablePreCapturedStyling.length,
                source: 'pre-captured',
              });
            } else {
              logCompositionEvent('input-commit-skip-canonical', evt, { 
                lineNum, 
                reason: 'no-pre-captured-styling',
              });
            }
          } catch (styleErr) {
            logCompositionEvent('input-commit-skip-canonical', evt, { lineNum, styleError: styleErr?.message });
          }
          return;
        }

        let targetCellIndex = computeTargetCellIndexFromSelection(rep.selStart[1], sanitizedCells);
        if (targetCellIndex === -1) {
          const editor = ed.ep_data_tables_editor;
          const lastClick = editor?.ep_data_tables_last_clicked;
          if (lastClick && lastClick.tblId === tableMetadata.tblId) {
            targetCellIndex = lastClick.cellIndex;
          }
        }
        if (targetCellIndex === -1) targetCellIndex = Math.min(tableMetadata.cols - 1, 0);

        logCompositionEvent('input-commit-rebuild', evt, {
          lineNum,
          targetCellIndex,
          canonicalLine,
          sanitizedCells,
        });
        const beforeDOM = lineEntry?.lineNode?.outerHTML;
        logCompositionEvent('input-commit-before-dom', evt, { lineNum, beforeDOM });
        applyCanonicalCellsToLine({
          lineNum,
          sanitizedCells,
          tableMetadata,
          targetCellIndex,
          evt,
          logTag: 'input-commit-canonical',
          suppressInputCommit: true,
        });
      };

      $inner.on('compositionstart', (evt) => {
        logCompositionEvent('compositionstart', evt);
        // Set global composition flag to prevent aggressive suppression
        __epDT_compositionActive = true;
        if (isAndroidUA() || isIOSUA()) return;
        try {
          const rep0 = ed.ace_getRep && ed.ace_getRep();
          if (rep0 && rep0.selStart) {
            let lineNum = rep0.selStart[0];
            let cellIndex = -1;
            let tableMetadata = null;
            let cellSnapshot = null;
            let originalTableLine = null; // Track where the table actually is
            
            // Recover cursor if browser moved it to wrong line between compositions
            let usedCursorRecovery = false;
            if (desktopComposition && desktopComposition._expectedCursor) {
              const stored = desktopComposition._expectedCursor;
              const timeSinceStored = Date.now() - stored.timestamp;
              
              if (timeSinceStored < 1000) {
                const currentLine = lineNum;
                const expectedLine = stored.lineNum;
                const expectedCell = stored.cellIndex;
                
                const lineChanged = currentLine !== expectedLine;
                const seemsWrong = lineChanged && (currentLine === expectedLine + 1);
                
                if (seemsWrong) {
                  const storedMeta = getTableMetadataForLine(expectedLine);
                  if (storedMeta && storedMeta.tblId === stored.tblId) {
                    lineNum = expectedLine;
                    cellIndex = expectedCell;
                    tableMetadata = storedMeta;
                    originalTableLine = expectedLine;
                    usedCursorRecovery = true;
                    
                    try {
                      ed.ace_callWithAce((aceInstance) => {
                        aceInstance.ace_performSelectionChange([expectedLine, stored.col], [expectedLine, stored.col], false);
                      }, 'compositionstart-cursor-recovery', true);
                    } catch (_) {}
                  }
                }
              }
              
              if (timeSinceStored > 1000 || usedCursorRecovery) {
                delete desktopComposition._expectedCursor;
              }
            }
            
            try {
              // Find table from DOM selection (works even if rep line is wrong)
              const domTarget = usedCursorRecovery ? null : getDomCellTargetFromSelection();
              const domFoundTable = !!(domTarget && domTarget.tblId);
              
              if (domFoundTable && typeof domTarget.lineNum === 'number' && domTarget.lineNum >= 0) {
                originalTableLine = domTarget.lineNum;
                lineNum = domTarget.lineNum;
                if (typeof domTarget.idx === 'number' && domTarget.idx >= 0) cellIndex = domTarget.idx;
              }
              
              if (!tableMetadata) tableMetadata = getTableMetadataForLine(lineNum);
              
              // DOM fallback for metadata
              if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
                try {
                  const entry = rep0.lines.atIndex(lineNum);
                  const lineNode = entry?.lineNode;
                  if (lineNode) {
                    const tableEl = lineNode.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
                    if (tableEl) {
                      const domTblId = tableEl.getAttribute('data-tblId') || tableEl.getAttribute('data-tblid');
                      const domRow = parseInt(tableEl.getAttribute('data-row') || '0', 10);
                      const tr = tableEl.querySelector('tbody > tr');
                      const domCols = tr ? tr.children.length : 0;
                      if (domTblId && domCols > 0) {
                        tableMetadata = { tblId: domTblId, row: domRow, cols: domCols };
                        originalTableLine = lineNum;
                      }
                    }
                  }
                } catch (_) {}
              } else {
                originalTableLine = lineNum;
              }
              
              if (tableMetadata && typeof tableMetadata.cols === 'number') {
                const entry = rep0.lines.atIndex(originalTableLine ?? lineNum);
                if (cellIndex < 0) {
                  // Fall back to RAW mapping to avoid normalization drift.
                  const rawMap = computeTargetCellIndexFromRaw(entry, rep0.selStart[1]);
                  cellIndex = rawMap.index;
                }
                // Capture snapshot of cell contents for corruption recovery
                // Try DOM-based cell extraction for higher fidelity
                // Capture snapshot of cell contents for corruption recovery
                if (entry && tableMetadata.cols > 0) {
                  const lineNode = entry?.lineNode;
                  const tableEl = lineNode?.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
                  if (tableEl) {
                    const tr = tableEl.querySelector('tbody > tr');
                    if (tr && tr.children.length === tableMetadata.cols) {
                      cellSnapshot = new Array(tableMetadata.cols);
                      for (let i = 0; i < tableMetadata.cols; i++) {
                        cellSnapshot[i] = sanitizeCellContent(tr.children[i]?.innerText || ' ');
                      }
                    }
                  }
                  if (!cellSnapshot) {
                    const cells = (entry.text || '').split(DELIMITER);
                    cellSnapshot = new Array(tableMetadata.cols);
                    for (let i = 0; i < tableMetadata.cols; i++) {
                      cellSnapshot[i] = sanitizeCellContent(cells[i] || ' ');
                    }
                  }
                }
              }
            } catch (metaErr) {
              console.debug('[ep_data_tables:compositionstart] metadata capture error', metaErr);
            }
            desktopComposition = {
              active: true,
              start: rep0.selStart.slice(),
              end: rep0.selEnd ? rep0.selEnd.slice() : rep0.selStart.slice(),
              lineNum, cellIndex,
              tblId: tableMetadata?.tblId ?? null,
              snapshot: cellSnapshot,
              snapshotMeta: tableMetadata ? { ...tableMetadata } : null,
              originalTableLine, usedCursorRecovery,
            };
            
            if (tableMetadata?.tblId && originalTableLine !== null) {
              __epDT_compositionOriginalLine = { tblId: tableMetadata.tblId, lineNum: originalTableLine, timestamp: Date.now() };
            }
          }
        } catch (err) {
          console.debug('[ep_data_tables:compositionstart] error', err);
          desktopComposition = { active: false, start: null, end: null, lineNum: null, cellIndex: -1, snapshot: null, snapshotMeta: null, originalTableLine: null };
        }
      });

      $inner.on('compositionupdate', (evt) => {
        logCompositionEvent('compositionupdate', evt);
        if (isAndroidUA() || isIOSUA()) return;
        try {
          if (desktopComposition.active) {
            const repN = ed.ace_getRep && ed.ace_getRep();
            if (repN && repN.selEnd) desktopComposition.end = repN.selEnd.slice();
          }
        } catch (_) {}
      });

      $inner.on('input', (evt) => {
        const nativeEvt = evt && (evt.originalEvent || evt);
        if (!nativeEvt) return;
        const isInsertType = typeof nativeEvt.inputType === 'string' && nativeEvt.inputType.startsWith('insert');
        if (nativeEvt.isComposing || isInsertType) {
          logCompositionEvent('input', evt);
        }
        if (suppressNextInputCommit) {
          suppressNextInputCommit = false;
          return;
        }
        if (!isAndroidUA() && !isIOSUA() &&
            !nativeEvt.isComposing &&
            (nativeEvt.inputType === 'insertCompositionText' || nativeEvt.inputType === 'insertText')) {
          handleDesktopCommitInput(evt);
        }
      });

      $inner.on('textInput', (evt) => {
        logCompositionEvent('textInput', evt);
      });

      const mobileSuggestionBlocker = (evt) => {
        const t = evt && evt.inputType || '';
        const dataStr = (evt && typeof evt.data === 'string') ? evt.data : '';
        const isProblem = (
          t === 'insertReplacementText' ||
          t === 'insertFromComposition' ||
          (t === 'insertText' && !evt.isComposing && (!dataStr || dataStr.length > 1))
        );
        if (!isProblem) return;

        try {
          const repQuick = ed.ace_getRep && ed.ace_getRep();
          if (!repQuick || !repQuick.selStart) return;
          const lineNumQuick = repQuick.selStart[0];
          let metaStrQuick = docManager && docManager.getAttributeOnLine
            ? docManager.getAttributeOnLine(lineNumQuick, ATTR_TABLE_JSON)
            : null;
          let metaQuick = null;
          if (metaStrQuick) { try { metaQuick = JSON.parse(metaStrQuick); } catch (_) {} }
          if (!metaQuick) metaQuick = getTableLineMetadata(lineNumQuick, ed, docManager);
          if (!metaQuick || typeof metaQuick.cols !== 'number') return;
        } catch (_) { return; }

        evt._epDataTablesHandled = true;
        if (evt.originalEvent) evt.originalEvent._epDataTablesHandled = true;
        evt.preventDefault();
        if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

        const capturedInputType = t;
        const capturedData = dataStr;

        setTimeout(() => {
          try {
            ed.ace_callWithAce((aceInstance) => {
              aceInstance.ace_fastIncorp(10);
              const rep = aceInstance.ace_getRep();
              if (!rep || !rep.selStart || !rep.selEnd) return;

              const selStart = [...rep.selStart];
              const selEnd = [...rep.selEnd];
              const lineNum = selStart[0];

              let metaStr = docManager && docManager.getAttributeOnLine
                ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON)
                : null;
              let tableMetadata = null;
              if (metaStr) { try { tableMetadata = JSON.parse(metaStr); } catch (_) {} }
              if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
              if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined') {
                return;
              }

              const initialHasSelection = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);

              let replacement = typeof capturedData === 'string' ? capturedData : '';
              if (!replacement) {
                if (capturedInputType === 'insertText' && !initialHasSelection) {
                  replacement = ' ';
                } else if (capturedInputType === 'insertReplacementText' || capturedInputType === 'insertFromComposition' || initialHasSelection) {
                  replacement = ' ';
                }
              }

              replacement = normalizeSoftWhitespace(
                (replacement || '')
                  .replace(new RegExp(DELIMITER, 'g'), ' ')
                  .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
              );

              if (!replacement) replacement = ' ';

              const lineEntry = rep.lines.atIndex(lineNum);
              const lineText = lineEntry?.text || '';
              const cells = lineText.split(DELIMITER);
              let currentOffset = 0;
              let targetCellIndex = -1;
              let cellStartCol = 0;
              let cellEndCol = 0;
              for (let i = 0; i < cells.length; i++) {
                const cellLen = cells[i]?.length ?? 0;
                const cellEndThis = currentOffset + cellLen;
                if (selStart[1] >= currentOffset && selStart[1] <= cellEndThis) {
                  targetCellIndex = i;
                  cellStartCol = currentOffset;
                  cellEndCol = cellEndThis;
                  break;
                }
                currentOffset += cellLen + DELIMITER.length;
              }

              if (targetCellIndex === -1) {
                aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, replacement);
                const repAfterFallback = aceInstance.ace_getRep();
                ed.ep_data_tables_applyMeta(
                  lineNum,
                  tableMetadata.tblId,
                  tableMetadata.row,
                  tableMetadata.cols,
                  repAfterFallback,
                  ed,
                  null,
                  docManager
                );
                const fallbackLineEntry = repAfterFallback.lines.atIndex(lineNum);
                const fallbackMaxLen = fallbackLineEntry ? fallbackLineEntry.text.length : 0;
                const fallbackStartCol = Math.min(Math.max(selStart[1], 0), fallbackMaxLen);
                const fallbackEndCol = Math.min(fallbackStartCol + replacement.length, fallbackMaxLen);
                const fallbackCaretPos = [lineNum, fallbackEndCol];
                aceInstance.ace_performSelectionChange(fallbackCaretPos, fallbackCaretPos, false);
                aceInstance.ace_fastIncorp(10);
                return;
              }

              if (selEnd[0] !== selStart[0]) {
                selEnd[0] = selStart[0];
                selEnd[1] = cellEndCol;
              }

              if (selEnd[1] > cellEndCol) {
                selEnd[1] = Math.min(selEnd[1], cellEndCol);
              }

              if (selEnd[1] < selStart[1]) selEnd[1] = selStart[1];

              aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, replacement);

              const repAfter = aceInstance.ace_getRep();
              const lineEntryAfter = repAfter.lines.atIndex(lineNum);
              const maxLen = lineEntryAfter ? lineEntryAfter.text.length : 0;
              const startCol = Math.min(Math.max(selStart[1], 0), maxLen);
              const endCol = Math.min(startCol + replacement.length, maxLen);

              if (endCol > startCol) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [lineNum, startCol],
                  [lineNum, endCol],
                  [[ATTR_CELL, String(targetCellIndex)]]
                );
              }

              ed.ep_data_tables_applyMeta(
                lineNum,
                tableMetadata.tblId,
                tableMetadata.row,
                tableMetadata.cols,
                repAfter,
                ed,
                null,
                docManager
              );

              const newCaretPos = [lineNum, endCol];
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              aceInstance.ace_fastIncorp(10);

              const editor = ed.ep_data_tables_editor;
              if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
                const freshLineText = lineEntryAfter ? lineEntryAfter.text : '';
                const freshCells = freshLineText.split(DELIMITER);
                let freshOffset = 0;
                for (let i = 0; i < targetCellIndex; i++) {
                  freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
                }
                const newRelativePos = newCaretPos[1] - freshOffset;
                editor.ep_data_tables_last_clicked = {
                  lineNum,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
                };
              }
            }, 'mobileSuggestionBlocker', true);
          } catch (e) {
            console.error('[ep_data_tables:mobileSuggestionBlocker] Error applying predictive text:', e);
          }
        }, 0);
      };

      const softBreakNormalizer = (rawEvt) => {
        try {
          const e = rawEvt && (rawEvt.originalEvent || rawEvt);
          if (!e || e._epDataTablesNormalized) return;
          const t = e.inputType || '';
          const dataStr = typeof e.data === 'string' ? e.data : '';

          // If this NBSP is flanked by ZWSPs we are inside an image placeholder.
          // In that case leave it untouched so caret math stays correct.
          if (dataStr === '\u00A0' && rep && rep.selStart) {
            const lineText = rep.lines.atIndex(rep.selStart[0])?.text || '';
            const pos = rep.selStart[1]; // caret is before the NBSP
            if (lineText.slice(pos - 1, pos + 2) === '\u200B\u00A0\u200B') return;
          }

          const hasSoftWs = /[\r\n\u00A0]/.test(dataStr); // include NBSP (U+00A0)
          const isSoftBreak = t === 'insertParagraph' || t === 'insertLineBreak' || hasSoftWs;
          if (!isSoftBreak) return;

          const rep = ed.ace_getRep && ed.ace_getRep();
          if (!rep || !rep.selStart) return;
          const lineNum = rep.selStart[0];
          let metaStr = docManager && docManager.getAttributeOnLine ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON) : null;
          let meta = null;
          if (metaStr) { try { meta = JSON.parse(metaStr); } catch (_) {} }
          if (!meta) meta = getTableLineMetadata(lineNum, ed, docManager);
          if (!meta || typeof meta.cols !== 'number') return;

          e.preventDefault();
          if (typeof e.stopImmediatePropagation === 'function') e.stopImmediatePropagation();
          e._epDataTablesNormalized = true;
          setTimeout(() => {
            try {
              ed.ace_callWithAce((aceInstance) => {
                aceInstance.ace_fastIncorp(10);
                const rep2 = aceInstance.ace_getRep();
                const start = rep2.selStart;
                const end = rep2.selEnd;
                // Replace the attempted soft-break with a single space
                aceInstance.ace_performDocumentReplaceRange(start, end, ' ');
              }, 'softBreakNormalizer', true);
            } catch (err) { console.error('[ep_data_tables:softBreakNormalizer] error', err); }
          }, 0);
        } catch (_) {}
      };

      const desktopCompositionSuppressor = (rawEvt) => {
        try {
          const e = rawEvt && (rawEvt.originalEvent || rawEvt);
          if (!e || isAndroidUA() || isIOSUA()) return;
          if (!desktopComposition.active) return;
          const t = e.inputType || '';
          if (e.isComposing && typeof t === 'string' && t.startsWith('insert')) {
            try { e.preventDefault(); } catch (_) {}
            try { if (typeof e.stopImmediatePropagation === 'function') e.stopImmediatePropagation(); } catch (_) {}
            try { if (typeof e.stopPropagation === 'function') e.stopPropagation(); } catch (_) {}
            e._epDataTablesHandled = true;
          }
        } catch (_) {}
      };

      const desktopCompositionInputSuppressor = (rawEvt) => {
        try {
          const e = rawEvt && (rawEvt.originalEvent || rawEvt);
          if (!e || isAndroidUA() || isIOSUA()) return;
          const t = e.inputType || '';
          // Suppress any insert events while the desktop composition is active.
          if (desktopComposition.active && e.isComposing && typeof t === 'string' && t.startsWith('insert')) {
            try { if (typeof e.stopImmediatePropagation === 'function') e.stopImmediatePropagation(); } catch (_) {}
            try { if (typeof e.stopPropagation === 'function') e.stopPropagation(); } catch (_) {}
            e._epDataTablesHandled = true;
          }
          // Also suppress the first post-composition desktop 'input' that carries the committed text
          // (Chrome often fires input.insertCompositionText with isComposing=false after compositionend).
          if (!e.isComposing && typeof t === 'string' && t === 'insertCompositionText' && suppressNextInputCommit) {
            try { e.preventDefault && e.preventDefault(); } catch (_) {}
            try { if (typeof e.stopImmediatePropagation === 'function') e.stopImmediatePropagation(); } catch (_) {}
            try { if (typeof e.stopPropagation === 'function') e.stopPropagation(); } catch (_) {}
            e._epDataTablesHandled = true;
            suppressNextInputCommit = false;
          }
        } catch (_) {}
      };

      if ($inner && $inner.length > 0 && $inner[0].addEventListener) {
        const el = $inner[0];
        // COMMENTED OUT FOR PRODUCTION - Uncomment for debugging IME/composition issues
        // ['beforeinput','input','textInput','compositionstart','compositionupdate','compositionend','keydown','keyup'].forEach((t) => {
        //   el.addEventListener(t, (ev) => logIMEEvent(ev, 'capture'), true);
        // });
        el.addEventListener('beforeinput', softBreakNormalizer, true);
        el.addEventListener('beforeinput', desktopCompositionSuppressor, true);
        el.addEventListener('input', desktopCompositionInputSuppressor, true);
        el.addEventListener('textInput', desktopCompositionInputSuppressor, true);
      }

      if ($inner && $inner.length > 0 && $inner[0].addEventListener) {
        $inner[0].addEventListener('beforeinput', mobileSuggestionBlocker, true);
      }

      if (isAndroidUA && isAndroidUA() && $inner && $inner.length > 0 && $inner[0].addEventListener) {
        $inner[0].addEventListener('textInput', (evt) => {
          const s = typeof evt.data === 'string' ? evt.data : '';
          if (s && s.length > 1) {
            mobileSuggestionBlocker({
              inputType: 'insertText',
              isComposing: false,
              data: s,
              preventDefault: () => { try { evt.preventDefault(); } catch (_) {} },
              stopImmediatePropagation: () => { try { evt.stopImmediatePropagation(); } catch (_) {} },
            });
          }
        }, true);
      }

      // Disable Grammarly on body and iframes (spellcheck is disabled per-table in buildTableFromDelimitedHTML)
      try {
        const body = innerDocBody[0] || innerDocBody;
        if (body) {
          body.setAttribute('data-gramm', 'false');
          body.setAttribute('data-enable-grammarly', 'false');
        }
        
        const outerFrame = document.querySelector('iframe[name="ace_outer"]');
        const innerFrame = outerFrame?.contentDocument?.querySelector('iframe[name="ace_inner"]');
        if (outerFrame) {
          outerFrame.setAttribute('data-gramm', 'false');
          outerFrame.setAttribute('data-enable-grammarly', 'false');
        }
        if (innerFrame) {
          innerFrame.setAttribute('data-gramm', 'false');
          innerFrame.setAttribute('data-enable-grammarly', 'false');
        }
      } catch (_) {}
      
      if (!$inner || $inner.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: $inner is not valid. Cannot attach listeners.`);
        return;
      }

    // Cut handler - intercept cut on table lines to protect delimiters
    $inner.on('cut', (evt) => {
      const cutLogPrefix = '[ep_data_tables:cutHandler]';

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;

      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      const hasSelectionInRep = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);

      // Block multi-line cut to protect table structure
      if (selStart[0] !== selEnd[0]) {
        evt.preventDefault();
        return;
      }

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch {}
      }

      if (!tableMetadata) {
        tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      }

      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        return; // Not a table line, allow default cut
      }

      // Block collapsed selection cut on table lines
      if (!hasSelectionInRep) {
        evt.preventDefault();
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;

      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;

        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      // Clamp selections that include delimiter to cell boundary
      const wouldClampStart = targetCellIndex > 0 && selStart[1] === cellStartCol - DELIMITER.length;
      const wouldClampEnd = targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length;

      if (wouldClampStart) selStart[1] = cellStartCol;
      if (wouldClampEnd) selEnd[1] = cellEndCol;

      // Block if selection spans cell boundaries
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      evt.preventDefault();

      try {
        const selectedText = lineText.substring(selStart[1], selEnd[1]);

        // Copy to clipboard
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(selectedText).catch(() => {});
        } else {
          const textArea = document.createElement('textarea');
          textArea.value = selectedText;
          document.body.appendChild(textArea);
          textArea.select();
          try { document.execCommand('copy'); } catch (_) {}
          document.body.removeChild(textArea);
        }

        // Delete selected text and preserve table structure
        ed.ace_callWithAce((aceInstance) => {
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');

          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          // Insert space if cell became empty to preserve delimiters
          if (cellTextAfterDeletion.length === 0) {
            const insertPos = [lineNum, selStart[1]];
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');
            aceInstance.ace_performDocumentApplyAttributesToRange(
              insertPos, [insertPos[0], insertPos[1] + 1], [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

          const repAfterCut = aceInstance.ace_getRep();
          ed.ep_data_tables_applyMeta(lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, repAfterCut, ed, null, docManager);

          const newCaretPos = [lineNum, selStart[1]];
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
        }, 'tableCutTextOperations', true);
      } catch (error) {
        console.error(`${cutLogPrefix} ERROR during cut operation:`, error);
      }
    });

    $inner.on('beforeinput', (evt) => {
      const deleteLogPrefix = '[ep_data_tables:beforeinputDeleteHandler]';

      if (!evt.originalEvent.inputType || !evt.originalEvent.inputType.startsWith('delete')) {
        return;
      }

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        console.warn(`${deleteLogPrefix} Could not get rep or selStart.`);
        return;
      }
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      const isAndroidChrome = isAndroidUA();
      const inputType = (evt.originalEvent && evt.originalEvent.inputType) || '';

      const isCollapsed = (selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);
      if (isCollapsed && isAndroidChrome && (inputType === 'deleteContentBackward' || inputType === 'deleteContentForward')) {
        let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
        let tableMetadata = null;
        if (lineAttrString) { try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {} }
        if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
          return;
        }

        const lineText = rep.lines.atIndex(lineNum)?.text || '';
        const cells = lineText.split(DELIMITER);
        let currentOffset = 0;
        let targetCellIndex = -1;
        let cellStartCol = 0;
        let cellEndCol = 0;
        for (let i = 0; i < cells.length; i++) {
          const cellLength = cells[i]?.length ?? 0;
          const cellEndColThisIteration = currentOffset + cellLength;
          if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
            targetCellIndex = i;
            cellStartCol = currentOffset;
            cellEndCol = cellEndColThisIteration;
            break;
          }
          currentOffset += cellLength + DELIMITER.length;
        }

        if (targetCellIndex === -1) return;

        const isBackward = inputType === 'deleteContentBackward';
        const caretCol = selStart[1];

        if ((isBackward && caretCol === cellStartCol) || (!isBackward && caretCol === cellEndCol)) {
          evt.preventDefault();
          return;
        }

        evt.preventDefault();
        try {
          ed.ace_callWithAce((aceInstance) => {
            // Refresh representation to account for any prior synchronous edits
            aceInstance.ace_fastIncorp(10);
            const freshRep = aceInstance.ace_getRep();

            const freshLineEntry = freshRep.lines.atIndex(lineNum);
            const freshText = (freshLineEntry && freshLineEntry.text) || '';
            const freshCells = freshText.split(DELIMITER);

            let offset = 0;
            let freshCellStart = 0;
            let freshCellEnd = 0;
            for (let i = 0; i < freshCells.length; i++) {
              const len = freshCells[i]?.length ?? 0;
              const end = offset + len;
              if (i === targetCellIndex) {
                freshCellStart = offset;
                freshCellEnd = end;
                break;
              }
              offset += len + DELIMITER.length;
            }

            // If caret is flush against delimiter after refresh, abort to protect structure
            if ((isBackward && caretCol <= freshCellStart) || (!isBackward && caretCol >= freshCellEnd)) {
              return;
            }

            const delStart = isBackward ? [lineNum, caretCol - 1] : [lineNum, caretCol];
            const delEnd   = isBackward ? [lineNum, caretCol]     : [lineNum, caretCol + 1];
            aceInstance.ace_performDocumentReplaceRange(delStart, delEnd, '');

            const repAfter = aceInstance.ace_getRep();

            ed.ep_data_tables_applyMeta(
              lineNum,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfter,
              ed,
              null,
              docManager
            );

            const newCaretCol = isBackward ? caretCol - 1 : caretCol;
            const newCaretPos = [lineNum, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            aceInstance.ace_fastIncorp(10);

            if (ed.ep_data_tables_editor && ed.ep_data_tables_editor.ep_data_tables_last_clicked && ed.ep_data_tables_editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
              const newRelativePos = newCaretCol - cellStartCol;
              ed.ep_data_tables_editor.ep_data_tables_last_clicked = {
                lineNum: lineNum,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos,
              };
            }
          }, 'tableCollapsedDeleteHandler', true);
        } catch (error) {
          console.error(`${deleteLogPrefix} ERROR handling collapsed delete:`, error);
        }
        return;
      }

      if (isCollapsed) {
        return;
      }

      if (selStart[0] !== selEnd[0]) {
        evt.preventDefault();
        return;
      }

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (lineAttrString) {
        try {
          tableMetadata = JSON.parse(lineAttrString);
        } catch {}
      }

      if (!tableMetadata) {
        tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      }

      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        return;
      }


      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;

      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;

        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      // Clamp selections that include delimiter to cell boundary
      const wouldClampStart = targetCellIndex > 0 && selStart[1] === cellStartCol - DELIMITER.length;
      const wouldClampEnd = targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length;

      if (wouldClampStart) selStart[1] = cellStartCol;
      if (wouldClampEnd) selEnd[1] = cellEndCol;

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      evt.preventDefault();

      try {
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${deleteLogPrefix}[ace_callWithAceOps]`;

          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');

          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
            const insertPos = [lineNum, selStart[1]];
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            const attrStart = insertPos;
            const attrEnd   = [insertPos[0], insertPos[1] + 1];
            aceInstance.ace_performDocumentApplyAttributesToRange(
              attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

          const repAfterDelete = aceInstance.ace_getRep();

          ed.ep_data_tables_applyMeta(
            lineNum,
            tableMetadata.tblId,
            tableMetadata.row,
            tableMetadata.cols,
            repAfterDelete,
            ed,
            null,
            docManager
          );

          const newCaretAbsoluteCol = (cellTextAfterDeletion.length === 0) ? selStart[1] + 1 : selStart[1];
          const newCaretPos = [lineNum, newCaretAbsoluteCol];
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);

        }, 'tableDeleteTextOperations', true);

      } catch (error) {
        console.error(`${deleteLogPrefix} ERROR during delete operation:`, error);
      }
    });

    $inner.on('beforeinput', (evt) => {
      const insertLogPrefix = '[ep_data_tables:beforeinputInsertHandler]';
      const inputType = evt.originalEvent && evt.originalEvent.inputType || '';

      if (!inputType || !inputType.startsWith('insert')) return;

      if ((evt && evt._epDataTablesHandled) || (evt.originalEvent && evt.originalEvent._epDataTablesHandled)) return;

      if (!isAndroidUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
      }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;
      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;
        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      if (inputType === 'insertParagraph' || inputType === 'insertLineBreak') {
        evt.preventDefault();
        evt.stopPropagation();
        if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();
        setTimeout(() => {
          try {
            ed.ace_callWithAce((aceInstance) => {
              aceInstance.ace_fastIncorp(10);
              const freshRep = aceInstance.ace_getRep();
              const freshSelStart = freshRep.selStart;
              const freshSelEnd = freshRep.selEnd;
              aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, ' ');

              const afterRep = aceInstance.ace_getRep();
              const maxLen = Math.max(0, afterRep.lines.atIndex(lineNum)?.text?.length || 0);
              const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
              const endCol = Math.min(startCol + 1, maxLen);
              if (endCol > startCol) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [lineNum, startCol], [lineNum, endCol], [[ATTR_CELL, String(targetCellIndex)]]
                );
              }

              ed.ep_data_tables_applyMeta(
                lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols,
                afterRep, ed, null, docManager
              );

              const newCaretPos = [lineNum, endCol];
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              aceInstance.ace_fastIncorp(10);
            }, 'iosSoftBreakToSpace', true);
          } catch (e) {
            console.error(`${autoLogPrefix} ERROR replacing soft break:`, e);
          }
        }, 0);
        return;
      }

      if (targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }

      if (inputType === 'insertParagraph' || inputType === 'insertLineBreak') {
        evt.preventDefault();
        try {
          navigateToCellBelow(lineNum, targetCellIndex, tableMetadata, ed, docManager);
        } catch (e) { console.error(`${insertLogPrefix} Error navigating on line break:`, e); }
        return;
      }

      if (suppressBeforeInputInsertTextDuringComposition && inputType === 'insertText') {
        evt.preventDefault();
        return;
      }

      if (suppressNextBeforeInputInsertTextOnce && inputType === 'insertText') {
        suppressNextBeforeInputInsertTextOnce = false;
        evt.preventDefault();
        return;
      }

      if (isAndroidChromeComposition) return;

      const rawData = evt.originalEvent && typeof evt.originalEvent.data === 'string' ? evt.originalEvent.data : '';
      if (!rawData) return;

      let insertedText = normalizeSoftWhitespace(
        rawData
          .replace(new RegExp(DELIMITER, 'g'), ' ')
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
          .replace(/\s+/g, ' ')
      );

      if (insertedText.length === 0) {
        evt.preventDefault();
        return;
      }

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        setTimeout(() => {
          ed.ace_callWithAce((aceInstance) => {
            aceInstance.ace_fastIncorp(10);
            const freshRep = aceInstance.ace_getRep();
            const freshSelStart = freshRep.selStart;
            const freshSelEnd = freshRep.selEnd;

            aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

            const repAfterReplace = aceInstance.ace_getRep();
            const freshLineIndex = freshSelStart[0];
            const freshLineEntry = repAfterReplace.lines.atIndex(freshLineIndex);
            const maxLen = Math.max(0, (freshLineEntry && freshLineEntry.text) ? freshLineEntry.text.length : 0);
            const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
            const endColRaw = startCol + insertedText.length;
            const endCol = Math.min(endColRaw, maxLen);
            if (endCol > startCol) {
              aceInstance.ace_performDocumentApplyAttributesToRange(
                [freshLineIndex, startCol], [freshLineIndex, endCol], [[ATTR_CELL, String(targetCellIndex)]]
              );
            }

            ed.ep_data_tables_applyMeta(
              freshLineIndex,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );

            const newCaretCol = endCol;
            const newCaretPos = [freshLineIndex, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            aceInstance.ace_fastIncorp(10);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
              const freshLineText = (freshLineEntry && freshLineEntry.text) || '';
              const freshCells = freshLineText.split(DELIMITER);
              let freshOffset = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
              }
              const newRelativePos = newCaretCol - freshOffset;
              editor.ep_data_tables_last_clicked = {
                lineNum: freshLineIndex,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos,
              };
            }
          }, 'tableInsertTextOperations', true);
        }, 0);
      } catch (error) {
        console.error(`${insertLogPrefix} ERROR during insert handling:`, error);
      }
    });

    $inner.on('beforeinput', (evt) => {
      const genericLogPrefix = '[ep_data_tables:beforeinputInsertTextGeneric]';
      const nativeEvt = evt.originalEvent || evt;
      const inputType = (nativeEvt && nativeEvt.inputType) || '';
      if (!inputType || !inputType.startsWith('insert')) return;

      // If we just committed via the desktop compositionend pipeline, skip this first beforeinput to avoid double-insert.
      if (suppressNextBeforeinputCommitOnce) {
        suppressNextBeforeinputCommitOnce = false;
        evt.preventDefault();
        return;
      }

      if (evt._epDataTablesNormalized || (evt.originalEvent && evt.originalEvent._epDataTablesNormalized)) return;
      const isComposing = !!(nativeEvt && nativeEvt.isComposing);
      const isCompositionText = inputType === 'insertCompositionText';
      const isCompositionCommit = isCompositionText && !isComposing;
      if (!evt._epDataTablesHandled) evt._epDataTablesHandled = false;
      if (nativeEvt && !nativeEvt._epDataTablesHandled) nativeEvt._epDataTablesHandled = false;

      if (!isCompositionCommit && isComposing) {
        logCompositionEvent('beforeinput-skip-active-composition', evt, { reason: 'active-composition', inputType });
        return;
      }
      if (isAndroidUA() || isIOSUA()) return;

      const dataPreview = typeof nativeEvt?.data === 'string' ? nativeEvt.data : '';
      logCompositionEvent('beforeinput', evt, {
        isCompositionCommit,
        dataPreview,
      });

      evt._epDataTablesHandled = true;
      if (nativeEvt) nativeEvt._epDataTablesHandled = true;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
      }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') return;
      const rawData = typeof nativeEvt?.data === 'string' ? nativeEvt.data : ' ';

      const insertedText = normalizeSoftWhitespace(
        rawData
          .replace(/[\u00A0\r\n\t]/g, ' ') // NBSP sanitized back to space for stability
          .replace(new RegExp(DELIMITER, 'g'), ' ')
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
          .replace(/\s+/g, ' ')
      );

      if (!insertedText) { evt.preventDefault(); return; }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;
      for (let i = 0; i < cells.length; i++) {
        const len = cells[i]?.length ?? 0;
        const end = currentOffset + len;
        if (selStart[1] >= currentOffset && selStart[1] <= end) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = end;
          break;
        }
        currentOffset += len + DELIMITER.length;
      }
      // Clamp selection when it includes the trailing delimiter so replacement stays within the cell
      if (targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) { evt.preventDefault(); logCompositionEvent('beforeinput-abort-outside-cell', evt, { selStart, selEnd, cellStartCol, cellEndCol }); return; }

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        logCompositionEvent('beforeinput-commit-routing', evt, {
          lineNum,
          targetCellIndex,
          insertedText,
        });
        ed.ace_callWithAce((ace) => {
          ace.ace_fastIncorp(10);
          const freshRep = ace.ace_getRep();
          const freshSelStart = freshRep.selStart;
          const freshSelEnd = freshRep.selEnd;

          // Validate line is safe to modify before replacement (prevents keyToNodeMap errors)
          if (!isLineSafeToModify(freshRep, freshSelStart[0], '[ep_data_tables:beforeinput-commit]')) {
            console.warn('[ep_data_tables:beforeinput-commit] line not safe to modify, skipping');
            return;
          }
          ace.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

          const afterRep = ace.ace_getRep();
          const lineEntry = afterRep.lines.atIndex(lineNum);
          const maxLen = lineEntry ? lineEntry.text.length : 0;
          const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
          const endCol = Math.min(startCol + insertedText.length, maxLen);
          if (endCol > startCol) {
            ace.ace_performDocumentApplyAttributesToRange([lineNum, startCol], [lineNum, endCol], [[ATTR_CELL, String(targetCellIndex)]]);
          }

          ed.ep_data_tables_applyMeta(lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, afterRep, ed, null, docManager);

          const newCaretPos = [lineNum, endCol];
          ace.ace_performSelectionChange(newCaretPos, newCaretPos, false);
          logCompositionEvent('beforeinput-commit-applied', evt, {
            lineNum,
            targetCellIndex,
            newCaretPos,
          });
        }, 'tableGenericInsertText', true);
      } catch (e) {
        console.error(`${genericLogPrefix} ERROR handling generic insertText:`, e);
      }
    });

    $inner.on('beforeinput', (evt) => {
      const autoLogPrefix = '[ep_data_tables:beforeinputAutoReplaceHandler]';
      const inputType = (evt.originalEvent && evt.originalEvent.inputType) || '';

      if ((evt && evt._epDataTablesHandled) || (evt.originalEvent && evt.originalEvent._epDataTablesHandled)) return;

      if (!isIOSUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      const dataStr = (evt.originalEvent && typeof evt.originalEvent.data === 'string') ? evt.originalEvent.data : '';
      const hasSelection = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);
      const looksLikeIOSAutoReplace = inputType === 'insertText' && dataStr.length > 1;
      const insertTextNull = inputType === 'insertText' && dataStr === '' && !hasSelection;
      const shouldHandle = INPUTTYPE_REPLACEMENT_TYPES.has(inputType) || looksLikeIOSAutoReplace || (inputType === 'insertText' && (hasSelection || insertTextNull));
      if (!shouldHandle) return;

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
      }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;
      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;
        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      if (targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      let insertedText = dataStr;
      if (!insertedText) {
        if (insertTextNull) {
          evt.preventDefault();
          evt.stopPropagation();
          if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

          setTimeout(() => {
            try {
              ed.ace_callWithAce((aceInstance) => {
                aceInstance.ace_fastIncorp(10);
                const freshRep = aceInstance.ace_getRep();
                const freshSelStart = freshRep.selStart;
                const freshSelEnd = freshRep.selEnd;
                aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, ' ');

                const afterRep = aceInstance.ace_getRep();
                const maxLen = Math.max(0, afterRep.lines.atIndex(lineNum)?.text?.length || 0);
                const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
                const endCol = Math.min(startCol + 1, maxLen);
                if (endCol > startCol) {
                  aceInstance.ace_performDocumentApplyAttributesToRange(
                    [lineNum, startCol], [lineNum, endCol], [[ATTR_CELL, String(targetCellIndex)]]
                  );
                }

                ed.ep_data_tables_applyMeta(
                  lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols,
                  afterRep, ed, null, docManager
                );

                const newCaretPos = [lineNum, endCol];
                aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                aceInstance.ace_fastIncorp(10);
              }, 'iosPredictiveCommit', true);
            } catch (e) {
              console.error(`${autoLogPrefix} ERROR fixing predictive commit:`, e);
            }
          }, 0);
          return;
        } else {
          if (INPUTTYPE_REPLACEMENT_TYPES.has(inputType) || hasSelection) {
            evt.preventDefault();
            evt.stopPropagation();
            if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();
          }
          return;
        }
      }

      insertedText = normalizeSoftWhitespace(
        insertedText
          .replace(new RegExp(DELIMITER, 'g'), ' ')
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
      );

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        setTimeout(() => {
          ed.ace_callWithAce((aceInstance) => {
            aceInstance.ace_fastIncorp(10);
            const freshRep = aceInstance.ace_getRep();
            const freshSelStart = freshRep.selStart;
            const freshSelEnd = freshRep.selEnd;

            aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

            const repAfterReplace = aceInstance.ace_getRep();
            const freshLineIndex = freshSelStart[0];
            const freshLineEntry = repAfterReplace.lines.atIndex(freshLineIndex);
            const maxLen = Math.max(0, (freshLineEntry && freshLineEntry.text) ? freshLineEntry.text.length : 0);
            const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
            const endColRaw = startCol + insertedText.length;
            const endCol = Math.min(endColRaw, maxLen);
            if (endCol > startCol) {
              aceInstance.ace_performDocumentApplyAttributesToRange(
                [freshLineIndex, startCol], [freshLineIndex, endCol], [[ATTR_CELL, String(targetCellIndex)]]
              );
            }

            ed.ep_data_tables_applyMeta(
              freshLineIndex,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );

            const newCaretCol = endCol;
            const newCaretPos = [freshLineIndex, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            aceInstance.ace_fastIncorp(10);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
              const freshLineText = (freshLineEntry && freshLineEntry.text) || '';
              const freshCells = freshLineText.split(DELIMITER);
              let freshOffset = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
              }
              const newRelativePos = newCaretCol - freshOffset;
              editor.ep_data_tables_last_clicked = {
                lineNum: freshLineIndex,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos,
              };
            }
          }, 'tableAutoReplaceTextOperations', true);
        }, 0);
      } catch (error) {
        console.error(`${autoLogPrefix} ERROR during auto-replace handling:`, error);
      }
    });

    $inner.on('compositionstart', (evt) => {
      // Set global composition flag (for Android path)
      __epDT_compositionActive = true;
      if (!isAndroidUA()) return;
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const lineNum = rep.selStart[0];
      let meta = null; let s = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      if (s) { try { meta = JSON.parse(s); } catch (_) {} }
      if (!meta) meta = getTableLineMetadata(lineNum, ed, docManager);
      if (!meta || typeof meta.cols !== 'number') return;
      isAndroidChromeComposition = true;
      handledCurrentComposition = false;
      suppressBeforeInputInsertTextDuringComposition = false;
    });
    $inner.on('compositionupdate', (evt) => {
      const compLogPrefix = '[ep_data_tables:compositionHandler]';

      if (!isAndroidUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) { try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {} }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') return;

      const d = evt.originalEvent && typeof evt.originalEvent.data === 'string' ? evt.originalEvent.data : '';
      if (evt.type === 'compositionupdate') {
        const isWhitespaceOnly = d && normalizeSoftWhitespace(d).trim() === '';
        if (!isWhitespaceOnly) return;

        const lineText = rep.lines.atIndex(lineNum)?.text || '';
        const cells = lineText.split(DELIMITER);
        let currentOffset = 0;
        let targetCellIndex = -1;
        let cellStartCol = 0;
        let cellEndCol = 0;
        for (let i = 0; i < cells.length; i++) {
          const cellLength = cells[i]?.length ?? 0;
          const cellEndColThisIteration = currentOffset + cellLength;
          if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
            targetCellIndex = i;
            cellStartCol = currentOffset;
            cellEndCol = cellEndColThisIteration;
            break;
          }
          currentOffset += cellLength + DELIMITER.length;
        }
        if (targetCellIndex === -1 || selEnd[1] > cellEndCol) return;

        evt.preventDefault();
        evt.stopPropagation();
        if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

        let insertedText = d
          .replace(/[\u00A0\r\n\t]/g, ' ')
          .replace(/\s+/g, ' ');
        if (insertedText.length === 0) insertedText = ' ';

        try {
          setTimeout(() => {
            ed.ace_callWithAce((aceInstance) => {
              aceInstance.ace_fastIncorp(10);
              const freshRep = aceInstance.ace_getRep();
              const freshSelStart = freshRep.selStart;
              const freshSelEnd = freshRep.selEnd;
              aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

              const repAfterReplace = aceInstance.ace_getRep();
              const freshLineIndex = freshSelStart[0];
              const freshLineEntry = repAfterReplace.lines.atIndex(freshLineIndex);
              const maxLen = Math.max(0, (freshLineEntry && freshLineEntry.text) ? freshLineEntry.text.length : 0);
              const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
              const endColRaw = startCol + insertedText.length;
              const endCol = Math.min(endColRaw, maxLen);
              if (endCol > startCol) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [freshLineIndex, startCol], [freshLineIndex, endCol], [[ATTR_CELL, String(targetCellIndex)]]
                );
              }

              ed.ep_data_tables_applyMeta(
                freshLineIndex,
                tableMetadata.tblId,
                tableMetadata.row,
                tableMetadata.cols,
                repAfterReplace,
                ed,
                null,
                docManager
              );

              const newCaretCol = endCol;
              const newCaretPos = [freshLineIndex, newCaretCol];
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              aceInstance.ace_fastIncorp(10);
              if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
                const freshLineText = (freshLineEntry && freshLineEntry.text) || '';
                const freshCells = freshLineText.split(DELIMITER);
                let freshOffset = 0;
                for (let i = 0; i < targetCellIndex; i++) {
                  freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
                }
                const newRelativePos = newCaretCol - freshOffset;
                editor.ep_data_tables_last_clicked = {
                  lineNum: freshLineIndex,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
                };
              }
            }, 'tableCompositionSpaceInsert', true);
          }, 0);
          suppressBeforeInputInsertTextDuringComposition = true;
        } catch (error) {
          console.error(`${compLogPrefix} ERROR inserting space during composition:`, error);
        }
      }
    });

    $inner.on('compositionend', (evt) => {
      // Clear global composition flag and set cooldown timer (for Android path)
      __epDT_compositionActive = false;
      __epDT_lastCompositionEndTime = Date.now();

      if (isAndroidChromeComposition) {
        logCompositionEvent('compositionend-android-handler', evt);
        isAndroidChromeComposition = false;
        handledCurrentComposition = false;
        suppressBeforeInputInsertTextDuringComposition = false;
      }
    });

    $inner.on('compositionend', (evt) => {
      // Clear global composition flag and set cooldown timer
      __epDT_compositionActive = false;
      __epDT_lastCompositionEndTime = Date.now();

      if (isAndroidUA() || isIOSUA()) return;
      const compLogPrefix = '[ep_data_tables:compositionEndDesktop]';
      const nativeEvt = evt.originalEvent || evt;
      const dataPreview = typeof nativeEvt?.data === 'string' ? nativeEvt.data : '';
      logCompositionEvent('compositionend-desktop-fired', evt, { data: dataPreview });

      // Capture desktopComposition state before async ops
      // which would overwrite desktopComposition and cause us to use the wrong state.
      const capturedComposition = desktopComposition ? { ...desktopComposition } : null;
      
        // Prevent the immediate post-composition input commit from running; we pipeline instead
        suppressNextInputCommit = true;
      requestAnimationFrame(() => {
        try {
          ed.ace_callWithAce((aceInstance) => {
            // CRITICAL GUARD: Only run table-related pipeline if we're actually editing a table
            // Check if compositionstart found a table (snapshotMeta set) or if DOM selection is in a table
            const domTargetCheck = getDomCellTargetFromSelection();
            const domIsInTable = !!(domTargetCheck && domTargetCheck.tblId);
            const compositionFoundTable = !!(capturedComposition && capturedComposition.snapshotMeta);
            
            if (!compositionFoundTable && !domIsInTable) {
              // User is NOT editing in a table - let normal Etherpad handle this
              logCompositionEvent('compositionend-desktop-skipped-not-in-table', evt, {
                compositionFoundTable,
                domIsInTable,
                lineNum: capturedComposition?.lineNum,
              });
              return;
            }
            
            // Pipeline: Apply committed IME string synchronously to the target cell
            const commitStrRaw = typeof nativeEvt?.data === 'string' ? nativeEvt.data : '';
            const commitStr = sanitizeCellContent(commitStrRaw || '');
            // Only suppress the next beforeinput commit if the IME provided a non-empty commit string.
            // Normalize using the same soft-whitespace rules the editor uses.
            const willCommit = typeof commitStrRaw === 'string' && normalizeSoftWhitespace(commitStrRaw).trim().length > 0;
            if (willCommit) suppressNextBeforeinputCommitOnce = true;
              const repNow = aceInstance.ace_getRep();
              const caret = repNow && repNow.selStart;
              if (!caret && (capturedComposition && typeof capturedComposition.lineNum !== 'number')) return;
              // Prefer the line captured at compositionstart to avoid caret drift.
              const pipelineLineNum = (capturedComposition && typeof capturedComposition.lineNum === 'number')
                ? capturedComposition.lineNum
                : caret[0];
            let metadata = getTableMetadataForLine(pipelineLineNum);
              let entry = repNow.lines.atIndex(pipelineLineNum);
            if (!entry) {
              logCompositionEvent('compositionend-desktop-no-line-entry', evt, { lineNum: pipelineLineNum });
              return;
            }
              // If tblId was captured at start, and current metadata doesn't match, relocate the line by tblId.
              if (capturedComposition && capturedComposition.tblId) {
                const currentTblId = metadata && metadata.tblId ? metadata.tblId : null;
                if (currentTblId !== capturedComposition.tblId) {
                  const relocatedLine = findLineNumByTblId(capturedComposition.tblId);
                  if (relocatedLine >= 0) {
                    const relocatedEntry = repNow.lines.atIndex(relocatedLine);
                    if (relocatedEntry) {
                      metadata = getTableMetadataForLine(relocatedLine) || metadata;
                      entry = relocatedEntry;
                    }
                  }
                }
              }
            // Fallback: derive metadata from DOM if missing (first-row/empty-cell cases)
            if (!metadata || typeof metadata.cols !== 'number') {
              try {
                const tableNode = entry.lineNode &&
                  entry.lineNode.querySelector &&
                  entry.lineNode.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
                if (tableNode) {
                  const domTblId = tableNode.getAttribute('data-tblId') || tableNode.getAttribute('data-tblid');
                  const domRowAttr = tableNode.getAttribute('data-row');
                  const tr = tableNode.querySelector('tbody > tr');
                  const domCols = tr ? tr.children.length : 0;
                  if (domTblId && domCols > 0) {
                    metadata = {
                      tblId: domTblId,
                      row: domRowAttr ? parseInt(domRowAttr, 10) : 0,
                      cols: domCols,
                    };
                    logCompositionEvent('compositionend-desktop-dom-meta', evt, {
                      lineNum: pipelineLineNum, tblId: metadata.tblId, cols: metadata.cols,
                    });
                  }
                }
              } catch (_) {}
            }
            
            // ENHANCED FALLBACK: If compositionstart missed metadata, try multiple recovery strategies
            if (!metadata || typeof metadata.cols !== 'number') {
              // Strategy 1: Use DOM selection to find the table we're actually in
              try {
                const domTarget = getDomCellTargetFromSelection();
                if (domTarget && domTarget.tblId && typeof domTarget.lineNum === 'number') {
                  const targetEntry = repNow.lines.atIndex(domTarget.lineNum);
                  if (targetEntry) {
                    const tableNode = targetEntry.lineNode?.querySelector(
                      `table.dataTable[data-tblId="${domTarget.tblId}"], table.dataTable[data-tblid="${domTarget.tblId}"]`
                    );
                    if (tableNode) {
                      const tr = tableNode.querySelector('tbody > tr');
                      const domCols = tr ? tr.children.length : 0;
                      if (domCols > 0) {
                        metadata = {
                          tblId: domTarget.tblId,
                          row: parseInt(tableNode.getAttribute('data-row') || '0', 10),
                          cols: domCols,
                        };
                        entry = targetEntry;
                        logCompositionEvent('compositionend-desktop-selection-fallback', evt, {
                          originalLine: pipelineLineNum,
                          recoveredLine: domTarget.lineNum,
                          tblId: metadata.tblId,
                          cols: metadata.cols,
                        });
                      }
                    }
                  }
                }
              } catch (_) {}
            }
            
            // Strategy 2: Use last_clicked state if available
            if ((!metadata || typeof metadata.cols !== 'number') && ed.ep_data_tables_last_clicked) {
              try {
                const lastClicked = ed.ep_data_tables_last_clicked;
                if (lastClicked.tblId && typeof lastClicked.lineNum === 'number') {
                  const targetEntry = repNow.lines.atIndex(lastClicked.lineNum);
                  if (targetEntry) {
                    const tableNode = targetEntry.lineNode?.querySelector(
                      `table.dataTable[data-tblId="${lastClicked.tblId}"], table.dataTable[data-tblid="${lastClicked.tblId}"]`
                    );
                    if (tableNode) {
                      const tr = tableNode.querySelector('tbody > tr');
                      const domCols = tr ? tr.children.length : 0;
                      if (domCols > 0) {
                        metadata = {
                          tblId: lastClicked.tblId,
                          row: parseInt(tableNode.getAttribute('data-row') || '0', 10),
                          cols: domCols,
                        };
                        entry = targetEntry;
                        logCompositionEvent('compositionend-desktop-lastclicked-fallback', evt, {
                          originalLine: pipelineLineNum,
                          recoveredLine: lastClicked.lineNum,
                          tblId: metadata.tblId,
                          cols: metadata.cols,
                          cellIndex: lastClicked.cellIndex,
                        });
                      }
                    }
                  }
                }
              } catch (_) {}
            }
            
            if (!metadata || typeof metadata.cols !== 'number') {
              logCompositionEvent('compositionend-desktop-no-meta', evt, { lineNum: pipelineLineNum });
              return;
            }
              const cellsNow = (entry.text || '').split(DELIMITER);
              while (cellsNow.length < metadata.cols) cellsNow.push(' ');
              // Prefer the cell index captured at compositionstart; otherwise compute using RAW mapping.
              let idx = (capturedComposition && capturedComposition.cellIndex >= 0)
                ? capturedComposition.cellIndex
                : (() => {
                    const selCol = (capturedComposition && capturedComposition.start) ? capturedComposition.start[1] : (caret ? caret[1] : 0);
                    const rawMap = computeTargetCellIndexFromRaw(entry, selCol);
                    return rawMap.index;
                  })();
            if (idx < 0) idx = Math.min(metadata.cols - 1, 0);

            // Compute relative selection in cell
              let baseOffset = 0;
              for (let i = 0; i < idx; i++) baseOffset += (cellsNow[i]?.length ?? 0) + DELIMITER.length;
              const sColAbs = (capturedComposition && capturedComposition.start) ? capturedComposition.start[1] : (caret ? caret[1] : 0);
              const eColAbs = (capturedComposition && capturedComposition.end) ? capturedComposition.end[1] : sColAbs;
              const currentCellLenRaw = cellsNow[idx]?.length ?? 0;
              const sCol = Math.max(0, sColAbs - baseOffset);
              const eCol = Math.max(sCol, Math.min(eColAbs - baseOffset, currentCellLenRaw));

              applyCellEditMinimal({
                    lineNum: pipelineLineNum,
                    tableMetadata: metadata,
                    cellIndex: idx,
                    relStart: sCol,
                    relEnd: eCol,
                    insertText: commitStr,
                    evt,
                    logTag: 'compositionend-desktop-pipeline-minimal',
                aceInstance,
              });

              // Position cursor at end of inserted text
              // Also store this position for recovery in the next compositionstart.
              try {
                const repAfterInsert = aceInstance.ace_getRep();
                const lineEntryAfter = repAfterInsert.lines.atIndex(pipelineLineNum);
                const lineTextAfter = lineEntryAfter?.text || '';
                const cellsAfter = lineTextAfter.split(DELIMITER);
                
                // Calculate where the cursor should be after insertion
                let expectedBaseOffset = 0;
                for (let i = 0; i < idx; i++) {
                  expectedBaseOffset += (cellsAfter[i]?.length ?? 0) + DELIMITER.length;
                }
                // Position cursor at: start of cell + relative start + length of inserted text
                const insertedTextLen = commitStr.length;
                const expectedCursorCol = expectedBaseOffset + sCol + insertedTextLen;
                
                // Explicitly set the selection to the expected position
                aceInstance.ace_performSelectionChange(
                  [pipelineLineNum, expectedCursorCol],
                  [pipelineLineNum, expectedCursorCol],
                  false
                );
                
                // Store the expected position for recovery in next compositionstart
                desktopComposition._expectedCursor = {
                  lineNum: pipelineLineNum,
                  col: expectedCursorCol,
                  cellIndex: idx,
                  tblId: metadata.tblId,
                  timestamp: Date.now(),
                };
                
                console.debug('[ep_data_tables:compositionend] cursor positioned', {
                  lineNum: pipelineLineNum,
                  col: expectedCursorCol,
                  cellIndex: idx,
                  insertedLen: insertedTextLen,
                });
              } catch (cursorErr) {
                console.debug('[ep_data_tables:compositionend] cursor positioning error', cursorErr?.message);
              }

            // Post-composition orphan detection and repair using snapshot
            setTimeout(() => {
              // Pre-check editor state
              // (keyToNodeMap.get(...) is undefined error)
              try {
                const preCheckRep = ed.ace_getRep && ed.ace_getRep();
                if (!preCheckRep || !preCheckRep.lines) {
                  console.debug('[ep_data_tables:compositionend-orphan-repair] skipped - rep unavailable in pre-check');
                  return;
                }
                // Additional sanity check - verify we can access line count without error
                const preCheckLineCount = preCheckRep.lines.length();
                if (typeof preCheckLineCount !== 'number' || preCheckLineCount < 1) {
                  console.debug('[ep_data_tables:compositionend-orphan-repair] skipped - invalid line count in pre-check');
                  return;
                }
              } catch (preCheckErr) {
                console.debug('[ep_data_tables:compositionend-orphan-repair] skipped - pre-check failed', preCheckErr?.message);
                return;
              }
              
              // Wrap entire ace_callWithAce in try-catch to handle internal Etherpad errors
              try {
                ed.ace_callWithAce((ace2) => {
                  try {
                  const repPost = ace2.ace_getRep();
                  if (!repPost || !repPost.lines || !docManager) return;
                  
                  const snapshotMeta = capturedComposition?.snapshotMeta;
                  const snapshotCells = capturedComposition?.snapshot;
                  const targetTblId = snapshotMeta?.tblId || metadata?.tblId;
                  const targetRow = snapshotMeta?.row ?? metadata?.row ?? 0;
                  
                  if (!targetTblId) return;
                  
                  // Find all lines with the same tblId/row (attribute-based AND DOM-based detection)
                  const matchingLines = [];
                  const seenLines = new Set();
                  const totalLines = repPost.lines.length();
                  for (let li = 0; li < totalLines; li++) {
                    // Method 1: Check line attribute
                    let attrStr = null;
                    try { attrStr = docManager.getAttributeOnLine(li, ATTR_TABLE_JSON); } catch (_) {}
                    if (attrStr) {
                      try {
                        const m = JSON.parse(attrStr);
                        if (m && m.tblId === targetTblId && m.row === targetRow) {
                          const entry = repPost.lines.atIndex(li);
                          matchingLines.push({ lineNum: li, text: entry?.text || '', meta: m, source: 'attr' });
                          seenLines.add(li);
                        }
                      } catch (_) {}
                    }
                    
                    // Method 2: DOM-based detection - find lines with tbljson-* OR tblCell-* class but no table
                    // NOTE: Browser may strip tbljson-* during IME but keep tblCell-*, so check both!
                    if (seenLines.has(li)) continue;
                    try {
                      const lineEntry = repPost.lines.atIndex(li);
                      const lineNode = lineEntry?.lineNode;
                      if (lineNode) {
                        const hasTable = lineNode.querySelector('table.dataTable');
                        if (!hasTable) {
                          // Check for tbljson-* spans (can decode metadata)
                          const tbljsonSpan = lineNode.querySelector('[class*="tbljson-"]');
                          if (tbljsonSpan) {
                            for (const cls of tbljsonSpan.classList) {
                              if (cls.startsWith('tbljson-')) {
                                try {
                                  const decoded = JSON.parse(atob(cls.substring(8)));
                                  if (decoded && decoded.tblId === targetTblId && decoded.row === targetRow) {
                                    const orphanText = lineEntry?.text || '';
                                    matchingLines.push({
                                      lineNum: li,
                                      text: orphanText,
                                      meta: decoded,
                                      source: 'dom-class-tbljson',
                                    });
                                    seenLines.add(li);
                                    console.debug('[ep_data_tables:compositionend-orphan-repair] tbljson DOM orphan found', {
                                      lineNum: li, tblId: decoded.tblId, row: decoded.row,
                                    });
                                    break;
                                  }
                                } catch (_) {}
                              }
                            }
                          }
                          
                          // Also check for tblCell-* spans (no encoded metadata, assume it's our target)
                          if (!seenLines.has(li)) {
                            const tblCellSpan = lineNode.querySelector('[class*="tblCell-"]');
                            if (tblCellSpan) {
                              const orphanText = lineEntry?.text || '';
                              // Only add if the text contains delimiters (suggests table-related content)
                              if (orphanText.includes(DELIMITER) || orphanText.trim().length > 0) {
                                matchingLines.push({
                                  lineNum: li,
                                  text: orphanText,
                                  meta: { tblId: targetTblId, row: targetRow, cols: metadata?.cols || 3 },
                                  source: 'dom-class-tblCell',
                                });
                                seenLines.add(li);
                                console.debug('[ep_data_tables:compositionend-orphan-repair] tblCell DOM orphan found', {
                                  lineNum: li, targetTblId, targetRow, textLen: orphanText.length,
                                });
                              }
                            }
                          }
                        }
                      }
                    } catch (_) {}
                  }
                  
                  // If more than one line matches, we have orphans - need to merge
                  if (matchingLines.length > 1) {
                    console.warn('[ep_data_tables:compositionend-orphan-repair] detected orphan lines', {
                      targetTblId, targetRow, lineCount: matchingLines.length,
                      lines: matchingLines.map(l => ({ line: l.lineNum, textLen: l.text.length })),
                      hasSnapshot: !!snapshotCells,
                    });
                    
                    // Find primary line using live DOM query
                    let primaryLine = null;
                    let orphans = [];
                    let lineWithTableIdx = -1;
                    
                    // Use live DOM query to find which ace-line actually has our table
                    try {
                      const innerDoc = $inner[0]?.ownerDocument || document;
                      const editorBody = innerDoc.getElementById('innerdocbody') || innerDoc.body;
                      if (editorBody) {
                        const allAceLines = editorBody.querySelectorAll('div.ace-line');
                        for (let ai = 0; ai < allAceLines.length; ai++) {
                          const aceLine = allAceLines[ai];
                          const tableEl = aceLine.querySelector(
                            `table.dataTable[data-tblId="${targetTblId}"][data-row="${targetRow}"], ` +
                            `table.dataTable[data-tblid="${targetTblId}"][data-row="${targetRow}"]`
                          );
                          if (tableEl) {
                            lineWithTableIdx = ai;
                            console.debug('[ep_data_tables:compositionend-orphan-repair] LIVE DOM found table', {
                              domIndex: ai, aceLineId: aceLine.id,
                            });
                            break;
                          }
                        }
                      }
                    } catch (domQueryErr) {
                      console.error('[ep_data_tables:compositionend-orphan-repair] live DOM query error', domQueryErr);
                    }
                    
                    // Find the primary line: prefer the one with the table, else lowest line number
                    if (lineWithTableIdx >= 0) {
                      // Find matching line for this DOM index
                      primaryLine = matchingLines.find(ml => ml.lineNum === lineWithTableIdx);
                      if (primaryLine) {
                        orphans = matchingLines.filter(ml => ml.lineNum !== lineWithTableIdx);
                        console.debug('[ep_data_tables:compositionend-orphan-repair] primary set via live DOM', {
                          lineNum: primaryLine.lineNum,
                        });
                      }
                    }
                    
                    // Fallback: if no line has a table, use the lowest line number
                    if (!primaryLine) {
                      matchingLines.sort((a, b) => a.lineNum - b.lineNum);
                      primaryLine = matchingLines[0];
                      orphans = matchingLines.slice(1);
                      console.debug('[ep_data_tables:compositionend-orphan-repair] no table DOM found, using lowest line', {
                        lineNum: primaryLine.lineNum,
                      });
                    }
                    
                    const expectedCols = primaryLine.meta.cols || (snapshotMeta?.cols) || 3;
                    
                    // Start with snapshot cells if available, otherwise use primary line
                    const baseCells = snapshotCells ? snapshotCells.slice() : primaryLine.text.split(DELIMITER);
                    while (baseCells.length < expectedCols) baseCells.push(' ');
                    
                    // If we have a valid committed string and cell index, apply it to the base
                    if (commitStr && commitStr.trim() && idx >= 0 && idx < expectedCols) {
                      // The committed text should be in the target cell
                      // Check if it's already there
                      if (!baseCells[idx].includes(commitStr.trim())) {
                        baseCells[idx] = sanitizeCellContent(baseCells[idx].trim() + commitStr.trim()) || ' ';
                      }
                    }
                    
                    // Merge content from orphan lines (preserve data, never delete content)
                    for (const orphan of orphans) {
                      const orphanSegs = orphan.text.split(DELIMITER);
                      for (let ci = 0; ci < Math.min(orphanSegs.length, expectedCols); ci++) {
                        const orphanContent = sanitizeCellContent(orphanSegs[ci] || '');
                        if (!orphanContent || orphanContent.trim() === '') continue;
                        // If base cell doesn't already contain this content, merge it
                        const baseTrimmed = (baseCells[ci] || '').trim();
                        const orphanTrimmed = orphanContent.trim();
                        if (baseTrimmed.includes(orphanTrimmed)) continue;
                        if (!baseTrimmed) {
                          baseCells[ci] = orphanContent;
                        } else {
                          baseCells[ci] = baseTrimmed + orphanTrimmed;
                        }
                      }
                    }
                    
                    // Sanitize all cells
                    for (let ci = 0; ci < expectedCols; ci++) {
                      baseCells[ci] = sanitizeCellContent(baseCells[ci] || ' ');
                    }
                    
                    const mergedText = baseCells.join(DELIMITER);
                    
                    // Update primary line with merged content
                    // First validate the line is safe to modify (prevents keyToNodeMap errors)
                    if (!isLineSafeToModify(repPost, primaryLine.lineNum, '[ep_data_tables:compositionend-orphan-repair] primary')) {
                      console.warn('[ep_data_tables:compositionend-orphan-repair] primary line not safe to modify, aborting merge');
                    } else {
                    const currentPrimaryText = repPost.lines.atIndex(primaryLine.lineNum)?.text || '';
                    if (mergedText !== currentPrimaryText) {
                      console.debug('[ep_data_tables:compositionend-orphan-repair] applying merged text', {
                        primaryLine: primaryLine.lineNum, from: currentPrimaryText.length, to: mergedText.length,
                      });
                      ace2.ace_performDocumentReplaceRange(
                        [primaryLine.lineNum, 0],
                        [primaryLine.lineNum, currentPrimaryText.length],
                        mergedText
                      );
                      
                      // Re-apply cell attributes
                      let offset = 0;
                      for (let ci = 0; ci < baseCells.length; ci++) {
                        const cellLen = baseCells[ci].length;
                        if (cellLen > 0) {
                          ace2.ace_performDocumentApplyAttributesToRange(
                            [primaryLine.lineNum, offset],
                            [primaryLine.lineNum, offset + cellLen],
                            [[ATTR_CELL, String(ci)]]
                          );
                        }
                        offset += cellLen;
                        if (ci < baseCells.length - 1) offset += DELIMITER.length;
                      }
                    }
                    } // end isLineSafeToModify check for primary line
                    
                    // Delete orphan lines bottom-up (content already merged)
                    if (!isDestructiveOperationSafe('compositionend-orphan-repair orphan removal')) {
                      console.debug('[ep_data_tables:compositionend-orphan-repair] skipping orphan removal (safe mode)');
                    } else {
                    orphans.sort((a, b) => b.lineNum - a.lineNum);
                    let orphansRemoved = 0;
                    let deletionOffset = 0;

                    for (const orphan of orphans) {
                      try {
                        // Re-fetch rep to get current state after previous deletions
                        const repCheck = ace2.ace_getRep();

                        // Lines are deleted bottom-up so indices above are stable
                        const adjustedLineNum = orphan.lineNum;

                        const totalLines = repCheck.lines.length();
                        if (adjustedLineNum >= totalLines || adjustedLineNum < 0) {
                          console.debug('[ep_data_tables:compositionend-orphan-repair] orphan line out of bounds, skipping', {
                            orphanLine: orphan.lineNum,
                            adjustedLine: adjustedLineNum,
                            totalLines,
                          });
                          continue;
                        }

                        if (!isLineSafeToModify(repCheck, adjustedLineNum, '[ep_data_tables:compositionend-orphan-repair]')) {
                          continue;
                        }

                        // Verify line content before deletion to prevent race condition
                        const lineEntry = repCheck.lines.atIndex(adjustedLineNum);
                        const currentLineText = lineEntry?.text || '';

                        // Skip empty lines (already cleaned)
                        if (!currentLineText.trim()) continue;

                        // Check if line has tbljson attribute matching target table
                        let lineAttr = null;
                        try {
                          lineAttr = docManager.getAttributeOnLine(adjustedLineNum, ATTR_TABLE_JSON);
                        } catch (_) {}

                        if (lineAttr) {
                          try {
                            const lineMeta = JSON.parse(lineAttr);
                            // Verify this is still for our target table
                            if (lineMeta.tblId !== targetTblId || lineMeta.row !== targetRow) {
                              console.debug('[ep_data_tables:compositionend-orphan-repair] line has different table, skipping', {
                                adjustedLine: adjustedLineNum,
                                lineTblId: lineMeta.tblId,
                                lineRow: lineMeta.row,
                                targetTblId,
                                targetRow,
                              });
                              continue;
                            }
                          } catch (_) {}
                        }

                        // Verify content similarity to catch major shifts and prevent data loss
                        const orphanOriginalText = orphan.text || '';
                        const textSimilarity = (currentLineText.length > 0 && orphanOriginalText.length > 0)
                          ? Math.min(currentLineText.length, orphanOriginalText.length) / Math.max(currentLineText.length, orphanOriginalText.length)
                          : 0;

                        if (orphanOriginalText.length > 0 && currentLineText.length > 0 && textSimilarity < 0.2) {
                          console.warn('[ep_data_tables:compositionend-orphan-repair] line content changed significantly, skipping deletion', {
                            adjustedLine: adjustedLineNum,
                            originalLen: orphanOriginalText.length,
                            currentLen: currentLineText.length,
                            similarity: textSimilarity.toFixed(2),
                          });
                          continue;
                        }

                      try {
                        if (docManager && typeof docManager.removeAttributeOnLine === 'function') {
                          docManager.removeAttributeOnLine(adjustedLineNum, ATTR_TABLE_JSON);
                        }
                        } catch (attrErr) {
                          console.debug('[ep_data_tables:compositionend-orphan-repair] removeAttribute failed', attrErr);
                        }

                      console.debug('[ep_data_tables:compositionend-orphan-repair] removing orphan (merged)', {
                        orphanLine: orphan.lineNum,
                        adjustedLine: adjustedLineNum,
                        contentVerified: true,
                      });
                      ace2.ace_performDocumentReplaceRange([adjustedLineNum, 0], [adjustedLineNum + 1, 0], '');
                        orphansRemoved++;
                        deletionOffset++; // Track that we deleted a line
                      } catch (orphanDeleteErr) {
                        console.error('[ep_data_tables:compositionend-orphan-repair] error deleting orphan line', {
                          orphanLine: orphan.lineNum,
                          error: orphanDeleteErr?.message || orphanDeleteErr,
                        });
                        // Check if this was a desync error
                        handleDomDesyncError(orphanDeleteErr, 'compositionend-orphan-repair');
                        // Don't break - try to continue with remaining orphans
                      }
                    }
                    }
                    try { ace2.ace_fastIncorp(5); } catch (incErr) {
                      console.debug('[ep_data_tables:compositionend-orphan-repair] fastIncorp error (non-fatal)', incErr?.message);
                    }
                    console.debug('[ep_data_tables:compositionend-orphan-repair] repair complete', {
                      primaryLine: primaryLine.lineNum, orphansRemoved,
                    });
                  }
                  } catch (innerErr) {
                    console.error('[ep_data_tables:compositionend-orphan-repair] inner callback error', {
                      error: innerErr?.message || innerErr,
                    });
                  }
                }, 'ep_data_tables:compositionend-orphan-repair', true);
              } catch (aceCallErr) {
                // This catches errors from ace_callWithAce itself (Etherpad internal state corruption)
                // When keyToNodeMap.get(...) is undefined, the internal state is corrupted
                // Don't rethrow - graceful degradation is better than crashing the editor
                console.warn('[ep_data_tables:compositionend-orphan-repair] ace_callWithAce failed - editor state may need refresh', {
                  error: aceCallErr?.message || aceCallErr,
                });
              }
            }, 50); // Small delay to let Etherpad process the cell edit first

            desktopComposition = { active: false, start: null, end: null, lineNum: null, cellIndex: -1, snapshot: null, snapshotMeta: null };
          }, 'tableDesktopCompositionEnd');
        } catch (compositionErr) {
          console.error(`${compLogPrefix} ERROR during desktop composition repair:`, compositionErr);
        }
      });
    });


    $inner.on('drop', (evt) => {
      const dropLogPrefix = '[ep_data_tables:dropHandler]';

      const targetEl = evt.target;
      if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('table.dataTable')) {
        evt.preventDefault();
        evt.stopPropagation();
        if (evt.originalEvent && evt.originalEvent.dataTransfer) {
          try { evt.originalEvent.dataTransfer.dropEffect = 'none'; } catch (_) {}
        }
        console.warn('[ep_data_tables] Drop prevented on table to protect structure.');
        return;
      }

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        return;
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let isTableLine = !!lineAttrString;

      if (!isTableLine) {
        const metadataFallback = getTableLineMetadata(lineNum, ed, docManager);
        isTableLine = !!metadataFallback;
      }

      if (isTableLine) {
      evt.preventDefault();
      evt.stopPropagation();
      console.warn('[ep_data_tables] Drop operation prevented to protect table structure. Please use copy/paste within table cells.');
      }
    });

    $inner.on('dragover', (evt) => {
      const dragLogPrefix = '[ep_data_tables:dragoverHandler]';

      const targetEl = evt.target;
      if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('table.dataTable')) {
        if (evt.originalEvent && evt.originalEvent.dataTransfer) {
          try { evt.originalEvent.dataTransfer.dropEffect = 'none'; } catch (_) {}
        }
        evt.preventDefault();
        evt.stopPropagation();
        return;
      }

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        return;
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let isTableLine = !!lineAttrString;

      if (!isTableLine) {
        isTableLine = !!getTableLineMetadata(lineNum, ed, docManager);
      }

      if (isTableLine) {
        evt.preventDefault();
      }
    });

    $inner.on('dragenter', (evt) => {
      const targetEl = evt.target;
      if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('table.dataTable')) {
        if (evt.originalEvent && evt.originalEvent.dataTransfer) {
          try { evt.originalEvent.dataTransfer.dropEffect = 'none'; } catch (_) {}
        }
        evt.preventDefault();
        evt.stopPropagation();
      }
    });

    $inner.on('paste', (evt) => {
      const pasteLogPrefix = '[ep_data_tables:pasteHandler]';

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        console.warn(`${pasteLogPrefix} Could not get rep or selStart.`);
        return;
      }
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      if (selStart[0] !== selEnd[0]) {
        evt.preventDefault();
        return;
      }

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (!lineAttrString) {
        const fallbackMeta = getTableLineMetadata(lineNum, ed, docManager);
        if (fallbackMeta) {
          tableMetadata = fallbackMeta;
          lineAttrString = JSON.stringify(fallbackMeta);
        }
      }

      if (!lineAttrString) {
        return;
      }

      try {
        if (!tableMetadata) {
          tableMetadata = JSON.parse(lineAttrString);
        }
        if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
          console.warn(`${pasteLogPrefix} Invalid table metadata for line ${lineNum}.`);
          return;
        }
      } catch(e) {
        console.error(`${pasteLogPrefix} ERROR parsing table metadata for line ${lineNum}:`, e);
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;

      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;

        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      /* allow "…cell content + delimiter" selections */
      if (targetCellIndex !== -1 &&
          selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      const clipboardData = evt.originalEvent.clipboardData || window.clipboardData;
      if (!clipboardData) {
        return;
      }

      const types = clipboardData.types || [];
      if (types.includes('text/html') && clipboardData.getData('text/html')) {
        return;
      }

      const pastedTextRaw = clipboardData.getData('text/plain');

      let pastedText = pastedTextRaw
        .replace(/(\r\n|\n|\r)/gm, " ")
        .replace(new RegExp(DELIMITER, 'g'), ' ')
        .replace(/\t/g, " ")
        .replace(/\s+/g, " ")
        .trim();


      if (typeof pastedText !== 'string' || pastedText.length === 0) {
        const types = clipboardData.types;
        if (types && types.includes('text/html')) {
        }
        return;
      }

      const currentCellText = cells[targetCellIndex] || '';
      const selectionLength = selEnd[1] - selStart[1];
      const newCellLength = currentCellText.length - selectionLength + pastedText.length;

      const MAX_CELL_LENGTH = 8000;
      if (newCellLength > MAX_CELL_LENGTH) {
        const truncatedPaste = pastedText.substring(0, MAX_CELL_LENGTH - (currentCellText.length - selectionLength));
        if (truncatedPaste.length === 0) {
          evt.preventDefault();
          return;
        }
        pastedText = truncatedPaste;
      }

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        ed.ace_callWithAce((aceInstance) => {
            const callAceLogPrefix = `${pasteLogPrefix}[ace_callWithAceOps]`;


            aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, pastedText);

            const repAfterReplace = aceInstance.ace_getRep();

            ed.ep_data_tables_applyMeta(
              lineNum,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );

            const newCaretCol = selStart[1] + pastedText.length;
            const newCaretPos = [lineNum, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);

            aceInstance.ace_fastIncorp(10);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
               const newRelativePos = newCaretCol - cellStartCol;
               editor.ep_data_tables_last_clicked = {
                  lineNum: lineNum,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
               };
            }

        }, 'tablePasteTextOperations', true);

      } catch (error) {
        console.error(`${pasteLogPrefix} CRITICAL ERROR during paste handling operation:`, error);
      }
    });


    const $iframeOuter = $('iframe[name="ace_outer"]');
    const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
    const innerDoc = $iframeInner.contents();
    const outerDoc = $iframeOuter.contents();


    $inner.on('mousedown', '.ep-data_tables-resize-handle', (evt) => {
      const resizeLogPrefix = '[ep_data_tables:resizeMousedown]';

      if (evt.button !== 0) {
        return;
      }

      const target = evt.target;
      const $target = $(target);
      const isImageRelated = $target.closest('.inline-image, .image-placeholder, .image-inner').length > 0;
      const isImageResizeHandle = $target.hasClass('image-resize-handle') || $target.closest('.image-resize-handle').length > 0;

      if (isImageRelated || isImageResizeHandle) {
        return;
      }

      evt.preventDefault();
      evt.stopPropagation();

      const handle = evt.target;
      const columnIndex = parseInt(handle.getAttribute('data-column'), 10);
      const table = handle.closest('table.dataTable');
      const lineNode = table.closest('div.ace-line');


      if (table && lineNode && !isNaN(columnIndex)) {
        const tblId = table.getAttribute('data-tblId');
        const rep = ed.ace_getRep();

        if (!rep || !rep.lines) {
          console.error(`${resizeLogPrefix} Cannot get editor representation`);
          return;
        }

        const lineNum = rep.lines.indexOfKey(lineNode.id);


        if (tblId && lineNum !== -1) {
          try {
            const lineAttrString = docManager.getAttributeOnLine(lineNum, 'tbljson');
            if (lineAttrString) {
              const metadata = JSON.parse(lineAttrString);
              if (metadata.tblId === tblId) {
                startColumnResize(table, columnIndex, evt.clientX, metadata, lineNum);

              } else {
              }
            } else {

              const rep = ed.ace_getRep();
              if (rep && rep.lines) {
                const lineEntry = rep.lines.atIndex(lineNum);
                if (lineEntry && lineEntry.lineNode) {
                  const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
                  if (tableInDOM) {
                    const domTblId = tableInDOM.getAttribute('data-tblId');
                    const domRow = tableInDOM.getAttribute('data-row');
                    if (domTblId === tblId && domRow !== null) {
                      const domCells = tableInDOM.querySelectorAll('td');
                      if (domCells.length > 0) {
                        const columnWidths = [];
                        domCells.forEach(cell => {
                          const style = cell.getAttribute('style') || '';
                          const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                          if (widthMatch) {
                            columnWidths.push(parseFloat(widthMatch[1]));
                          } else {
                            columnWidths.push(100 / domCells.length);
                          }
                        });

                        const reconstructedMetadata = {
                          tblId: domTblId,
                          row: parseInt(domRow, 10),
                          cols: domCells.length,
                          columnWidths: columnWidths
                        };

                        startColumnResize(table, columnIndex, evt.clientX, reconstructedMetadata, lineNum);

    } else {
                      }
                    } else {
                    }
                  } else {
                  }
                } else {
                }
              } else {
              }
            }
          } catch (e) {
            console.error(`${resizeLogPrefix} Error getting table metadata:`, e);
          }
        } else {
        }
      } else {
      }
    });

    const setupGlobalHandlers = () => {
      const mouseupLogPrefix = '[ep_data_tables:resizeMouseup]';
      const mousemoveLogPrefix = '[ep_data_tables:resizeMousemove]';

      const handleMousemove = (evt) => {
        if (isResizing) {
          evt.preventDefault();
          updateColumnResize(evt.clientX);
        }
      };

      const handleMouseup = (evt) => {

        if (isResizing) {
          evt.preventDefault();
          evt.stopPropagation();

          setTimeout(() => {
            finishColumnResize(ed, docManager);
          }, 50);
        } else {
        }
      };


      $(document).on('mousemove', handleMousemove);
      $(document).on('mouseup', handleMouseup);

      if (outerDoc.length > 0) {
        outerDoc.on('mousemove', handleMousemove);
        outerDoc.on('mouseup', handleMouseup);
      }

      if (innerDoc.length > 0) {
        innerDoc.on('mousemove', handleMousemove);
        innerDoc.on('mouseup', handleMouseup);
      }

      $inner.on('mousemove', handleMousemove);
      $inner.on('mouseup', handleMouseup);

      const failsafeMouseup = (evt) => {
        if (isResizing) {
          if (evt.type === 'mouseup' || evt.type === 'mousedown' || evt.type === 'click') {
            setTimeout(() => {
              if (isResizing) {
                finishColumnResize(ed, docManager);
              }
            }, 100);
          }
        }
      };

      document.addEventListener('mouseup', failsafeMouseup, true);
      document.addEventListener('mousedown', failsafeMouseup, true);
      document.addEventListener('click', failsafeMouseup, true);

      const preventTableDrag = (evt) => {
        const target = evt.target;
        const inTable = target && typeof target.closest === 'function' && target.closest('table.dataTable');
        if (inTable) {
          evt.preventDefault();
          evt.stopPropagation();
          if (evt.originalEvent && evt.originalEvent.dataTransfer) {
            try { evt.originalEvent.dataTransfer.effectAllowed = 'none'; } catch (_) {}
          }
          return false;
        }
      };

      $inner.on('dragstart', preventTableDrag);
      $inner.on('drag', preventTableDrag);
      $inner.on('dragend', preventTableDrag);

      if (innerDoc.length > 0) {
        innerDoc.on('dragstart', preventTableDrag);
        innerDoc.on('drag', preventTableDrag);
        innerDoc.on('dragend', preventTableDrag);
      }
      if (outerDoc.length > 0) {
        outerDoc.on('dragstart', preventTableDrag);
        outerDoc.on('drag', preventTableDrag);
        outerDoc.on('dragend', preventTableDrag);
      }
      $(document).on('dragstart', preventTableDrag);
      $(document).on('drag', preventTableDrag);
      $(document).on('dragend', preventTableDrag);
    };

    setupGlobalHandlers();

      } catch (e) {
        console.error(`${callWithAceLogPrefix} ERROR: Exception while attaching listeners:`, e);
      }
    }; // End of attachListeners function

    // Start the retry process to access iframes and attach all listeners
    tryGetIframeBody(0);

  }, 'tablePasteAndResizeListeners', true);

  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo, attributeString = null, documentAttributeManager = null) {
    const funcName = 'applyTableLineMetadataAttribute';

    let finalMetadata;

    if (attributeString) {
      try {
        const providedMetadata = JSON.parse(attributeString);
        if (providedMetadata.columnWidths && Array.isArray(providedMetadata.columnWidths) && providedMetadata.columnWidths.length === numCols) {
          finalMetadata = providedMetadata;
        } else {
          finalMetadata = providedMetadata;
           }
         } catch (e) {
        finalMetadata = null;
      }
    }

    if (!finalMetadata || !finalMetadata.columnWidths) {
      let columnWidths = null;

      try {
        const lineEntry = rep.lines.atIndex(lineNum);
        if (lineEntry && lineEntry.lineNode) {
          const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
          if (tableInDOM) {
            const domTblId = tableInDOM.getAttribute('data-tblId');
            if (domTblId === tblId) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    columnWidths.push(100 / numCols);
                  }
                });
              }
            }
          }
             }
           } catch (e) {
      }

      finalMetadata = finalMetadata || {
        tblId: tblId,
        row: rowIndex,
        cols: numCols
      };

      if (columnWidths && columnWidths.length === numCols) {
        finalMetadata.columnWidths = columnWidths;
      }
    }

    const finalAttributeString = JSON.stringify(finalMetadata);

    try {
       const lineEntry = rep.lines.atIndex(lineNum);
       if (!lineEntry) {
           return;
       }
       const lineLength = Math.max(1, lineEntry.text.length);

       const attributes = [[ATTR_TABLE_JSON, finalAttributeString]];
       const start = [lineNum, 0];
       const end = [lineNum, lineLength];

       editorInfo.ace_performDocumentApplyAttributesToRange(start, end, attributes);

    } catch(e) {
        console.error(`[ep_data_tables] ${logPrefix}:${funcName}: Error applying metadata attribute on line ${lineNum}:`, e);
    }
  }

  /** Insert a fresh rows×cols blank table at the caret */
  ed.ace_createTableViaAttributes = (rows = 2, cols = 2) => {
    const funcName = 'ace_createTableViaAttributes';
    rows = Math.max(1, rows); cols = Math.max(1, cols);

    const tblId   = rand();
    const initialCellContent = ' ';
    const lineTxt = Array.from({ length: cols }).fill(initialCellContent).join(DELIMITER);
    const block = Array.from({ length: rows }).fill(lineTxt).join('\n') + '\n';

    const currentRepInitial = ed.ace_getRep(); 
    if (!currentRepInitial || !currentRepInitial.selStart || !currentRepInitial.selEnd) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation or selection via ace_getRep(). Aborting.`);
        return;
    }
    const start = currentRepInitial.selStart;
    const end = currentRepInitial.selEnd;
    const initialStartLine = start[0];

    ed.ace_performDocumentReplaceRange(start, end, block);
    ed.ace_fastIncorp(20);

    const currentRep = ed.ace_getRep();
    if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get updated rep after text insertion. Cannot apply attributes reliably.`);
        return; 
    }

    for (let r = 0; r < rows; r++) {
      const lineNumToApply = initialStartLine + r;

      const lineEntry = currentRep.lines.atIndex(lineNumToApply);
      if (!lineEntry) {
        continue;
      }
      const lineText = lineEntry.text || '';
      const cells = lineText.split(DELIMITER);
      let offset = 0;

      for (let c = 0; c < cols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) {
          const cellStart = [lineNumToApply, offset];
          const cellEnd = [lineNumToApply, offset + cellContent.length];
          ed.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < cols - 1) {
          offset += DELIMITER.length;
        }
      }

      applyTableLineMetadataAttribute(lineNumToApply, tblId, r, cols, currentRep, ed, null, null); 
    }
    ed.ace_fastIncorp(20);

    const finalCaretLine = initialStartLine + rows;
    const finalCaretPos = [finalCaretLine, 0];
    try {
      ed.ace_performSelectionChange(finalCaretPos, finalCaretPos, false);
    } catch(e) {
       console.error(`[ep_data_tables] ${funcName}: Error setting caret position after table creation:`, e);
    }

  };

  ed.ace_doDatatableOptions = (action) => {
    const funcName = 'ace_doDatatableOptions';

    const editor = ed.ep_data_tables_editor;
    if (!editor) {
      console.error(`[ep_data_tables] ${funcName}: Could not get editor reference.`);
      return;
    }

    const lastClick = editor.ep_data_tables_last_clicked;
    if (!lastClick || !lastClick.tblId) {
      console.warn('[ep_data_tables] No table selected. Please click on a table cell first.');
      return;
    }


    try {
      const currentRep = ed.ace_getRep();
      if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation.`);
        return;
      }

      const docManager = ed.ep_data_tables_docManager;
      if (!docManager) {
        console.error(`[ep_data_tables] ${funcName}: Could not get document attribute manager from stored reference.`);
        return;
      }


      const tableLines = [];
      const totalLines = currentRep.lines.length();

      // Bidirectional search from clicked line (tables are contiguous)
      const startLine = lastClick.lineNum;
      const MAX_SEARCH_RADIUS = 200;
      let consecutiveNonTableUp = 0;
      let consecutiveNonTableDown = 0;
      const EARLY_TERMINATION_THRESHOLD = 5;

      const checkLineForTable = (lineIndex) => {
        if (lineIndex < 0 || lineIndex >= totalLines) return null;
        try {
          let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

          if (!lineAttrString) {
            const lineEntry = currentRep.lines.atIndex(lineIndex);
            if (lineEntry && lineEntry.lineNode) {
              const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
              if (tableInDOM) {
                const domTblId = tableInDOM.getAttribute('data-tblId');
                const domRow = tableInDOM.getAttribute('data-row');
                if (domTblId && domRow !== null) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length > 0) {
                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length
                    };
                    lineAttrString = JSON.stringify(reconstructedMetadata);
                  }
                }
              }
            }
          }

          if (lineAttrString) {
            const lineMetadata = JSON.parse(lineAttrString);
            if (lineMetadata.tblId === lastClick.tblId) {
              const lineEntry = currentRep.lines.atIndex(lineIndex);
              if (lineEntry) {
                return {
                  lineIndex,
                  row: lineMetadata.row,
                  cols: lineMetadata.cols,
                  lineText: lineEntry.text,
                  metadata: lineMetadata
                };
              }
            }
          }
        } catch (e) {
          // Ignore errors for individual lines
        }
        return null;
      };

      // Check the clicked line first
      const clickedLineResult = checkLineForTable(startLine);
      if (clickedLineResult) {
        tableLines.push(clickedLineResult);
      }

      // Search upward and downward simultaneously
      for (let offset = 1; offset <= MAX_SEARCH_RADIUS; offset++) {
        // Check if we should stop early (found all rows or hit termination threshold)
        const shouldStopUp = consecutiveNonTableUp >= EARLY_TERMINATION_THRESHOLD;
        const shouldStopDown = consecutiveNonTableDown >= EARLY_TERMINATION_THRESHOLD;

        if (shouldStopUp && shouldStopDown) {
          console.debug(`[ep_data_tables] ${funcName}: Early termination - hit ${EARLY_TERMINATION_THRESHOLD} consecutive non-table lines in both directions`);
          break;
        }

        // Search upward
        if (!shouldStopUp) {
          const upLine = startLine - offset;
          if (upLine >= 0) {
            const upResult = checkLineForTable(upLine);
            if (upResult) {
              tableLines.push(upResult);
              consecutiveNonTableUp = 0; // Reset counter on finding a table line
            } else {
              consecutiveNonTableUp++;
            }
          }
        }

        // Search downward
        if (!shouldStopDown) {
          const downLine = startLine + offset;
          if (downLine < totalLines) {
            const downResult = checkLineForTable(downLine);
            if (downResult) {
              tableLines.push(downResult);
              consecutiveNonTableDown = 0; // Reset counter on finding a table line
            } else {
              consecutiveNonTableDown++;
            }
          }
        }
      }

      if (tableLines.length === 0) {
        return;
      }

      tableLines.sort((a, b) => a.row - b.row);

      const numRows = tableLines.length;
      const numCols = tableLines[0].cols;

      let targetRowIndex = -1;

      targetRowIndex = tableLines.findIndex(line => line.lineIndex === lastClick.lineNum);

      if (targetRowIndex === -1) {
        const clickedLineEntry = currentRep.lines.atIndex(lastClick.lineNum);
        if (clickedLineEntry && clickedLineEntry.lineNode) {
          const clickedTable = clickedLineEntry.lineNode.querySelector('table.dataTable[data-tblId="' + lastClick.tblId + '"]');
          if (clickedTable) {
            const clickedRowAttr = clickedTable.getAttribute('data-row');
            if (clickedRowAttr !== null) {
              const clickedRowNum = parseInt(clickedRowAttr, 10);
              targetRowIndex = tableLines.findIndex(line => line.row === clickedRowNum);
            }
          }
        }
      }

      if (targetRowIndex === -1) {
        targetRowIndex = 0;
      }

      const targetColIndex = lastClick.cellIndex || 0;


      let newNumCols = numCols;
      let success = false;

      switch (action) {
        case 'addTblRowA':
          success = addTableRowAboveWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;

        case 'addTblRowB':
          success = addTableRowBelowWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;

        case 'addTblColL':
          newNumCols = numCols + 1;
          success = addTableColumnLeftWithText(tableLines, targetColIndex, ed, docManager);
          break;

        case 'addTblColR':
          newNumCols = numCols + 1;
          success = addTableColumnRightWithText(tableLines, targetColIndex, ed, docManager);
          break;

        case 'delTblRow':
          const rowConfirmMessage = `Are you sure you want to delete Row ${targetRowIndex + 1} and all content within?`;
          if (!confirm(rowConfirmMessage)) {
            return;
          }
          success = deleteTableRowWithText(tableLines, targetRowIndex, ed, docManager);
          break;

        case 'delTblCol':
          const colConfirmMessage = `Are you sure you want to delete Column ${targetColIndex + 1} and all content within?`;
          if (!confirm(colConfirmMessage)) {
            return;
          }
          newNumCols = numCols - 1;
          success = deleteTableColumnWithText(tableLines, targetColIndex, ed, docManager);
          break;

        default:
          return;
      }

      if (!success) {
        console.error(`[ep_data_tables] ${funcName}: Table operation failed for action: ${action}`);
        return;
      }


    } catch (error) {
      console.error(`[ep_data_tables] ${funcName}: Error during table operation:`, error);
    }
  };

  function addTableRowAboveWithText(tableLines, targetRowIndex, numCols, tblId, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      const newLineText = Array.from({ length: numCols }).fill(' ').join(DELIMITER);
      const insertLineIndex = targetLine.lineIndex;

      editorInfo.ace_performDocumentReplaceRange([insertLineIndex, 0], [insertLineIndex, 0], newLineText + '\n');

      const rep = editorInfo.ace_getRep();
      const cells = newLineText.split(DELIMITER);
      let offset = 0;
      for (let c = 0; c < numCols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) {
          const cellStart = [insertLineIndex, offset];
          const cellEnd = [insertLineIndex, offset + cellContent.length];
          editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < numCols - 1) {
          offset += DELIMITER.length;
        }
      }

      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        try {
          const rep = editorInfo.ace_getRep();
          const lineEntry = rep.lines.atIndex(targetLine.lineIndex + 1);
          if (lineEntry && lineEntry.lineNode) {
            const tableInDOM = lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${tblId}"]`);
            if (tableInDOM) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    columnWidths.push(100 / numCols);
                  }
                });
              }
            }
          }
        } catch (e) {
          console.error('[ep_data_tables] addTableRowAbove: Error extracting column widths from DOM:', e);
        }
      }

      for (let i = targetRowIndex; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1;
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };

        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      const newMetadata = { tblId, row: targetLine.metadata.row, cols: numCols, columnWidths };
      applyTableLineMetadataAttribute(insertLineIndex, tblId, targetLine.metadata.row, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding row above with text:', e);
      return false;
    }
  }

  function addTableRowBelowWithText(tableLines, targetRowIndex, numCols, tblId, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      const newLineText = Array.from({ length: numCols }).fill(' ').join(DELIMITER);
      const insertLineIndex = targetLine.lineIndex + 1;

      editorInfo.ace_performDocumentReplaceRange([insertLineIndex, 0], [insertLineIndex, 0], newLineText + '\n');

      const rep = editorInfo.ace_getRep();
      const cells = newLineText.split(DELIMITER);
      let offset = 0;
      for (let c = 0; c < numCols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) {
          const cellStart = [insertLineIndex, offset];
          const cellEnd = [insertLineIndex, offset + cellContent.length];
          editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < numCols - 1) {
          offset += DELIMITER.length;
        }
      }

      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        try {
          const rep = editorInfo.ace_getRep();
          const lineEntry = rep.lines.atIndex(targetLine.lineIndex);
          if (lineEntry && lineEntry.lineNode) {
            const tableInDOM = lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${tblId}"]`);
            if (tableInDOM) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    columnWidths.push(100 / numCols);
                  }
                });
              }
            }
          }
        } catch (e) {
          console.error('[ep_data_tables] addTableRowBelow: Error extracting column widths from DOM:', e);
        }
      }

      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1;
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };

        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      const newMetadata = { tblId, row: targetLine.metadata.row + 1, cols: numCols, columnWidths };
      applyTableLineMetadataAttribute(insertLineIndex, tblId, targetLine.metadata.row + 1, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding row below with text:', e);
      return false;
    }
  }

  function addTableColumnLeftWithText(tableLines, targetColIndex, editorInfo, docManager) {
    const funcName = 'addTableColumnLeftWithText';
    __epDT_columnOperationInProgress = true;
    try {
      // First, update metadata for ALL rows BEFORE modifying content
      // This prevents postWriteCanonicalize from seeing stale column counts
      const newColCount = tableLines[0].cols + 1;
      const equalWidth = 100 / newColCount;
      const normalizedWidths = Array(newColCount).fill(equalWidth);
      
      for (const tableLine of tableLines) {
        const newMetadata = { ...tableLine.metadata, cols: newColCount, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, newColCount, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Now insert the new column content for each row
      for (const tableLine of tableLines) {
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        const lineText = lineEntry?.text || tableLine.lineText;
        const cells = lineText.split(DELIMITER);

        let insertPos = 0;
        for (let i = 0; i < targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0) + DELIMITER.length;
        }

        const textToInsert = ' ' + DELIMITER;
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];

        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);

        // Get fresh rep after text change
        const repAfter = editorInfo.ace_getRep();
        const lineEntryAfter = repAfter.lines.atIndex(tableLine.lineIndex);
        if (lineEntryAfter) {
          const newLineText = lineEntryAfter.text || '';
          const newCells = newLineText.split(DELIMITER);
          
          // First, REMOVE existing tblCell-* attributes from the ENTIRE line
          // This prevents overlapping cell markers
          const lineLen = newLineText.length;
          if (lineLen > 0) {
            for (let oldCell = 0; oldCell < tableLine.cols; oldCell++) {
              editorInfo.ace_performDocumentApplyAttributesToRange(
                [tableLine.lineIndex, 0], 
                [tableLine.lineIndex, lineLen], 
                [[ATTR_CELL, '']]  // Empty value removes the attribute
              );
            }
          }
          
          // Now apply fresh tblCell-* attributes to each cell
          let offset = 0;
          for (let c = 0; c < newColCount; c++) {
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) {
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
              console.debug(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL}=${c} to Line ${tableLine.lineIndex} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding column left with text:', e);
      return false;
    } finally {
      __epDT_columnOperationInProgress = false;
    }
  }

  function addTableColumnRightWithText(tableLines, targetColIndex, editorInfo, docManager) {
    const funcName = 'addTableColumnRightWithText';
    __epDT_columnOperationInProgress = true;
    try {
      // First, update metadata for ALL rows BEFORE modifying content
      // This prevents postWriteCanonicalize from seeing stale column counts
      const newColCount = tableLines[0].cols + 1;
      const equalWidth = 100 / newColCount;
      const normalizedWidths = Array(newColCount).fill(equalWidth);
      
      for (const tableLine of tableLines) {
        const newMetadata = { ...tableLine.metadata, cols: newColCount, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, newColCount, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Now insert the new column content for each row
      for (const tableLine of tableLines) {
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        const lineText = lineEntry?.text || tableLine.lineText;
        const cells = lineText.split(DELIMITER);

        let insertPos = 0;
        for (let i = 0; i <= targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0);
          if (i < targetColIndex) insertPos += DELIMITER.length;
        }

        const textToInsert = DELIMITER + ' ';
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];

        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);

        // Get fresh rep after text change
        const repAfter = editorInfo.ace_getRep();
        const lineEntryAfter = repAfter.lines.atIndex(tableLine.lineIndex);
        if (lineEntryAfter) {
          const newLineText = lineEntryAfter.text || '';
          const newCells = newLineText.split(DELIMITER);
          
          // First, REMOVE existing tblCell-* attributes from the ENTIRE line
          // This prevents overlapping cell markers
          const lineLen = newLineText.length;
          if (lineLen > 0) {
            for (let oldCell = 0; oldCell < tableLine.cols; oldCell++) {
              editorInfo.ace_performDocumentApplyAttributesToRange(
                [tableLine.lineIndex, 0], 
                [tableLine.lineIndex, lineLen], 
                [[ATTR_CELL, '']]  // Empty value removes the attribute
              );
            }
          }
          
          // Now apply fresh tblCell-* attributes to each cell
          let offset = 0;
          for (let c = 0; c < newColCount; c++) {
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) {
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
              console.debug(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL}=${c} to Line ${tableLine.lineIndex} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding column right with text:', e);
      return false;
    } finally {
      __epDT_columnOperationInProgress = false;
    }
  }

  function deleteTableRowWithText(tableLines, targetRowIndex, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];

      if (targetRowIndex === 0) {
        const insertStart = [targetLine.lineIndex, 0];
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertStart, '\n');

        const deleteStart = [targetLine.lineIndex + 1, 0];
        const deleteEnd = [targetLine.lineIndex + 2, 0];
        editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      } else {
      const deleteStart = [targetLine.lineIndex, 0];
      const deleteEnd = [targetLine.lineIndex + 1, 0];
      editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      }

      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        try {
          const rep = editorInfo.ace_getRep();
          for (const tableLine of tableLines) {
            if (tableLine.lineIndex !== targetLine.lineIndex) {
              const lineEntry = rep.lines.atIndex(tableLine.lineIndex >= targetLine.lineIndex ? tableLine.lineIndex - 1 : tableLine.lineIndex);
              if (lineEntry && lineEntry.lineNode) {
                const tableInDOM = lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${targetLine.metadata.tblId}"]`);
                if (tableInDOM) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length === targetLine.metadata.cols) {
                    columnWidths = [];
                    domCells.forEach(cell => {
                      const style = cell.getAttribute('style') || '';
                      const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                      if (widthMatch) {
                        columnWidths.push(parseFloat(widthMatch[1]));
                      } else {
                        columnWidths.push(100 / targetLine.metadata.cols);
                      }
                    });
                    break;
                  }
                }
              }
            }
          }
        } catch (e) {
          console.error('[ep_data_tables] deleteTableRow: Error extracting column widths from DOM:', e);
        }
      }

      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex - 1;
        const newRowIndex = tableLines[i].metadata.row - 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };

        applyTableLineMetadataAttribute(lineToUpdate, tableLines[i].metadata.tblId, newRowIndex, tableLines[i].cols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error deleting row with text:', e);
      return false;
    }
  }
  function deleteTableColumnWithText(tableLines, targetColIndex, editorInfo, docManager) {
    const funcName = 'deleteTableColumnWithText';
    __epDT_columnOperationInProgress = true;
    try {
      const newColCount = tableLines[0].cols - 1;
      if (newColCount <= 0) {
        console.warn('[ep_data_tables] Cannot delete column - would result in 0 columns');
        return false;
      }
      
      // First, update metadata for ALL rows BEFORE modifying content
      // This prevents postWriteCanonicalize from seeing stale column counts
      const equalWidth = 100 / newColCount;
      const normalizedWidths = Array(newColCount).fill(equalWidth);
      
      for (const tableLine of tableLines) {
        const newMetadata = { ...tableLine.metadata, cols: newColCount, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, newColCount, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Now delete the column content for each row
      for (const tableLine of tableLines) {
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        const lineText = lineEntry?.text || tableLine.lineText;
        const cells = lineText.split(DELIMITER);

        if (targetColIndex >= cells.length) {
          console.debug(`[ep_data_tables] ${funcName}: Target column ${targetColIndex} doesn't exist in line with ${cells.length} columns`);
          continue;
        }

        let deleteStart = 0;
        let deleteEnd = 0;

        for (let i = 0; i < targetColIndex; i++) {
          deleteStart += (cells[i]?.length ?? 0) + DELIMITER.length;
        }

        deleteEnd = deleteStart + (cells[targetColIndex]?.length ?? 0);

        if (targetColIndex === 0 && cells.length > 1) {
          deleteEnd += DELIMITER.length;
        } else if (targetColIndex > 0) {
          deleteStart -= DELIMITER.length;
        }

        console.debug(`[ep_data_tables] ${funcName}: Deleting column ${targetColIndex} from line ${tableLine.lineIndex}: chars ${deleteStart}-${deleteEnd}`);

        const rangeStart = [tableLine.lineIndex, deleteStart];
        const rangeEnd = [tableLine.lineIndex, deleteEnd];

        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, '');

        // Get fresh rep after text change and re-apply tblCell attributes
        const repAfter = editorInfo.ace_getRep();
        const lineEntryAfter = repAfter.lines.atIndex(tableLine.lineIndex);
        if (lineEntryAfter) {
          const newLineText = lineEntryAfter.text || '';
          const newCells = newLineText.split(DELIMITER);
          
          // First, REMOVE existing tblCell-* attributes from the ENTIRE line
          const lineLen = newLineText.length;
          if (lineLen > 0) {
            editorInfo.ace_performDocumentApplyAttributesToRange(
              [tableLine.lineIndex, 0], 
              [tableLine.lineIndex, lineLen], 
              [[ATTR_CELL, '']]  // Empty value removes the attribute
            );
          }
          
          // Now apply fresh tblCell-* attributes to each remaining cell
          let offset = 0;
          for (let c = 0; c < newColCount; c++) {
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) {
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
              console.debug(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL}=${c} to Line ${tableLine.lineIndex} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error deleting column with text:', e);
      return false;
    } finally {
      __epDT_columnOperationInProgress = false;
    }
  }


};

exports.aceEditorCSS                = () => { 
  return ['ep_data_tables/static/css/datatables-editor.css', 'ep_data_tables/static/css/caret.css'];
};

exports.aceRegisterBlockElements = () => ['table'];

// Ensure contentcollector treats <table> as a block when collecting
exports.ccRegisterBlockElements = () => ['table'];

// Emit canonical row text during DOM->changeset collection to avoid IME collapse
// Also SUPPRESS orphan tbljson spans that are NOT inside a table to prevent line fragmentation
exports.collectContentLineText = (hookName, context) => {
  try {
    const {cc, state, node} = context || {};
    if (!node || node.nodeType !== (typeof Node !== 'undefined' ? Node.TEXT_NODE : 3)) return;
    const parentEl = node.parentElement || (node.parentNode && node.parentNode.nodeType === 1 ? node.parentNode : null);
    if (!parentEl || typeof parentEl.closest !== 'function') return;
    
    const table = parentEl.closest('table.dataTable[data-tblId], table.dataTable[data-tblid]');
    
    // Check for orphan table spans (tbljson-* or tblCell-*) not inside a table
    if (!table) {
      const hasTableRelatedClass = (el) => {
        if (!el) return false;
        if (el.classList && el.classList.length > 0) {
          for (const cls of el.classList) {
            if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) return true;
          }
        }
        if (el.className) {
          const classStr = typeof el.className === 'string' ? el.className : String(el.className);
          if (classStr.includes('tbljson-') || classStr.includes('tblCell-')) return true;
        }
        return false;
      };
      
      let checkEl = parentEl;
      let foundOrphanTableSpan = false;
      let detectedClass = null;
      const docBody = node.ownerDocument?.body || document.body;
      while (checkEl && checkEl !== docBody) {
        if (hasTableRelatedClass(checkEl)) {
          foundOrphanTableSpan = true;
          try { for (const cls of (checkEl.classList || [])) if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) { detectedClass = cls; break; } } catch (_) {}
          break;
        }
        checkEl = checkEl.parentElement;
      }
      
      if (foundOrphanTableSpan) {
        // Check if this line has a rendered table - if not, allow fresh paste content through
        try {
          const lineDiv = parentEl.closest('div.ace-line');
          if (lineDiv && !lineDiv.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]')) {
            console.debug('[ep_data_tables:collector] allowing fresh tbljson content (no table yet)', { lineId: lineDiv.id || null, detectedClass });
            return;
          }
        } catch (_) {}
        
        // During composition, skip orphan suppression to avoid discarding active input
        if (isInCompositionCooldown()) {
          console.debug('[ep_data_tables:collector] skipping orphan suppression (composition cooldown)');
          return;
        }

        context.text = '';
        return;
      }
      
      // Suppress stray content on table lines that isn't inside the table
      try {
        const lineDiv = parentEl.closest('div.ace-line');
        if (lineDiv) {
          const tableInLine = lineDiv.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
          if (tableInLine) {
            // Skip suppression during composition to avoid discarding active input
            if (isInCompositionCooldown()) {
              console.debug('[ep_data_tables:collector] skipping non-table content suppression (composition cooldown)');
              return;
            }

            context.text = '';
            return;
          }
        }
      } catch (_) {}
      
      // Not a table-related element, let normal collection proceed
      return;
    }

    // Only emit once per ace line during collection
    if (state && state._epDT_emittedCanonical) { 
      context.text = ''; 
      return; 
    }
    if (state) state._epDT_emittedCanonical = true;

    const tr = table.querySelector('tbody > tr');
    if (!tr || !tr.children || tr.children.length === 0) return;

    // Sanitize cell text, preserving zero-width chars for images (they use U+200B as placeholder)
    const sanitize = (s, hasImage) => {
      let x = (s || '').replace(new RegExp(DELIMITER, 'g'), ' ');
      if (!hasImage) {
        x = x.replace(/[\u200B\u200C\u200D\uFEFF]/g, '');
      }
      if (!x) x = ' ';
      return x;
    };
    
    // Detect image cells by class or ZWS placeholder chars
    const cellHasImage = (td) => {
      if (td.querySelector('[class*="image:"], [class*="inline-image"], span.image-placeholder')) return true;
      const text = td.innerText || td.textContent || '';
      return /[\u200B\u200C\u200D\uFEFF]/.test(text);
    };
    
    const cells = Array.from(tr.children).map((td) => sanitize(td.innerText || '', cellHasImage(td)));
    const canonical = cells.join(DELIMITER);
    
    // Emit canonical text for first text node only; subsequent nodes get suppressed.
    // Styling is captured before emit and restored after to preserve both structure and styling.
    
    // Get table metadata for cache key fallback (when lineId unavailable)
    const tableTblId = table.getAttribute('data-tblId') || table.getAttribute('data-tblid') || '';
    const tableRowAttr = table.getAttribute('data-row');
    const cacheKeyFallback = tableTblId && tableRowAttr != null ? `${tableTblId}:${tableRowAttr}` : null;
    
    // Resolve the ace-line div for styling capture
    let lineDiv = parentEl.closest && parentEl.closest('div.ace-line');
    if (!lineDiv && table) {
      lineDiv = table.closest('div.ace-line');
    }
    
    // Capture styling before emitting canonical
    const cacheKey = (lineDiv && lineDiv.id) ? lineDiv.id : cacheKeyFallback;
    
    if (lineDiv || cacheKeyFallback) {
      try {
        const editorInfo = EP_DT_EDITOR_INFO;
        if (editorInfo && typeof editorInfo.ep_data_tables_extractStyling === 'function') {
          const extractTarget = lineDiv || table.closest('div.ace-line') || table.parentElement;
          const capturedStyling = editorInfo.ep_data_tables_extractStyling(extractTarget);
          if (capturedStyling && capturedStyling.length > 0 && cacheKey) {
            cacheStylingForLine(cacheKey, capturedStyling, cells);
            
            // Schedule styling restoration after collection to batch text + styling into one changeset
            const keyForRestore = cacheKey;
            const tblIdForRestore = tableTblId;
            const rowForRestore = tableRowAttr != null ? parseInt(tableRowAttr, 10) : 0;
            
            // setTimeout(0) runs after collection but before changesetTracker sends to server
            setTimeout(() => {
              try {
                if (editorInfo && typeof editorInfo.ep_data_tables_restoreCachedStyling === 'function') {
                  editorInfo.ep_data_tables_restoreCachedStyling(keyForRestore, tblIdForRestore, rowForRestore);
                }
              } catch (restoreErr) {
                console.debug('[ep_data_tables:collector] styling restore error (non-fatal)', restoreErr?.message);
              }
            }, 0);
          }
        }
      } catch (captureErr) {
        console.debug('[ep_data_tables:collector] styling capture error (non-fatal)', captureErr?.message);
      }
    }
    
    // Emit canonical text to preserve table structure
    context.text = canonical;

    // Apply/refresh tbljson line attribute
    const tblId = table.getAttribute('data-tblId') || table.getAttribute('data-tblid') || '';
    const rowStr = table.getAttribute('data-row');
    const row = rowStr != null ? parseInt(rowStr, 10) : 0;
    const meta = { tblId, row, cols: cells.length };
    try { cc && cc.doAttrib && state && cc.doAttrib(state, `${ATTR_TABLE_JSON}::${JSON.stringify(meta)}`); } catch (_) {}
  } catch (_) {
    // swallow collection-time errors to avoid breaking contentcollector
  }
};

const startColumnResize = (table, columnIndex, startX, metadata, lineNum) => {
  const funcName = 'startColumnResize';

  isResizing = true;
  resizeStartX = startX;
  resizeCurrentX = startX;
  resizeTargetTable = table;
  resizeTargetColumn = columnIndex;
  resizeTableMetadata = metadata;
  resizeLineNum = lineNum;

  const numCols = metadata.cols;
  resizeOriginalWidths = metadata.columnWidths ? [...metadata.columnWidths] : Array(numCols).fill(100 / numCols);


  createResizeOverlay(table, columnIndex);

  document.body.style.userSelect = 'none';
  document.body.style.webkitUserSelect = 'none';
  document.body.style.mozUserSelect = 'none';
  document.body.style.msUserSelect = 'none';
};

const createResizeOverlay = (table, columnIndex) => {
  if (resizeOverlay) {
    resizeOverlay.remove();
  }

  const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
  if ($innerIframe.length === 0) {
    console.error('[ep_data_tables] createResizeOverlay: Could not find inner iframe');
    return;
  }

  const innerDocBody = $innerIframe.contents().find('body')[0];
  const padOuter = $('iframe[name="ace_outer"]').contents().find('body');

  if (!innerDocBody || padOuter.length === 0) {
    console.error('[ep_data_tables] createResizeOverlay: Could not find required container elements');
          return;
      }

  const tblId = table.getAttribute('data-tblId');
  if (!tblId) {
    console.error('[ep_data_tables] createResizeOverlay: No tblId found on table');
    return;
  }

  const allTableRows = innerDocBody.querySelectorAll(`table.dataTable[data-tblId="${tblId}"]`);
  if (allTableRows.length === 0) {
    console.error('[ep_data_tables] createResizeOverlay: No table rows found for tblId:', tblId);
    return;
  }

  let minTop = Infinity;
  let maxBottom = -Infinity;
  let tableLeft = 0;
  let tableWidth = 0;

  Array.from(allTableRows).forEach((tableRow, index) => {
    const rect = tableRow.getBoundingClientRect();
    minTop = Math.min(minTop, rect.top);
    maxBottom = Math.max(maxBottom, rect.bottom);

    if (index === 0) {
      tableLeft = rect.left;
      tableWidth = rect.width;
    }
  });

  const totalTableHeight = maxBottom - minTop;


  let innerBodyRect, innerIframeRect, outerBodyRect;
  let scrollTopInner, scrollLeftInner, scrollTopOuter, scrollLeftOuter;

  try {
    innerBodyRect = innerDocBody.getBoundingClientRect();
    innerIframeRect = $innerIframe[0].getBoundingClientRect();
    outerBodyRect = padOuter[0].getBoundingClientRect();
    scrollTopInner = innerDocBody.scrollTop;
    scrollLeftInner = innerDocBody.scrollLeft;
    scrollTopOuter = padOuter.scrollTop();
    scrollLeftOuter = padOuter.scrollLeft();
  } catch (e) {
    console.error('[ep_data_tables] createResizeOverlay: Error getting container rects/scrolls:', e);
    return;
  }

  const tableTopRelInner = minTop - innerBodyRect.top + scrollTopInner;
  const tableLeftRelInner = tableLeft - innerBodyRect.left + scrollLeftInner;

  const innerFrameTopRelOuter = innerIframeRect.top - outerBodyRect.top + scrollTopOuter;
  const innerFrameLeftRelOuter = innerIframeRect.left - outerBodyRect.left + scrollLeftOuter;

  const overlayTopOuter = innerFrameTopRelOuter + tableTopRelInner;
  const overlayLeftOuter = innerFrameLeftRelOuter + tableLeftRelInner;

  const outerPadding = window.getComputedStyle(padOuter[0]);
  const outerPaddingTop = parseFloat(outerPadding.paddingTop) || 0;
  const outerPaddingLeft = parseFloat(outerPadding.paddingLeft) || 0;

  const MANUAL_OFFSET_TOP = 6;
  const MANUAL_OFFSET_LEFT = 39;

  const finalOverlayTop = overlayTopOuter + outerPaddingTop + MANUAL_OFFSET_TOP;
  const finalOverlayLeft = overlayLeftOuter + outerPaddingLeft + MANUAL_OFFSET_LEFT;

  const tds = table.querySelectorAll('td');
  const tds_array = Array.from(tds);
  let linePosition = 0;

  if (columnIndex < tds_array.length) {
    const currentTd = tds_array[columnIndex];
    const currentTdRect = currentTd.getBoundingClientRect();
    const currentRelativeLeft = currentTdRect.left - tableLeft;
    const currentWidth = currentTdRect.width;
    linePosition = currentRelativeLeft + currentWidth;
  }

  resizeOverlay = document.createElement('div');
  resizeOverlay.className = 'ep-data_tables-resize-overlay';
  resizeOverlay.style.cssText = `
    position: absolute;
    left: ${finalOverlayLeft}px;
    top: ${finalOverlayTop}px;
    width: ${tableWidth}px;
    height: ${totalTableHeight}px;
    pointer-events: none;
    z-index: 1000;
    background: transparent;
    box-sizing: border-box;
  `;

  const resizeLine = document.createElement('div');
  resizeLine.className = 'resize-line';
  resizeLine.style.cssText = `
    position: absolute;
    left: ${linePosition}px;
    top: 0;
    width: 2px;
    height: 100%;
    background: #1a73e8;
    z-index: 1001;
  `;
  resizeOverlay.appendChild(resizeLine);

  padOuter.append(resizeOverlay);

};

const updateColumnResize = (currentX) => {
  if (!isResizing || !resizeTargetTable || !resizeOverlay) return;

  resizeCurrentX = currentX;
  const deltaX = currentX - resizeStartX;

  const tblId = resizeTargetTable.getAttribute('data-tblId');
  if (!tblId) return;

  const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
  const innerDocBody = $innerIframe.contents().find('body')[0];
  const firstTableRow = innerDocBody.querySelector(`table.dataTable[data-tblId="${tblId}"]`);

  if (!firstTableRow) return;

  const tableRect = firstTableRow.getBoundingClientRect();
  const deltaPercent = (deltaX / tableRect.width) * 100;

  const newWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;

  if (nextColumn < newWidths.length) {
    const transfer = Math.min(deltaPercent, newWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(newWidths[currentColumn] - 5));

    newWidths[currentColumn] += actualTransfer;
    newWidths[nextColumn] -= actualTransfer;

    const resizeLine = resizeOverlay.querySelector('.resize-line');
    if (resizeLine) {
      const newColumnWidth = (newWidths[currentColumn] / 100) * tableRect.width;

      const tds = firstTableRow.querySelectorAll('td');
      const tds_array = Array.from(tds);

      if (currentColumn < tds_array.length) {
        const currentTd = tds_array[currentColumn];
        const currentTdRect = currentTd.getBoundingClientRect();
        const currentRelativeLeft = currentTdRect.left - tableRect.left;

        const newLinePosition = currentRelativeLeft + newColumnWidth;
        resizeLine.style.left = newLinePosition + 'px';
      }
    }
  }
};

const finishColumnResize = (editorInfo, docManager) => {
  if (!isResizing || !resizeTargetTable) {
                      return;
                  }

  const funcName = 'finishColumnResize';

  const tableRect = resizeTargetTable.getBoundingClientRect();
  const deltaX = resizeCurrentX - resizeStartX;
  const deltaPercent = (deltaX / tableRect.width) * 100;


  const finalWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;

  if (nextColumn < finalWidths.length) {
    const transfer = Math.min(deltaPercent, finalWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(finalWidths[currentColumn] - 5));

    finalWidths[currentColumn] += actualTransfer;
    finalWidths[nextColumn] -= actualTransfer;

  }

  const totalWidth = finalWidths.reduce((sum, width) => sum + width, 0);
  if (totalWidth > 0) {
    finalWidths.forEach((width, index) => {
      finalWidths[index] = (width / totalWidth) * 100;
    });
  }


  if (resizeOverlay) {
    resizeOverlay.remove();
    resizeOverlay = null;
  }

  document.body.style.userSelect = '';
  document.body.style.webkitUserSelect = '';
  document.body.style.mozUserSelect = '';
  document.body.style.msUserSelect = '';

  isResizing = false;

  editorInfo.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = `${funcName}[ace_callWithAce]`;

    try {
      const rep = ace.ace_getRep();
      if (!rep || !rep.lines) {
        console.error(`${callWithAceLogPrefix}: Invalid rep`);
        return;
      }

      const tableLines = [];
      const totalLines = rep.lines.length();

      // Bidirectional search from resize line
      const startLine = resizeLineNum >= 0 ? resizeLineNum : 0;
      const MAX_SEARCH_RADIUS = 200;
      let consecutiveNonTableUp = 0;
      let consecutiveNonTableDown = 0;
      const EARLY_TERMINATION_THRESHOLD = 5;

      const checkLineForTableResize = (lineIndex) => {
        if (lineIndex < 0 || lineIndex >= totalLines) return null;
        try {
          let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

          if (lineAttrString) {
            const lineMetadata = JSON.parse(lineAttrString);
            if (lineMetadata.tblId === resizeTableMetadata.tblId) {
              return { lineIndex, metadata: lineMetadata };
            }
          } else {
            const lineEntry = rep.lines.atIndex(lineIndex);
            if (lineEntry && lineEntry.lineNode) {
              const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
              if (tableInDOM) {
                const domTblId = tableInDOM.getAttribute('data-tblId');
                const domRow = tableInDOM.getAttribute('data-row');
                if (domTblId === resizeTableMetadata.tblId && domRow !== null) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length > 0) {
                    const columnWidths = [];
                    domCells.forEach(cell => {
                      const style = cell.getAttribute('style') || '';
                      const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                      if (widthMatch) {
                        columnWidths.push(parseFloat(widthMatch[1]));
                      } else {
                        columnWidths.push(100 / domCells.length);
                      }
                    });

                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length,
                      columnWidths: columnWidths
                    };
                    return { lineIndex, metadata: reconstructedMetadata };
                  }
                }
              }
            }
          }
        } catch (e) {
          // Ignore errors for individual lines
        }
        return null;
      };

      // Check the starting line first
      const startResult = checkLineForTableResize(startLine);
      if (startResult) {
        tableLines.push(startResult);
      }

      // Bidirectional search
      for (let offset = 1; offset <= MAX_SEARCH_RADIUS; offset++) {
        const shouldStopUp = consecutiveNonTableUp >= EARLY_TERMINATION_THRESHOLD;
        const shouldStopDown = consecutiveNonTableDown >= EARLY_TERMINATION_THRESHOLD;

        if (shouldStopUp && shouldStopDown) break;

        if (!shouldStopUp) {
          const upLine = startLine - offset;
          if (upLine >= 0) {
            const upResult = checkLineForTableResize(upLine);
            if (upResult) {
              tableLines.push(upResult);
              consecutiveNonTableUp = 0;
            } else {
              consecutiveNonTableUp++;
            }
          }
        }

        if (!shouldStopDown) {
          const downLine = startLine + offset;
          if (downLine < totalLines) {
            const downResult = checkLineForTableResize(downLine);
            if (downResult) {
              tableLines.push(downResult);
              consecutiveNonTableDown = 0;
            } else {
              consecutiveNonTableDown++;
            }
          }
        }
      }


      for (const tableLine of tableLines) {
        const updatedMetadata = { ...tableLine.metadata, columnWidths: finalWidths };
        const updatedMetadataString = JSON.stringify(updatedMetadata);

        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (!lineEntry) {
          console.error(`${callWithAceLogPrefix}: Could not get line entry for line ${tableLine.lineIndex}`);
          continue;
        }

        const lineLength = Math.max(1, lineEntry.text.length);
        const rangeStart = [tableLine.lineIndex, 0];
        const rangeEnd = [tableLine.lineIndex, lineLength];


        ace.ace_performDocumentApplyAttributesToRange(rangeStart, rangeEnd, [
          [ATTR_TABLE_JSON, updatedMetadataString]
        ]);
      }


    } catch (error) {
      console.error(`${callWithAceLogPrefix}: Error applying updated metadata:`, error);
    }
  }, 'applyTableResizeToAllRows', true);


  resizeStartX = 0;
  resizeCurrentX = 0;
  resizeTargetTable = null;
  resizeTargetColumn = -1;
  resizeOriginalWidths = [];
  resizeTableMetadata = null;
  resizeLineNum = -1;

};

exports.aceUndoRedo = (hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceUndoRedo]';

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd) {
    return;
  }

  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];

  let hasTableLines = false;
  let tableLines = [];

  for (let line = startLine; line <= endLine; line++) {
    const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
    if (lineAttrString) {
      hasTableLines = true;
      tableLines.push(line);
    }
  }

  if (!hasTableLines) {
    return;
  }


  try {
    for (const line of tableLines) {
      const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
      if (!lineAttrString) continue;

      const tableMetadata = JSON.parse(lineAttrString);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
        const lineText = ctx.rep.lines.atIndex(line)?.text || '';
        const cells = lineText.split(DELIMITER);

        if (cells.length > 1) {
          const newMetadata = {
            cols: cells.length,
            rows: 1,
            cells: cells.map((_, i) => ({ col: i, row: 0 }))
          };

          ctx.documentAttributeManager.setAttributeOnLine(line, ATTR_TABLE_JSON, JSON.stringify(newMetadata));
        } else {
          ctx.documentAttributeManager.removeAttributeOnLine(line, ATTR_TABLE_JSON);
        }
      }
    }
  } catch (e) {
    console.error(`${logPrefix} Error during undo/redo validation:`, e);
  }
};





