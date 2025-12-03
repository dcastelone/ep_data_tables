const ATTR_TABLE_JSON = 'tbljson';
// Global guard: Chrome/macOS sometimes loads the same plugin script twice due to a
// subtle preload/import timing quirk, causing duplicate aceInitialized hooks that
// explode into large duplicate-row changesets.  Bail early on re-entry.
if (typeof window !== 'undefined') {
  if (window.__epDataTablesLoaded) {
    console.debug('[ep_data_tables] Duplicate client_hooks.js load suppressed');
    // eslint-disable-next-line no-useless-return
    return; // Abort evaluation of the rest of the module
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
// Track original table line from compositionstart for corruption recovery
// This helps postWriteCanonicalize know where the table SHOULD be, not just where the attribute ended up
let __epDT_compositionOriginalLine = { tblId: null, lineNum: null, timestamp: 0 };
// Flag to prevent postWriteCanonicalize from interfering during column operations
let __epDT_columnOperationInProgress = false;
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
         // log(`${funcName}: Found metadata via attribute for line ${lineNum}`);
          return metadata;
        }
      } catch (e) {
       // log(`${funcName}: Invalid JSON in tbljson attribute on line ${lineNum}:`, e.message);
      }
    }

   // log(`${funcName}: No valid attribute on line ${lineNum}, checking DOM.`);
    const rep = editorInfo.ace_getRep();

    const lineEntry = rep.lines.atIndex(lineNum);
    const lineNode = lineEntry?.lineNode;

    if (!lineNode) {
     // log(`${funcName}: Could not find line node in rep for line ${lineNum}`);
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
           // log(`${funcName}: Reconstructed metadata from DOM for line ${lineNum}:`, metadata);
            return metadata;
          } catch (e) {
            console.error(`${funcName}: Failed to decode/parse tbljson class on line ${lineNum}:`, e);
            return null;
          }
        }
      }
    }

   // log(`${funcName}: Could not find table metadata for line ${lineNum} in DOM.`);
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
 // log(`${funcName}: START - Current: Line=${currentLineNum}, Cell=${currentCellIndex}, Shift=${shiftKey}`);

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

   // log(`${funcName}: Target coordinates - Row=${targetRow}, Col=${targetCol}`);

  const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);
    if (targetLineNum === -1) {
     // log(`${funcName}: Could not find line for target row ${targetRow}`);
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
 // log(`${funcName}: START - Current: Line=${currentLineNum}, Cell=${currentCellIndex}`);

  try {
  const targetRow = tableMetadata.row + 1;
    const targetCol = currentCellIndex;

   // log(`${funcName}: Target coordinates - Row=${targetRow}, Col=${targetCol}`);

  const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);

  if (targetLineNum !== -1) {
     // log(`${funcName}: Found line for target row ${targetRow}, navigating.`);
      return navigateToCell(targetLineNum, targetCol, editorInfo, docManager);
    } else {
     // log(`${funcName}: Could not find next row. Creating new line after table.`);
  const rep = editorInfo.ace_getRep();
      const lineTextLength = rep.lines.atIndex(currentLineNum).text.length;
      const endOfLinePos = [currentLineNum, lineTextLength];

      editorInfo.ace_performSelectionChange(endOfLinePos, endOfLinePos, false);
  editorInfo.ace_performDocumentReplaceRange(endOfLinePos, endOfLinePos, '\n');

      editorInfo.ace_updateBrowserSelectionFromRep();
  editorInfo.ace_focus();

      const editor = editorInfo.editor;
      if (editor) editor.ep_data_tables_last_clicked = null;
     // log(`${funcName}: Cleared last click info as we have exited the table.`);

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
 // log(`${funcName}: Searching for tblId=${tblId}, row=${targetRow}`);

  try {
  const rep = editorInfo.ace_getRep();
    if (!rep || !rep.lines) {
     // log(`${funcName}: Could not get rep or rep.lines`);
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
               // log(`${funcName}: Found target via DOM: line ${lineIndex}`);
                return lineIndex;
              }
            }
          }
        }

        if (lineAttrString) {
          const lineMetadata = JSON.parse(lineAttrString);
          if (lineMetadata.tblId === tblId && lineMetadata.row === targetRow) {
           // log(`${funcName}: Found target via attribute: line ${lineIndex}`);
            return lineIndex;
          }
        }
      } catch (e) {
        continue;
      }
    }

   // log(`${funcName}: Target row not found`);
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
 // log(`${funcName}: START - Target: Line=${targetLineNum}, Cell=${targetCellIndex}`);
  let targetPos;

  try {
    const rep = editorInfo.ace_getRep();
    if (!rep || !rep.lines) {
     // log(`${funcName}: Could not get rep or rep.lines`);
      return false;
    }

    const lineEntry = rep.lines.atIndex(targetLineNum);
    if (!lineEntry) {
     // log(`${funcName}: Could not get line entry for line ${targetLineNum}`);
      return false;
    }

    const lineText = lineEntry.text || '';
    const cells = lineText.split(DELIMITER);

    if (targetCellIndex >= cells.length) {
     // log(`${funcName}: Target cell ${targetCellIndex} doesn't exist (only ${cells.length} cells)`);
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
       // log(`${funcName}: Pre-emptively updated stored click info:`, editor.ep_data_tables_last_clicked);
      } else {
       // log(`${funcName}: Could not get table metadata for target line ${targetLineNum}, cannot update click info.`);
      }
    } catch (e) {
     // log(`${funcName}: Could not update stored click info before navigation:`, e.message);
    }

    try {
    editorInfo.ace_performSelectionChange(targetPos, targetPos, false);
     // log(`${funcName}: Updated internal selection to [${targetPos}]`);

    editorInfo.ace_updateBrowserSelectionFromRep();
     // log(`${funcName}: Called updateBrowserSelectionFromRep to sync visual caret.`);

    editorInfo.ace_focus();
     // log(`${funcName}: Editor focused.`);

    } catch(e) {
      console.error(`[ep_data_tables] ${funcName}: Error during direct navigation update:`, e);
      return false;
    }

  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during cell navigation:`, e);
    return false;
  }

 // log(`${funcName}: Navigation considered successful.`);
  return true;
}

exports.collectContentPre = (hook, ctx) => {
  const funcName = 'collectContentPre';
  const node = ctx.domNode;
  const state = ctx.state;
  const cc = ctx.cc;

 // log(`${funcName}: *** ENTRY POINT *** Hook: ${hook}, Node: ${node?.tagName}.${node?.className}`);

  if (node?.classList?.contains('ace-line')) {
    const tableNode = node.querySelector('table.dataTable[data-tblId]');
    if (tableNode) {
     // log(`${funcName}: Found ace-line with rendered table. Attempting reconstruction from DOM.`);

    const docManager = cc.documentAttributeManager;
    const rep = cc.rep;
    const lineNum = rep?.lines?.indexOfKey(node.id);

      if (typeof lineNum === 'number' && lineNum >= 0 && docManager) {
       // log(`${funcName}: Processing line ${lineNum} (NodeID: ${node.id}) for DOM reconstruction.`);
    try {
      const existingAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
         // log(`${funcName}: Line ${lineNum} existing ${ATTR_TABLE_JSON} attribute: '${existingAttrString}'`);

          if (existingAttrString) {
      const existingMetadata = JSON.parse(existingAttrString);
            if (existingMetadata && typeof existingMetadata.tblId !== 'undefined' &&
                typeof existingMetadata.row !== 'undefined' && typeof existingMetadata.cols === 'number') {
             // log(`${funcName}: Line ${lineNum} existing metadata is valid:`, existingMetadata);

      const trNode = tableNode.querySelector('tbody > tr');
              if (trNode) {
               // log(`${funcName}: Line ${lineNum} found <tr> node for cell content extraction.`);
                let cellPlainTexts = Array.from(trNode.children).map((td) => extractCellPlainText(td));

      if (cellPlainTexts.length !== existingMetadata.cols) {
                   // log(`${funcName}: WARNING Line ${lineNum}: Reconstructed cell count (${cellPlainTexts.length}) does not match metadata cols (${existingMetadata.cols}). Padding/truncating.`);
        while (cellPlainTexts.length < existingMetadata.cols) cellPlainTexts.push(' ');
                    if (cellPlainTexts.length > existingMetadata.cols) cellPlainTexts.length = existingMetadata.cols;
                }

                const canonicalLineText = cellPlainTexts.join(DELIMITER);
                state.line = canonicalLineText;
               // log(`${funcName}: Line ${lineNum} successfully reconstructed ctx.state.line: "${canonicalLineText.substring(0, 200)}..."`);

                state.lineAttributes = state.lineAttributes || [];
                state.lineAttributes = state.lineAttributes.filter(attr => attr[0] !== ATTR_TABLE_JSON);
      state.lineAttributes.push([ATTR_TABLE_JSON, existingAttrString]);
               // log(`${funcName}: Line ${lineNum} ensured ${ATTR_TABLE_JSON} attribute is in state.lineAttributes.`);

               // log(`${funcName}: Line ${lineNum} reconstruction complete. Returning undefined to prevent default DOM collection.`);
                return undefined;
              } else {
               // log(`${funcName}: ERROR Line ${lineNum}: Could not find tbody > tr in rendered table for reconstruction.`);
              }
            } else {
             // log(`${funcName}: ERROR Line ${lineNum}: Invalid or incomplete existing metadata from line attribute:`, existingMetadata);
            }
          } else {
           // log(`${funcName}: WARNING Line ${lineNum}: No existing ${ATTR_TABLE_JSON} attribute found for reconstruction, despite table DOM presence. Table may be malformed or attribute lost.`);
            const domTblId = tableNode.getAttribute('data-tblId');
            const domRow = tableNode.getAttribute('data-row');
            const trNode = tableNode.querySelector('tbody > tr');
            if (domTblId && domRow !== null && trNode && trNode.children.length > 0) {
               // log(`${funcName}: Line ${lineNum} FALLBACK: Attempting reconstruction using table DOM attributes as ${ATTR_TABLE_JSON} was missing.`);
                const domCols = trNode.children.length;
                const tempMetadata = {tblId: domTblId, row: parseInt(domRow, 10), cols: domCols};
                const tempAttrString = JSON.stringify(tempMetadata);
               // log(`${funcName}: Line ${lineNum} FALLBACK: Constructed temporary metadata: ${tempAttrString}`);

                let cellPlainTexts = Array.from(trNode.children).map((td) => extractCellPlainText(td));

                if (cellPlainTexts.length !== domCols) {
                    // log(`${funcName}: WARNING Line ${lineNum} (Fallback): Reconstructed cell count (${cellPlainTexts.length}) does not match DOM cols (${domCols}).`);
                     while(cellPlainTexts.length < domCols) cellPlainTexts.push(' ');
                     if(cellPlainTexts.length > domCols) cellPlainTexts.length = domCols;
                }

                const canonicalLineText = cellPlainTexts.join(DELIMITER);
                state.line = canonicalLineText;
                state.lineAttributes = state.lineAttributes || [];
                state.lineAttributes = state.lineAttributes.filter(attr => attr[0] !== ATTR_TABLE_JSON);
                state.lineAttributes.push([ATTR_TABLE_JSON, tempAttrString]);
               // log(`${funcName}: Line ${lineNum} FALLBACK: Successfully reconstructed line using DOM attributes. Returning undefined.`);
                return undefined;
            } else {
                // log(`${funcName}: Line ${lineNum} FALLBACK: Could not reconstruct from DOM attributes due to missing info.`);
            }
          }
    } catch (e) {
          console.error(`[ep_data_tables] ${funcName}: Line ${lineNum} error during DOM reconstruction:`, e);
         // log(`${funcName}: Line ${lineNum} Exception details:`, { message: e.message, stack: e.stack });
        }
      } else {
       // log(`${funcName}: Could not get valid line number (${lineNum}), rep, or docManager for DOM reconstruction of ace-line.`);
      }
    } else {
     // log(`${funcName}: Node is ace-line but no rendered table.dataTable[data-tblId] found. Allowing normal processing for: ${node?.className}`);
    }
  } else {
   // log(`${funcName}: Node is not an ace-line (or node is null). Node: ${node?.tagName}.${node?.className}. Allowing normal processing.`);
  }


  const classes = ctx.cls ? ctx.cls.split(' ') : [];
  let appliedAttribFromClass = false;
  if (classes.length > 0) {
   // log(`${funcName}: Secondary path - Checking classes on node ${node?.tagName}.${node?.className}: [${classes.join(', ')}]`);
  for (const cls of classes) {
    if (cls.startsWith('tbljson-')) {
       // log(`${funcName}: Secondary path - Found tbljson class: ${cls} on node ${node?.tagName}.${node?.className}`);
        const encodedMetadata = cls.substring(8);
      try {
        const decodedMetadata = dec(encodedMetadata);
        if (decodedMetadata) {
          cc.doAttrib(state, `${ATTR_TABLE_JSON}::${decodedMetadata}`);
            appliedAttribFromClass = true;
           // log(`${funcName}: Secondary path - Applied attribute to OP via cc.doAttrib for class ${cls.substring(0, 20)}... on ${node?.tagName}`);
          } else {
           // log(`${funcName}: Secondary path - ERROR - Decoded metadata is null or empty for class ${cls}`);
        }
      } catch (e) {
          console.error(`[ep_data_tables] ${funcName}: Secondary path - Error processing tbljson class ${cls} on ${node?.tagName}:`, e);
      }
        break; 
      }
    }
    if (!appliedAttribFromClass && classes.some(c => c.startsWith('tbljson-'))) {
       // log(`${funcName}: Secondary path - Found tbljson- class but failed to apply attribute.`);
    } else if (!classes.some(c => c.startsWith('tbljson-'))) {
       // log(`${funcName}: Secondary path - No tbljson- class found on this node.`);
    }
  } else {
    // log(`${funcName}: Secondary path - Node ${node?.tagName}.${node?.className} has no ctx.cls or classes array is empty.`);
  }

 // log(`${funcName}: *** EXIT POINT *** For Node: ${node?.tagName}.${node?.className}. Applied from class: ${appliedAttribFromClass}`);
};

exports.aceAttribsToClasses = (hook, ctx) => {
  const funcName = 'aceAttribsToClasses';
 // log(`>>>> ${funcName}: Called with key: ${ctx.key}`); // log entry
  if (ctx.key === ATTR_TABLE_JSON) {
   // log(`${funcName}: Processing ATTR_TABLE_JSON.`);
    const rawJsonValue = ctx.value;
   // log(`${funcName}: Received raw attribute value (ctx.value):`, rawJsonValue);

    let parsedMetadataForLog = '[JSON Parse Error]';
    try {
        parsedMetadataForLog = JSON.parse(rawJsonValue);
       // log(`${funcName}: Value parsed for logging:`, parsedMetadataForLog);
    } catch(e) {
       // log(`${funcName}: Error parsing raw JSON value for logging:`, e);
    }

    const className = `tbljson-${enc(rawJsonValue)}`;
   // log(`${funcName}: Generated class name by encoding raw JSON: ${className}`);
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
 // log(`${funcName}: START`, { metadata, innerHTMLSegments });

  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined') {
    console.error(`[ep_data_tables] ${funcName}: Invalid or missing metadata. Aborting.`);
   // log(`${funcName}: END - Error`);
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
 // log(`${funcName}: Joined all cellsHtml:`, cellsHtml);

  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';
 // log(`${funcName}: First row class applied: '${firstRowClass}'`);

  const tableHtml = `<table class="dataTable${firstRowClass}" writingsuggestions="false" autocorrect="off" autocapitalize="off" spellcheck="false" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse: collapse; table-layout: fixed;" draggable="false"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
 // log(`${funcName}: Generated final table HTML:`, tableHtml);
 // log(`${funcName}: END - Success`);
  return tableHtml;
}

exports.acePostWriteDomLineHTML = function (hook_name, args, cb) {
  const funcName = 'acePostWriteDomLineHTML';
  const node = args?.node;
  const nodeId = node?.id;
  const lineNumFromArgs = args?.lineNumber;
  let resolvedRep = null;
  let resolvedLineNum = -1;
  try {
    if (node && nodeId) {
      // Resolve the line number via DOM sibling index to avoid touching internal rep key maps here.
      resolvedLineNum = _getLineNumberOfElement(node);
      console.debug('[ep_data_tables:acePostWriteDomLineHTML] resolve-dom-index', { nodeId, resolvedLineNum });
      // Avoid calling ace_callWithAce during domline render to prevent re-entrancy issues.
      console.debug('[ep_data_tables:acePostWriteDomLineHTML] resolve-rep-skip', { reason: 'avoid-ace-call-during-postwrite' });
    } else {
      console.debug('[ep_data_tables:acePostWriteDomLineHTML] resolve-skip', {
        hasEditorInfo: !!EP_DT_EDITOR_INFO, hasNode: !!node, nodeId,
      });
    }
  } catch (e) {
    console.error('[ep_data_tables:acePostWriteDomLineHTML] resolve-exception', e);
  }
  const lineNum = (typeof lineNumFromArgs === 'number') ? lineNumFromArgs : resolvedLineNum;
  const logPrefix = '[ep_data_tables:acePostWriteDomLineHTML]';

 // log(`${logPrefix} ----- START ----- NodeID: ${nodeId} LineNum: ${lineNum}`);
  if (!node || !nodeId) {
     // log(`${logPrefix} ERROR - Received invalid node or node without ID. Aborting.`);
      console.error(`[ep_data_tables] ${funcName}: Received invalid node or node without ID.`);
    return cb();
  }

 // log(`${logPrefix} NodeID#${nodeId}: COMPLETE DOM STRUCTURE DEBUG:`);
 // log(`${logPrefix} NodeID#${nodeId}: Node tagName: ${node.tagName}`);
 // log(`${logPrefix} NodeID#${nodeId}: Node className: ${node.className}`);
 // log(`${logPrefix} NodeID#${nodeId}: Node innerHTML length: ${node.innerHTML?.length || 0}`);
 // log(`${logPrefix} NodeID#${nodeId}: Node innerHTML (first 500 chars): "${(node.innerHTML || '').substring(0, 500)}"`);
 // log(`${logPrefix} NodeID#${nodeId}: Node children count: ${node.children?.length || 0}`);

  if (node.children) {
    for (let i = 0; i < Math.min(node.children.length, 10); i++) {
      const child = node.children[i];
     // log(`${logPrefix} NodeID#${nodeId}: Child[${i}] tagName: ${child.tagName}, className: "${child.className}", innerHTML length: ${child.innerHTML?.length || 0}`);
      if (child.className && child.className.includes('tbljson-')) {
       // log(`${logPrefix} NodeID#${nodeId}: *** FOUND TBLJSON CLASS ON CHILD[${i}] ***`);
      }
    }
  }

  let rowMetadata = null;
  let encodedJsonString = null;

 // log(`${logPrefix} NodeID#${nodeId}: Searching for tbljson-* class...`);

  function findTbljsonClass(element, depth = 0, path = '') {
    const indent = '  '.repeat(depth);
   // log(`${logPrefix} NodeID#${nodeId}: ${indent}Searching element: ${element.tagName || 'unknown'}, path: ${path}`);

    if (element.classList) {
     // log(`${logPrefix} NodeID#${nodeId}: ${indent}Element has ${element.classList.length} classes: [${Array.from(element.classList).join(', ')}]`);
      for (const cls of element.classList) {
          if (cls.startsWith('tbljson-')) {
         // log(`${logPrefix} NodeID#${nodeId}: ${indent}*** FOUND TBLJSON CLASS: ${cls.substring(8)} at depth ${depth}, path: ${path} ***`);
          return cls.substring(8);
        }
      }
    } else {
     // log(`${logPrefix} NodeID#${nodeId}: ${indent}Element has no classList`);
    }

    if (element.children) {
     // log(`${logPrefix} NodeID#${nodeId}: ${indent}Element has ${element.children.length} children`);
      for (let i = 0; i < element.children.length; i++) {
        const child = element.children[i];
        const childPath = `${path}>${child.tagName}[${i}]`;
        const found = findTbljsonClass(child, depth + 1, childPath);
        if (found) {
         // log(`${logPrefix} NodeID#${nodeId}: ${indent}Returning found result from child: ${found}`);
          return found;
      }
    }
    } else {
     // log(`${logPrefix} NodeID#${nodeId}: ${indent}Element has no children`);
    }

   // log(`${logPrefix} NodeID#${nodeId}: ${indent}No tbljson class found in this element or its children`);
    return null;
  }

 // log(`${logPrefix} NodeID#${nodeId}: Starting recursive search for tbljson class...`);
  encodedJsonString = findTbljsonClass(node, 0, 'ROOT');

  if (encodedJsonString) {
   // log(`${logPrefix} NodeID#${nodeId}: *** SUCCESS: Found encoded tbljson class: ${encodedJsonString} ***`);
  } else {
   // log(`${logPrefix} NodeID#${nodeId}: *** NO TBLJSON CLASS FOUND ***`);
  } 

  if (!encodedJsonString) {
     // log(`${logPrefix} NodeID#${nodeId}: No tbljson-* class found. Assuming not a table line. END.`);

     // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Node tag: ${node.tagName}, Node classes:`, Array.from(node.classList || []));
     // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Node innerHTML (first 200 chars): "${(node.innerHTML || '').substring(0, 200)}"`);

      if (node.children && node.children.length > 0) {
        for (let i = 0; i < Math.min(node.children.length, 5); i++) {
          const child = node.children[i];
         // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Child ${i} tag: ${child.tagName}, classes:`, Array.from(child.classList || []));
        }
      }

      const existingTable = node.querySelector('table.dataTable[data-tblId]');
      if (existingTable) {
        const existingTblId = existingTable.getAttribute('data-tblId');
        const existingRow = existingTable.getAttribute('data-row');
       // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Found orphaned table! TblId: ${existingTblId}, Row: ${existingRow}`);

        if (existingTblId && existingRow !== null) {
         // log(`${logPrefix} NodeID#${nodeId}: POTENTIAL ISSUE - Table exists but no tbljson class. This may be a post-resize issue.`);

          const tableCells = existingTable.querySelectorAll('td');
         // log(`${logPrefix} NodeID#${nodeId}: Table has ${tableCells.length} cells`);

          if (lineNum !== undefined && args?.documentAttributeManager) {
            try {
              const currentLineAttr = args.documentAttributeManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
             // log(`${logPrefix} NodeID#${nodeId}: Current line ${lineNum} tbljson attribute: ${currentLineAttr || 'NULL'}`);
            } catch (e) {
             // log(`${logPrefix} NodeID#${nodeId}: Error getting line attribute:`, e);
            }
          }
        }
      }

      return cb(); 
  }

  const existingTable = node.querySelector('table.dataTable[data-tblId]');
  if (existingTable) {
     // log(`${logPrefix} NodeID#${nodeId}: Table already exists in DOM. Skipping innerHTML replacement.`);
      return cb();
  }

 // log(`${logPrefix} NodeID#${nodeId}: Decoding and parsing metadata...`);
  try {
    const decoded = dec(encodedJsonString);
     // log(`${logPrefix} NodeID#${nodeId}: Decoded string: ${decoded}`);
      if (!decoded) throw new Error('Decoded string is null or empty.');
    rowMetadata = JSON.parse(decoded);
     // log(`${logPrefix} NodeID#${nodeId}: Parsed rowMetadata:`, rowMetadata);

      if (!rowMetadata || typeof rowMetadata.tblId === 'undefined' || typeof rowMetadata.row === 'undefined' || typeof rowMetadata.cols !== 'number') {
          throw new Error('Invalid or incomplete metadata (missing tblId, row, or cols).');
      }
     // log(`${logPrefix} NodeID#${nodeId}: Metadata validated successfully.`);

  } catch(e) { 
     // log(`${logPrefix} NodeID#${nodeId}: FATAL ERROR - Failed to decode/parse/validate tbljson metadata. Rendering cannot proceed.`, e);
      console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Failed to decode/parse/validate tbljson.`, encodedJsonString, e);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Invalid table metadata attribute found.</div>';
     // log(`${logPrefix} NodeID#${nodeId}: Rendered error message in node. END.`);
    return cb();
  }

  const delimitedTextFromLine = node.innerHTML;
 // log(`${logPrefix} NodeID#${nodeId}: Using node.innerHTML for delimited text to preserve styling.`);
 // log(`${logPrefix} NodeID#${nodeId}: Raw innerHTML length: ${delimitedTextFromLine?.length || 0}`);
 // log(`${logPrefix} NodeID#${nodeId}: Raw innerHTML (first 1000 chars): "${(delimitedTextFromLine || '').substring(0, 1000)}"`);

  const delimiterCount = (delimitedTextFromLine || '').split(DELIMITER).length - 1;
 // log(`${logPrefix} NodeID#${nodeId}: Delimiter '${DELIMITER}' count in innerHTML: ${delimiterCount}`);
 // log(`${logPrefix} NodeID#${nodeId}: Expected delimiters for ${rowMetadata.cols} columns: ${rowMetadata.cols - 1}`);

  let pos = -1;
  const delimiterPositions = [];
  while ((pos = delimitedTextFromLine.indexOf(DELIMITER, pos + 1)) !== -1) {
    delimiterPositions.push(pos);
   // log(`${logPrefix} NodeID#${nodeId}: Delimiter found at position ${pos}, context: "${delimitedTextFromLine.substring(Math.max(0, pos - 20), pos + 21)}"`);
  }
 // log(`${logPrefix} NodeID#${nodeId}: All delimiter positions: [${delimiterPositions.join(', ')}]`);

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

 // log(`${logPrefix} NodeID#${nodeId}: *** SEGMENT ANALYSIS ***`);
 // log(`${logPrefix} NodeID#${nodeId}: Split resulted in ${htmlSegments.length} segments`);
  for (let i = 0; i < htmlSegments.length; i++) {
    const segment = htmlSegments[i] || '';
   // log(`${logPrefix} NodeID#${nodeId}: Segment[${i}] length: ${segment.length}`);
   // log(`${logPrefix} NodeID#${nodeId}: Segment[${i}] content (first 200 chars): "${segment.substring(0, 200)}"`);
    if (segment.length > 200) {
     // log(`${logPrefix} NodeID#${nodeId}: Segment[${i}] content (chars 200-400): "${segment.substring(200, 400)}"`);
    }
    if (segment.length > 400) {
     // log(`${logPrefix} NodeID#${nodeId}: Segment[${i}] content (chars 400-600): "${segment.substring(400, 600)}"`);
    }
    if (segment.includes('image:') || segment.includes('image-placeholder') || segment.includes('currently-selected')) {
     // log(`${logPrefix} NodeID#${nodeId}: *** SEGMENT[${i}] CONTAINS IMAGE CONTENT ***`);
    }
    try {
      const tblCellMatches = segment.match(/\btblCell-(\d+)\b/g) || [];
      const tbljsonMatches = segment.match(/\btbljson-[A-Za-z0-9_-]+\b/g) || [];
      const uniqueCells = Array.from(new Set(tblCellMatches));
      if (uniqueCells.length > 1) {
        console.warn('[ep_data_tables][diag] segment contains multiple tblCell-* markers', { segIndex: i, uniqueCells });
      }
    } catch (_) {}
  }

 // log(`${logPrefix} NodeID#${nodeId}: Parsed HTML segments (${htmlSegments.length}):`, htmlSegments.map(s => (s || '').substring(0,50) + (s && s.length > 50 ? '...' : '')));

  let finalHtmlSegments = htmlSegments;

  // Flag to track if we should skip the early return after mismatch detection
  let skipMismatchReturn = false;

  if (htmlSegments.length !== rowMetadata.cols) {
     // log(`${logPrefix} NodeID#${nodeId}: *** MISMATCH DETECTED *** - Attempting reconstruction.`);
    console.warn('[ep_data_tables][diag] Segment/column mismatch', { nodeId, lineNum, segs: htmlSegments.length, cols: rowMetadata.cols, tblId: rowMetadata.tblId, row: rowMetadata.row });
    const hasImageSelected = delimitedTextFromLine.includes('currently-selected');
    const hasImageContent = delimitedTextFromLine.includes('image:');
    if (hasImageSelected) {
      // note only
    }
    
    // Skip scheduling mismatch repair during column operations - the column operation will handle it
    // Also skip the early return so the table can still be rendered with current segments
    if (__epDT_columnOperationInProgress) {
      console.debug('[ep_data_tables][diag] mismatch-skip-scheduling', { nodeId, reason: 'column-operation-in-progress' });
      skipMismatchReturn = true;
      // Don't schedule repair, continue to table building below
    }
    
    // Defer canonicalization to after the render tick; do not mutate DOM here.
    // But skip if column operation is in progress
    else if (nodeId && !__epDT_postWriteScheduled.has(nodeId)) {
      __epDT_postWriteScheduled.add(nodeId);
      const fallbackLineNum = lineNum; // DOM-derived index as fallback
      const capturedTblId = rowMetadata.tblId;
      const capturedRow = rowMetadata.row;
      const expectedCols = rowMetadata.cols;
      setTimeout(() => {
        try {
          if (!EP_DT_EDITOR_INFO) return;
          const ed = EP_DT_EDITOR_INFO;
          const docManager = ed.ep_data_tables_docManager || null;
          
          // Skip if a column operation is in progress - let the column operation handle its own cleanup
          if (__epDT_columnOperationInProgress) {
            console.debug('[ep_data_tables:postWriteCanonicalize] skipped - column operation in progress');
            return;
          }
          
          // CRITICAL: Pre-check editor state before calling ace_callWithAce
          // Prevents crash when Etherpad's internal keyToNodeMap is corrupted
          try {
            const preCheckRep = ed.ace_getRep && ed.ace_getRep();
            if (!preCheckRep || !preCheckRep.lines) {
              console.debug('[ep_data_tables:postWriteCanonicalize] skipped - rep unavailable in pre-check');
              return;
            }
          } catch (preCheckErr) {
            console.debug('[ep_data_tables:postWriteCanonicalize] skipped - pre-check failed', preCheckErr?.message);
            return;
          }
          
          try {
          ed.ace_callWithAce((ace) => {
            try {
              const rep = ace.ace_getRep();
              if (!rep || !rep.lines) return;
              
              // Check if we have original line info from compositionstart
              const origLineInfo = __epDT_compositionOriginalLine;
              const hasOriginalLineInfo = origLineInfo && 
                origLineInfo.tblId === capturedTblId && 
                typeof origLineInfo.lineNum === 'number' &&
                (Date.now() - origLineInfo.timestamp) < 5000; // Only use if recent (within 5 seconds)
              
              console.debug('[ep_data_tables:postWriteCanonicalize] start', {
                nodeId, fallbackLineNum, capturedTblId, capturedRow, expectedCols,
                hasDocManager: !!docManager,
                originalLineInfo: hasOriginalLineInfo ? origLineInfo.lineNum : null,
              });
              if (!docManager || typeof docManager.getAttributeOnLine !== 'function') {
                console.debug('[ep_data_tables:postWriteCanonicalize] abort: no documentAttributeManager available');
                return;
              }
              // Robust line resolution: primary by attribute scan (tblId/row), then key lookup, then fallback.
              let ln = -1;
              let attrScanLine = -1; // Track where attr scan finds the table (may differ from original)
              try {
                const totalScan = rep.lines.length();
                for (let i = 0; i < totalScan; i++) {
                  let sAttr = null;
                  try { sAttr = docManager.getAttributeOnLine(i, ATTR_TABLE_JSON); } catch (_) {}
                  if (!sAttr) continue;
                  try {
                    const m = JSON.parse(sAttr);
                    if (m && m.tblId === capturedTblId && m.row === capturedRow) {
                      attrScanLine = i;
                      ln = i;
                      console.debug('[ep_data_tables:postWriteCanonicalize] found-by-attr-scan', { ln, tblId: m.tblId, row: m.row, cols: m.cols });
                      break;
                    }
                  } catch (_) { /* ignore parse errors */ }
                }
              } catch (scanErr) {
                console.error('[ep_data_tables:postWriteCanonicalize] line scan error', scanErr);
              }
              
              // CRITICAL: If we have original line info and it differs from attr-scan,
              // check if the original line still has a valid table DOM
              if (hasOriginalLineInfo && attrScanLine >= 0 && attrScanLine !== origLineInfo.lineNum) {
                console.debug('[ep_data_tables:postWriteCanonicalize] line shift detected', {
                  attrScanLine,
                  originalLine: origLineInfo.lineNum,
                  shift: attrScanLine - origLineInfo.lineNum,
                });
                
                // Check if the original line still has a table DOM
                const origEntry = rep.lines.atIndex(origLineInfo.lineNum);
                if (origEntry?.lineNode) {
                  const origTableEl = origEntry.lineNode.querySelector(
                    `table.dataTable[data-tblId="${capturedTblId}"], table.dataTable[data-tblid="${capturedTblId}"]`
                  );
                  if (origTableEl) {
                    console.debug('[ep_data_tables:postWriteCanonicalize] original line still has table DOM, preferring it', {
                      originalLine: origLineInfo.lineNum,
                      attrLine: attrScanLine,
                    });
                    ln = origLineInfo.lineNum;
                  }
                }
              }
              
              if (ln < 0) {
                try {
                  if (rep.lines && typeof rep.lines.indexOfKey === 'function') {
                    ln = rep.lines.indexOfKey(nodeId);
                    console.debug('[ep_data_tables:postWriteCanonicalize] found-by-indexOfKey', { ln, nodeId });
                  }
                } catch (eIdx) {
                  console.error('[ep_data_tables:postWriteCanonicalize] indexOfKey error; using fallback line', eIdx);
                  ln = (typeof fallbackLineNum === 'number') ? fallbackLineNum : -1;
                }
              }
              if (typeof ln !== 'number' || ln < 0) {
                console.debug('[ep_data_tables:postWriteCanonicalize] abort: invalid line', { nodeId, ln });
                return;
              }
              let attrStr = null;
              try { attrStr = docManager.getAttributeOnLine(ln, ATTR_TABLE_JSON); } catch (_) { attrStr = null; }
              if (!attrStr) {
                console.debug('[ep_data_tables:postWriteCanonicalize] abort: no tbljson on resolved line', { ln, nodeId, capturedTblId, capturedRow });
                return;
              }
              let metaAttr = null;
              try { metaAttr = JSON.parse(attrStr); } catch (_) { metaAttr = null; }
              if (!metaAttr || metaAttr.tblId !== capturedTblId || metaAttr.row !== capturedRow) {
                console.debug('[ep_data_tables:postWriteCanonicalize] abort: meta mismatch', { ln, metaAttr, capturedTblId, capturedRow });
                return;
              }
              if (typeof metaAttr.cols !== 'number') {
                console.debug('[ep_data_tables:postWriteCanonicalize] abort: invalid cols', { ln, metaAttr });
                return;
              }
              const entry = rep.lines.atIndex(ln);
              const currentText = entry?.text || '';
              const segs = currentText.split(DELIMITER);
              const needed = metaAttr.cols;
              const sanitize = (s) => {
                const x = normalizeSoftWhitespace((s || '').replace(new RegExp(DELIMITER, 'g'), ' ').replace(/[\u200B\u200C\u200D\uFEFF]/g, ' '));
                return x || ' ';
              };
              const cells = new Array(needed);
                for (let i = 0; i < needed; i++) cells[i] = sanitize(segs[i] || ' ');
              const canonical = cells.join(DELIMITER);
              
              // CRITICAL: Check if line actually has a table DOM element
              // Text matching is NOT sufficient - the table structure may have been destroyed
              let lineHasTableDOM = false;
              try {
                const lineNode = entry?.lineNode;
                if (lineNode && typeof lineNode.querySelector === 'function') {
                  const tableEl = lineNode.querySelector(
                    `table.dataTable[data-tblId="${capturedTblId}"], table.dataTable[data-tblid="${capturedTblId}"]`
                  );
                  lineHasTableDOM = !!tableEl;
                }
              } catch (_) {}
              
              const textMatches = canonical === currentText;
              const needsRepair = !textMatches || !lineHasTableDOM;
              
              // CRITICAL: Extract ALL character-level attributes BEFORE replacement to preserve them
              // Uses a generic approach - any class that isn't table-related is treated as a styling attribute
              const extractedStyling = [];
              try {
                const lineNode = entry?.lineNode;
                if (lineNode && lineHasTableDOM) {
                  const tableEl = lineNode.querySelector('table.dataTable');
                  if (tableEl) {
                    const tds = tableEl.querySelectorAll('td');
                    tds.forEach((td, cellIdx) => {
                      // Get all styled spans in this cell (skip delimiters and caret anchors)
                      const spans = td.querySelectorAll('span:not(.ep-data_tables-delim):not(.ep-data_tables-caret-anchor)');
                      let relPos = 0; // Position relative to cell start
                      spans.forEach((span) => {
                        const text = (span.textContent || '').replace(/\u00A0/g, ' '); // Normalize nbsp
                        const textLen = text.length;
                        if (textLen === 0) return;
                        
                        // Extract ALL classes except table-related ones
                        // Convert class names to Etherpad attribute format
                        const stylingAttrs = [];
                        if (span.classList) {
                          for (const cls of span.classList) {
                            // Skip table-related classes
                            if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) continue;
                            if (cls === 'ace-line' || cls.startsWith('ep-data_tables-')) continue;
                            
                            // Parse class name to attribute key-value pair
                            // Etherpad uses formats like: "author-xyz", "font-size:12", "bold"
                            if (cls.includes(':')) {
                              // Format: "key:value" (e.g., "font-size:12", "color:red")
                              const colonIdx = cls.indexOf(':');
                              const key = cls.substring(0, colonIdx);
                              const value = cls.substring(colonIdx + 1);
                              stylingAttrs.push([key, value]);
                            } else if (cls.includes('-')) {
                              // Format: "key-value" (e.g., "author-a1b2c3")
                              const dashIdx = cls.indexOf('-');
                              const key = cls.substring(0, dashIdx);
                              const value = cls.substring(dashIdx + 1);
                              stylingAttrs.push([key, value]);
                            } else {
                              // Format: "key" only (e.g., "bold", "italic")
                              // These are boolean attributes set to "true"
                              stylingAttrs.push([cls, 'true']);
                            }
                          }
                        }
                        
                        if (stylingAttrs.length > 0) {
                          extractedStyling.push({
                            cellIdx,
                            relStart: relPos,
                            len: textLen,
                            text: text, // Store text for matching
                            attrs: stylingAttrs,
                          });
                        }
                        relPos += textLen;
                      });
                    });
                  }
                }
                if (extractedStyling.length > 0) {
                  console.debug('[ep_data_tables:postWriteCanonicalize] extracted styling', {
                    ln, count: extractedStyling.length, 
                    sample: extractedStyling.slice(0, 2).map(s => ({ cell: s.cellIdx, attrs: s.attrs })),
                  });
                }
              } catch (extractErr) {
                console.debug('[ep_data_tables:postWriteCanonicalize] styling extraction error (non-fatal)', extractErr?.message);
              }
              
              if (needsRepair) {
                console.debug('[ep_data_tables:postWriteCanonicalize] line needs repair', {
                  ln,
                  textMatches,
                  lineHasTableDOM,
                  fromLen: currentText.length,
                  toLen: canonical.length,
                });
                if (!textMatches) {
                ace.ace_performDocumentReplaceRange([ln, 0], [ln, currentText.length], canonical);
                }
                // Note: If lineHasTableDOM is false, orphan detection below will handle finding the real table
              } else {
                console.debug('[ep_data_tables:postWriteCanonicalize] line already canonical', { ln, lineHasTableDOM });
              }
              let offset = 0;
              for (let i = 0; i < cells.length; i++) {
                const len = cells[i].length;
                if (len > 0) {
                  ace.ace_performDocumentApplyAttributesToRange([ln, offset], [ln, offset + len], [[ATTR_CELL, String(i)]]);
                }
                offset += len;
                if (i < cells.length - 1) offset += DELIMITER.length;
              }
              
              // Re-apply extracted styling attributes using TEXT-BASED MATCHING
              // This handles cases where text content shifts (e.g., Grammarly changes)
              if (extractedStyling.length > 0) {
                let appliedCount = 0;
                try {
                  for (const style of extractedStyling) {
                    const cellIdx = style.cellIdx;
                    const cellContent = cells[cellIdx] || '';
                    const styledText = style.text;
                    
                    if (!styledText || styledText.length === 0 || !cellContent) continue;
                    
                    // Find where the styled text appears in the NEW cell content
                    let foundPos = cellContent.indexOf(styledText);
                    
                    // If not found, try case-insensitive match (for capitalization changes)
                    if (foundPos === -1) {
                      const lowerCell = cellContent.toLowerCase();
                      const lowerStyled = styledText.toLowerCase();
                      foundPos = lowerCell.indexOf(lowerStyled);
                    }
                    
                    if (foundPos === -1) continue; // Text not found in cell
                    
                    // Calculate absolute position in the line
                    let absStart = 0;
                    for (let c = 0; c < cellIdx; c++) {
                      absStart += (cells[c]?.length || 0) + DELIMITER.length;
                    }
                    absStart += foundPos;
                    const absEnd = absStart + styledText.length;
                    
                    if (style.attrs.length > 0) {
                      ace.ace_performDocumentApplyAttributesToRange(
                        [ln, absStart],
                        [ln, absEnd],
                        style.attrs
                      );
                      appliedCount++;
                    }
                  }
                  if (appliedCount > 0) {
                    console.debug('[ep_data_tables:postWriteCanonicalize] re-applied styling', {
                      ln, applied: appliedCount, total: extractedStyling.length,
                    });
                  }
                } catch (applyErr) {
                  console.debug('[ep_data_tables:postWriteCanonicalize] styling re-apply error (non-fatal)', applyErr?.message);
                }
              }
              
              try {
                ed.ep_data_tables_applyMeta(ln, metaAttr.tblId, metaAttr.row, metaAttr.cols, ace.ace_getRep(), ed, JSON.stringify(metaAttr), docManager);
              } catch (metaErr) {
                console.error('[ep_data_tables:postWriteCanonicalize] meta apply error', metaErr);
              }
              
              // CRITICAL: Skip orphan detection for fresh paste content
              // If the PRIMARY line doesn't have a rendered table yet, this is fresh content
              // that should be left to render normally, not aggressively merged.
              // Orphan detection is only appropriate when there's an EXISTING table that got corrupted.
              if (!lineHasTableDOM) {
                console.debug('[ep_data_tables:postWriteCanonicalize] skipping orphan detection (no table on primary line - fresh paste)', {
                  ln, capturedTblId, capturedRow,
                });
                // Skip to the end - let normal rendering proceed
                return;
              }
              
              // Detect orphan lines: lines with same tblId/row attribute OR
              // lines with tbljson-* CSS classes but no table element (DOM-based orphan detection)
              // This preserves user data that was split into multiple lines due to composition/IME corruption
              const orphanLines = [];
              const seenOrphanLines = new Set(); // Avoid duplicate detection
              try {
                const total = rep.lines.length();
                for (let li = 0; li < total; li++) {
                  if (li === ln) continue;
                  
                  // Method 1: Check line attribute (existing approach)
                  let sOther = null;
                  try { sOther = docManager.getAttributeOnLine(li, ATTR_TABLE_JSON); } catch (_) { sOther = null; }
                  if (sOther) {
                  let mOther = null;
                  try { mOther = JSON.parse(sOther); } catch (_) { mOther = null; }
                  if (mOther && mOther.tblId === capturedTblId && mOther.row === capturedRow) {
                      const orphanEntry = rep.lines.atIndex(li);
                      const orphanText = orphanEntry?.text || '';
                      orphanLines.push({ lineNum: li, text: orphanText, meta: mOther, source: 'attr' });
                      seenOrphanLines.add(li);
                      continue;
                    }
                  }
                  
                  // Method 2: DOM-based detection - look for tbljson-* OR tblCell-* class spans WITHOUT a table parent
                  // This catches orphan lines that have CSS classes but no line-level attribute
                  // NOTE: Browser may strip tbljson-* during IME but keep tblCell-*, so check both!
                  if (seenOrphanLines.has(li)) continue;
                  try {
                    const lineEntry = rep.lines.atIndex(li);
                    const lineNode = lineEntry?.lineNode;
                    if (lineNode) {
                      // Check if this line has tbljson-* or tblCell-* class elements but NO table element
                      const hasTable = lineNode.querySelector('table.dataTable');
                      if (!hasTable) {
                        // Check for tbljson-* spans first (can decode metadata)
                        const tbljsonSpan = lineNode.querySelector('[class*="tbljson-"]');
                        if (tbljsonSpan) {
                          // Extract tblId from the class
                          for (const cls of tbljsonSpan.classList) {
                            if (cls.startsWith('tbljson-')) {
                              try {
                                const decoded = JSON.parse(atob(cls.substring(8)));
                                if (decoded && decoded.tblId === capturedTblId && decoded.row === capturedRow) {
                                  const orphanText = lineEntry?.text || '';
                                  orphanLines.push({
                                    lineNum: li,
                                    text: orphanText,
                                    meta: decoded,
                                    source: 'dom-class-tbljson',
                                  });
                                  seenOrphanLines.add(li);
                                  console.debug('[ep_data_tables:postWriteCanonicalize] detected tbljson DOM orphan', {
                                    lineNum: li, tblId: decoded.tblId, row: decoded.row, textLen: orphanText.length,
                                  });
                                  break;
                                }
                              } catch (_) {}
                            }
                          }
                        }
                        
                        // Also check for tblCell-* spans (no encoded metadata)
                        // This catches orphans where browser stripped tbljson-* but kept tblCell-*
                        // CRITICAL: Only use this fallback if we can verify this line belongs to OUR table
                        // Don't assume random tblCell-* spans belong to us - they could be fresh paste!
                        if (!seenOrphanLines.has(li)) {
                          const tblCellSpan = lineNode.querySelector('[class*="tblCell-"]');
                          if (tblCellSpan) {
                            // SAFETY CHECK: Verify this line has OUR tbljson attribute, just row might differ
                            // If it has tbljson for a DIFFERENT tblId, it's NOT our orphan
                            let belongsToOurTable = false;
                            try {
                              const lineAttr = docManager.getAttributeOnLine(li, ATTR_TABLE_JSON);
                              if (lineAttr) {
                                const lineMeta = JSON.parse(lineAttr);
                                if (lineMeta && lineMeta.tblId === capturedTblId) {
                                  belongsToOurTable = true;
                                }
                              } else {
                                // No tbljson attribute - check if there's a tbljson-* CLASS that we can decode
                                const anyTbljsonSpan = lineNode.querySelector('[class*="tbljson-"]');
                                if (anyTbljsonSpan) {
                                  for (const cls of anyTbljsonSpan.classList) {
                                    if (cls.startsWith('tbljson-')) {
                                      try {
                                        const decoded = JSON.parse(atob(cls.substring(8)));
                                        if (decoded && decoded.tblId === capturedTblId) {
                                          belongsToOurTable = true;
                                        }
                                      } catch (_) {}
                                      break;
                                    }
                                  }
                                }
                              }
                            } catch (_) {}
                            
                            if (belongsToOurTable) {
                              const orphanText = lineEntry?.text || '';
                              // Only add if the text contains delimiters or meaningful content
                              if (orphanText.includes(DELIMITER) || orphanText.trim().length > 0) {
                                orphanLines.push({
                                  lineNum: li,
                                  text: orphanText,
                                  meta: { tblId: capturedTblId, row: capturedRow, cols: expectedCols },
                                  source: 'dom-class-tblCell',
                                });
                                seenOrphanLines.add(li);
                                console.debug('[ep_data_tables:postWriteCanonicalize] detected tblCell DOM orphan', {
                                  lineNum: li, capturedTblId, capturedRow, textLen: orphanText.length,
                                });
                              }
                            } else {
                              console.debug('[ep_data_tables:postWriteCanonicalize] skipping tblCell line (different table)', {
                                lineNum: li, capturedTblId,
                              });
                            }
                          }
                        }
                      }
                    }
                  } catch (domErr) {
                    console.error('[ep_data_tables:postWriteCanonicalize] DOM orphan detection error', domErr);
                  }
                }
              } catch (scanOrphanErr) {
                console.error('[ep_data_tables:postWriteCanonicalize] orphan scan error', scanOrphanErr);
                  }
              
              if (orphanLines.length > 0) {
                // CRITICAL: Re-evaluate primary line selection using LIVE DOM query
                // The lineNode references from rep.lines might be stale after DOM mutations.
                // Query the actual document to find which ace-line currently has the table.
                let mainEntry = rep.lines.atIndex(ln);
                
                // Use live DOM query to find the line that actually has our table
                let lineWithTableIdx = -1;
                  try {
                  const innerDoc = node?.ownerDocument || document;
                  const editorBody = innerDoc.getElementById('innerdocbody') || innerDoc.body;
                  if (editorBody) {
                    const allAceLines = editorBody.querySelectorAll('div.ace-line');
                    for (let ai = 0; ai < allAceLines.length; ai++) {
                      const aceLine = allAceLines[ai];
                      // Look for our specific table by tblId and row
                      const tableEl = aceLine.querySelector(
                        `table.dataTable[data-tblId="${capturedTblId}"][data-row="${capturedRow}"], ` +
                        `table.dataTable[data-tblid="${capturedTblId}"][data-row="${capturedRow}"]`
                      );
                      if (tableEl) {
                        lineWithTableIdx = ai;
                        console.debug('[ep_data_tables:postWriteCanonicalize] LIVE DOM found table', {
                          domIndex: ai, aceLineId: aceLine.id, tblId: capturedTblId, row: capturedRow,
                        });
                              break;
                            }
                          }
                  }
                } catch (domQueryErr) {
                  console.error('[ep_data_tables:postWriteCanonicalize] live DOM query error', domQueryErr);
                }
                
                // Check if we need to swap primary
                // If the live DOM shows the table is on a different line than our current primary,
                // we need to swap to use that line as the primary
                if (lineWithTableIdx >= 0 && lineWithTableIdx !== ln) {
                  // Find if this line is in our orphan list
                  const orphanIdx = orphanLines.findIndex(o => o.lineNum === lineWithTableIdx);
                  if (orphanIdx >= 0) {
                    console.warn('[ep_data_tables:postWriteCanonicalize] SWAPPING primary via LIVE DOM: table found on orphan line', {
                      oldPrimary: ln,
                      newPrimary: lineWithTableIdx,
                      orphanIdx,
                    });
                    
                    // Move current "primary" to orphan list
                    const formerPrimary = {
                      lineNum: ln,
                      text: mainEntry?.text || '',
                      meta: metaAttr,
                      source: 'swapped-to-orphan',
                    };

                    // Swap: orphan becomes primary
                    const newPrimary = orphanLines[orphanIdx];
                    ln = newPrimary.lineNum;
                    metaAttr = newPrimary.meta;
                    mainEntry = rep.lines.atIndex(ln);
                    
                    // Remove new primary from orphan list and add former primary
                    orphanLines.splice(orphanIdx, 1);
                    orphanLines.push(formerPrimary);
                    
                    // Re-read cells from new primary
                    const newText = mainEntry?.text || '';
                    const newSegs = newText.split(DELIMITER);
                    for (let i = 0; i < needed; i++) {
                      cells[i] = sanitize(newSegs[i] || ' ');
                    }
                  }
                } else if (lineWithTableIdx < 0) {
                  // No line has the table - this is bad, but we should still try to preserve content
                  console.warn('[ep_data_tables:postWriteCanonicalize] WARNING: no line has table DOM for tblId/row', {
                    capturedTblId, capturedRow, primaryLn: ln, orphanCount: orphanLines.length,
                  });
                }
                
                console.warn('[ep_data_tables:postWriteCanonicalize] orphan lines detected - merging content', {
                  keepLine: ln, orphans: orphanLines.map(o => ({ line: o.lineNum, textLen: o.text.length })),
                  tableFoundOnLine: lineWithTableIdx,
                });
                
                // CRITICAL FIX: Extract cell content from LIVE DOM table, not from line text!
                // The line text may contain non-table content (like "1" from a previous line)
                // that got mixed in due to IME corruption. The DOM table is the source of truth.
                const mainText = mainEntry?.text || '';
                const mainSegs = mainText.split(DELIMITER);
                const mergedCells = new Array(needed);
                
                // Try to get clean cell content from the DOM table
                let usedDomContent = false;
                if (lineWithTableIdx >= 0) {
                  try {
                    const innerDoc = node?.ownerDocument || document;
                    const editorBody = innerDoc.getElementById('innerdocbody') || innerDoc.body;
                    const allAceLines = editorBody?.querySelectorAll('div.ace-line');
                    const tableAceLine = allAceLines?.[lineWithTableIdx];
                    const tableEl = tableAceLine?.querySelector(
                      `table.dataTable[data-tblId="${capturedTblId}"], table.dataTable[data-tblid="${capturedTblId}"]`
                    );
                    if (tableEl) {
                      const tr = tableEl.querySelector('tbody > tr');
                      if (tr && tr.children.length === needed) {
                        for (let i = 0; i < needed; i++) {
                          const td = tr.children[i];
                          // Extract text content, excluding delimiters and special elements
                          let cellText = '';
                          for (const child of td.childNodes) {
                            if (child.nodeType === 3) { // Text node
                              cellText += child.textContent || '';
                            } else if (child.nodeType === 1) { // Element
                              const el = child;
                              // Skip delimiter spans and caret anchors
                              if (el.classList?.contains('ep-data_tables-delim')) continue;
                              if (el.classList?.contains('ep-data_tables-caret-anchor')) continue;
                              if (el.classList?.contains('ep-data_tables-resize-handle')) continue;
                              cellText += el.textContent || '';
                            }
                          }
                          mergedCells[i] = sanitize(cellText || ' ');
                        }
                        usedDomContent = true;
                        console.debug('[ep_data_tables:postWriteCanonicalize] using DOM table content for merge base', {
                          cells: mergedCells.map(c => c.slice(0, 20)),
                        });
                      }
                    }
                  } catch (domExtractErr) {
                    console.error('[ep_data_tables:postWriteCanonicalize] DOM content extraction error', domExtractErr);
                  }
                }
                
                // Fallback to line text if DOM extraction failed
                if (!usedDomContent) {
                for (let i = 0; i < needed; i++) {
                  mergedCells[i] = sanitize(mainSegs[i] || ' ');
                  }
                  console.debug('[ep_data_tables:postWriteCanonicalize] using line text for merge base (DOM extraction failed)');
                }

                // Extract content from orphan lines and merge into cells
                        for (const orphan of orphanLines) {
                  const orphanSegs = orphan.text.split(DELIMITER);
                  // Each orphan segment represents content that should be in corresponding cell
                  // Only merge if orphan has non-trivial content
                  for (let i = 0; i < Math.min(orphanSegs.length, needed); i++) {
                    const orphanContent = sanitize(orphanSegs[i] || '');
                    // Skip if orphan cell is empty or just whitespace
                    if (!orphanContent || orphanContent.trim() === '') continue;
                    // Skip if main cell already contains this content (avoid duplication)
                    const mainCellTrimmed = (mergedCells[i] || '').trim();
                    const orphanTrimmed = orphanContent.trim();
                    if (mainCellTrimmed.includes(orphanTrimmed)) continue;
                    // If main cell is empty/whitespace, replace; otherwise append
                    if (!mainCellTrimmed) {
                      mergedCells[i] = orphanContent;
                    } else {
                      // Append orphan content to existing cell content (preserve user data)
                      mergedCells[i] = mainCellTrimmed + orphanTrimmed;
                                  }
                    console.debug('[ep_data_tables:postWriteCanonicalize] merged orphan content', {
                      cellIdx: i, orphanLine: orphan.lineNum, orphanContent, mergedResult: mergedCells[i],
                    });
                  }
                }
                
                // Apply merged content to main line
                const mergedCanonical = mergedCells.join(DELIMITER);
                if (mergedCanonical !== mainText) {
                  console.debug('[ep_data_tables:postWriteCanonicalize] applying merged canonical line', {
                    ln, fromLen: mainText.length, toLen: mergedCanonical.length, mergedCells,
                  });
                  ace.ace_performDocumentReplaceRange([ln, 0], [ln, mainText.length], mergedCanonical);

                  // Re-apply cell attributes after merge
                  let mergeOffset = 0;
                  for (let i = 0; i < mergedCells.length; i++) {
                    const cellLen = mergedCells[i].length;
                        if (cellLen > 0) {
                      ace.ace_performDocumentApplyAttributesToRange([ln, mergeOffset], [ln, mergeOffset + cellLen], [[ATTR_CELL, String(i)]]);
                        }
                    mergeOffset += cellLen;
                    if (i < mergedCells.length - 1) mergeOffset += DELIMITER.length;
                  }
                  
                  // CRITICAL: Re-apply line-level tbljson attribute after merge
                  // Without this, the line won't render as a table!
                  try {
                    const repAfterMerge = ace.ace_getRep();
                    ed.ep_data_tables_applyMeta(ln, capturedTblId, capturedRow, expectedCols, repAfterMerge, ed, JSON.stringify(metaAttr), docManager);
                    console.debug('[ep_data_tables:postWriteCanonicalize] re-applied tbljson after merge', {
                      ln, tblId: capturedTblId, row: capturedRow, cols: expectedCols,
                    });
                  } catch (metaMergeErr) {
                    console.error('[ep_data_tables:postWriteCanonicalize] failed to re-apply tbljson after merge', metaMergeErr);
                  }
                      }

                // Now safely delete orphan lines (bottom-up to preserve line numbers)
                // CRITICAL: Validate each line before operating to prevent keyToNodeMap errors
                orphanLines.sort((a, b) => b.lineNum - a.lineNum).forEach((orphan) => {
                    try {
                      // Re-fetch rep to get current state after any previous deletions
                      const repCheck = ace.ace_getRep();
                      if (!repCheck || !repCheck.lines) {
                        console.warn('[ep_data_tables:postWriteCanonicalize] rep invalid, skipping orphan removal');
                        return;
                      }
                      const currentLineCount = repCheck.lines.length();
                      if (orphan.lineNum >= currentLineCount) {
                        console.debug('[ep_data_tables:postWriteCanonicalize] orphan line no longer exists', {
                          orphanLine: orphan.lineNum, currentLineCount,
                        });
                        return;
                      }
                      // Verify line entry exists
                      const orphanEntry = repCheck.lines.atIndex(orphan.lineNum);
                      if (!orphanEntry) {
                        console.debug('[ep_data_tables:postWriteCanonicalize] orphan line has no entry', {
                          orphanLine: orphan.lineNum,
                        });
                        return;
                      }
                      
                      try {
                        if (docManager && typeof docManager.removeAttributeOnLine === 'function') {
                        docManager.removeAttributeOnLine(orphan.lineNum, ATTR_TABLE_JSON);
                        }
                    } catch (remErr) {
                      console.debug('[ep_data_tables:postWriteCanonicalize] removeAttributeOnLine error (non-fatal)', remErr?.message);
                    }
                    console.debug('[ep_data_tables:postWriteCanonicalize] removing orphan line (content already merged)', {
                      orphanLine: orphan.lineNum,
                    });
                    ace.ace_performDocumentReplaceRange([orphan.lineNum, 0], [orphan.lineNum + 1, 0], '');
                  } catch (orphanRemErr) {
                    console.error('[ep_data_tables:postWriteCanonicalize] orphan line removal error', {
                      orphanLine: orphan.lineNum,
                      error: orphanRemErr?.message || orphanRemErr,
                    });
                    }
                  });
                  
                // Clean up spurious blank lines between this table row and the next
                // These can be created when orphan content gets merged and lines shift
                try {
                  const repAfterOrphanRemoval = ace.ace_getRep();
                  const currentLineNum = ln; // The line we just merged into
                  const nextRowNum = capturedRow + 1;
                  
                  // Find the line with the next table row (if any)
                  let nextRowLineNum = -1;
                  const totalAfter = repAfterOrphanRemoval.lines.length();
                  for (let li = currentLineNum + 1; li < totalAfter && li < currentLineNum + 10; li++) {
                    try {
                      const attrStr = docManager.getAttributeOnLine(li, ATTR_TABLE_JSON);
                      if (attrStr) {
                        const meta = JSON.parse(attrStr);
                        if (meta && meta.tblId === capturedTblId && meta.row === nextRowNum) {
                          nextRowLineNum = li;
                          break;
                        }
                      }
                    } catch (_) {}
                  }
                  
                  // If there are blank lines between current row and next row, remove them
                  if (nextRowLineNum > currentLineNum + 1) {
                    const blankLinesToRemove = [];
                    for (let li = currentLineNum + 1; li < nextRowLineNum; li++) {
                      const lineEntry = repAfterOrphanRemoval.lines.atIndex(li);
                      const lineText = lineEntry?.text || '';
                      // Check if line is blank (empty or just whitespace)
                      if (!lineText.trim() || lineText === '\n') {
                        blankLinesToRemove.push(li);
                      }
                    }
                    
                    // Remove blank lines bottom-up to preserve line numbers
                    blankLinesToRemove.sort((a, b) => b - a).forEach((blankLineNum) => {
                      try {
                        console.debug('[ep_data_tables:postWriteCanonicalize] removing spurious blank line between table rows', {
                          blankLineNum, betweenRows: [capturedRow, nextRowNum],
                        });
                        ace.ace_performDocumentReplaceRange([blankLineNum, 0], [blankLineNum + 1, 0], '');
                      } catch (blankRemErr) {
                        console.error('[ep_data_tables:postWriteCanonicalize] blank line removal error', blankRemErr);
                      }
                    });
                  }
                } catch (cleanupErr) {
                  console.error('[ep_data_tables:postWriteCanonicalize] blank line cleanup error', cleanupErr);
                }
                  
                  try { ace.ace_fastIncorp(5); } catch (_) {}
                }
              
              // ALWAYS run blank line cleanup after repair, even if there were no orphans
              // Grammarly and other extensions can create blank lines without triggering orphan detection
              try {
                const repFinal = ace.ace_getRep();
                if (repFinal && repFinal.lines) {
                  const finalLineNum = ln; // The line we repaired
                  const totalFinal = repFinal.lines.length();
                  
                  // Look for blank lines immediately after this table row
                  const blankLinesToClean = [];
                  for (let li = finalLineNum + 1; li < totalFinal && li < finalLineNum + 5; li++) {
                    const lineEntry = repFinal.lines.atIndex(li);
                    const lineText = lineEntry?.text || '';
                    
                    // Check if this is a table row (should stop scanning)
                    let isTableRow = false;
                    try {
                      const attrStr = docManager.getAttributeOnLine(li, ATTR_TABLE_JSON);
                      if (attrStr) isTableRow = true;
                    } catch (_) {}
                    
                    if (isTableRow) break; // Stop at next table row
                    
                    // Check if blank line (empty or just whitespace)
                    if (!lineText.trim() || lineText === '\n') {
                      blankLinesToClean.push(li);
                    } else {
                      break; // Stop at first non-blank, non-table line
                    }
                  }
                  
                  // Remove blank lines bottom-up
                  if (blankLinesToClean.length > 0) {
                    blankLinesToClean.sort((a, b) => b - a).forEach((blankLineNum) => {
                      try {
                        const checkRep = ace.ace_getRep();
                        if (checkRep && blankLineNum < checkRep.lines.length()) {
                          console.debug('[ep_data_tables:postWriteCanonicalize] removing blank line after repair', {
                            blankLineNum, afterRow: capturedRow,
                          });
                          ace.ace_performDocumentReplaceRange([blankLineNum, 0], [blankLineNum + 1, 0], '');
                        }
                      } catch (_) {}
                    });
                  }
                }
              } catch (finalCleanupErr) {
                console.debug('[ep_data_tables:postWriteCanonicalize] final blank cleanup error (non-fatal)', finalCleanupErr?.message);
              }
              
              try { ace.ace_fastIncorp(5); } catch (_) {}
              } catch (innerErr) {
                console.error('[ep_data_tables:postWriteCanonicalize] inner callback error', {
                  error: innerErr?.message || innerErr,
                });
            }
          }, 'ep_data_tables:postwrite-canonicalize', true);
          } catch (aceCallErr) {
            // This catches errors from ace_callWithAce itself (Etherpad internal state corruption)
            console.warn('[ep_data_tables:postWriteCanonicalize] ace_callWithAce failed - editor state may need refresh', {
              error: aceCallErr?.message || aceCallErr,
            });
          }
        } finally {
          __epDT_postWriteScheduled.delete(nodeId);
          // Clear the original line info if it was for this table
          if (__epDT_compositionOriginalLine.tblId === capturedTblId) {
            __epDT_compositionOriginalLine = { tblId: null, lineNum: null, timestamp: 0 };
          }
        }
      }, 0);
    }
    // Stop here; let the next render reflect the canonicalized text.
    // BUT: if column operation is in progress, continue to render the table with current segments
    if (!skipMismatchReturn) {
    return cb();
    }
    // Otherwise, fall through to table building below
  } else {
     // log(`${logPrefix} NodeID#${nodeId}: Segment count matches metadata cols (${rowMetadata.cols}). Using original segments.`);
  }

 // log(`${logPrefix} NodeID#${nodeId}: Calling buildTableFromDelimitedHTML...`);
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
     // log(`${logPrefix} NodeID#${nodeId}: Received new table HTML from helper. Replacing content.`);

      const tbljsonElement = findTbljsonElement(node);

      if (tbljsonElement && tbljsonElement.parentElement && tbljsonElement.parentElement !== node) {
        const parentTag = tbljsonElement.parentElement.tagName.toLowerCase();
    const blockElements = ['center', 'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'right', 'left', 'ul', 'ol', 'li', 'code'];

        if (blockElements.includes(parentTag)) {
         // log(`${logPrefix} NodeID#${nodeId}: Preserving block element ${parentTag} and replacing its content with table.`);
          tbljsonElement.parentElement.innerHTML = newTableHTML;
    } else {
         // log(`${logPrefix} NodeID#${nodeId}: Parent element ${parentTag} is not a block element, replacing entire node content.`);
      node.innerHTML = newTableHTML;
    }
      } else {
       // log(`${logPrefix} NodeID#${nodeId}: No nested block element found, replacing entire node content.`);
      node.innerHTML = newTableHTML;
      }

     // log(`${logPrefix} NodeID#${nodeId}: Successfully replaced content with new table structure.`);
  } catch (renderError) {
     // log(`${logPrefix} NodeID#${nodeId}: ERROR during table building or rendering.`, renderError);
      console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Error building/rendering table.`, renderError);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Failed to render table structure.</div>';
     // log(`${logPrefix} NodeID#${nodeId}: Rendered build/render error message in node. END.`);
      return cb();
  }


 // log(`${logPrefix}: ----- END ----- NodeID: ${nodeId}`);
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
 // log(`${logPrefix} START Key='${evt?.key}' Code=${evt?.keyCode} Type=${evt?.type} Modifiers={ctrl:${evt?.ctrlKey},alt:${evt?.altKey},meta:${evt?.metaKey},shift:${evt?.shiftKey}}`, { selStart: rep?.selStart, selEnd: rep?.selEnd });

  if (!rep || !rep.selStart || !editorInfo || !evt || !docManager) {
   // log(`${logPrefix} Skipping - Missing critical context.`);
    return false;
  }

  const reportedLineNum = rep.selStart[0];
  const reportedCol = rep.selStart[1]; 
 // log(`${logPrefix} Reported caret from rep: Line=${reportedLineNum}, Col=${reportedCol}`);

  let tableMetadata = null;
  let lineAttrString = null;
  try {
   // log(`${logPrefix} DEBUG: Attempting to get ${ATTR_TABLE_JSON} attribute from line ${reportedLineNum}`);
    lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
   // log(`${logPrefix} DEBUG: getAttributeOnLine returned: ${lineAttrString ? `"${lineAttrString}"` : 'null/undefined'}`);

    if (typeof docManager.getAttributesOnLine === 'function') {
      try {
        const allAttribs = docManager.getAttributesOnLine(reportedLineNum);
       // log(`${logPrefix} DEBUG: All attributes on line ${reportedLineNum}:`, allAttribs);
      } catch(e) {
       // log(`${logPrefix} DEBUG: Error getting all attributes:`, e);
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
           // log(`${logPrefix} DEBUG: Found table in DOM without attribute! TblId=${domTblId}, Row=${domRow}`);
            const domCells = tableInDOM.querySelectorAll('td');
            if (domTblId && domRow !== null && domCells.length > 0) {
             // log(`${logPrefix} DEBUG: Attempting to reconstruct metadata from DOM...`);
              const reconstructedMetadata = {
                tblId: domTblId,
                row: parseInt(domRow, 10),
                cols: domCells.length
              };
              lineAttrString = JSON.stringify(reconstructedMetadata);
             // log(`${logPrefix} DEBUG: Reconstructed metadata: ${lineAttrString}`);
            }
          }
        }
      } catch(e) {
       // log(`${logPrefix} DEBUG: Error checking DOM for table:`, e);
      }
    }

    if (lineAttrString) {
        tableMetadata = JSON.parse(lineAttrString);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
            // log(`${logPrefix} Line ${reportedLineNum} has attribute, but metadata invalid/missing cols.`);
             tableMetadata = null;
        }
    } else {
       // log(`${logPrefix} DEBUG: No ${ATTR_TABLE_JSON} attribute found on line ${reportedLineNum}`);
    }
  } catch(e) {
    console.error(`${logPrefix} Error checking/parsing line attribute for line ${reportedLineNum}.`, e);
    tableMetadata = null;
  }

  const editor = editorInfo.editor;
  const lastClick = editor?.ep_data_tables_last_clicked;
 // log(`${logPrefix} Reading stored click/caret info:`, lastClick);

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
     // log(`${logPrefix} Attempting to validate stored click info for Line=${lastClick.lineNum}...`);
      let storedLineAttrString = null;
      let storedLineMetadata = null;
      try {
         // log(`${logPrefix} DEBUG: Getting ${ATTR_TABLE_JSON} attribute from stored line ${lastClick.lineNum}`);
          storedLineAttrString = docManager.getAttributeOnLine(lastClick.lineNum, ATTR_TABLE_JSON);
         // log(`${logPrefix} DEBUG: Stored line attribute result: ${storedLineAttrString ? `"${storedLineAttrString}"` : 'null/undefined'}`);

          if (storedLineAttrString) {
            storedLineMetadata = JSON.parse(storedLineAttrString);
           // log(`${logPrefix} DEBUG: Parsed stored metadata:`, storedLineMetadata);
          }

          if (storedLineMetadata && typeof storedLineMetadata.cols === 'number' && storedLineMetadata.tblId === lastClick.tblId) {
             // log(`${logPrefix} Stored click info VALIDATED (Metadata OK and tblId matches). Trusting stored state.`);
              trustedLastClick = true;
              currentLineNum = lastClick.lineNum; 
              targetCellIndex = lastClick.cellIndex;
              metadataForTargetLine = storedLineMetadata; 
              lineAttrString = storedLineAttrString;

              lineText = rep.lines.atIndex(currentLineNum)?.text || '';
              cellTexts = lineText.split(DELIMITER);
             // log(`${logPrefix} Using Line=${currentLineNum}, CellIndex=${targetCellIndex}. Text: "${lineText}"`);

              if (cellTexts.length !== metadataForTargetLine.cols) {
                 // log(`${logPrefix} WARNING: Stored cell count mismatch for trusted line ${currentLineNum}.`);
              }

              cellStartCol = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                  cellStartCol += (cellTexts[i]?.length ?? 0) + DELIMITER.length;
              }
              precedingCellsOffset = cellStartCol;
             // log(`${logPrefix} Calculated cellStartCol=${cellStartCol} from trusted cellIndex=${targetCellIndex}.`);

              if (typeof lastClick.relativePos === 'number' && lastClick.relativePos >= 0) {
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(lastClick.relativePos, currentCellTextLength));
                 // log(`${logPrefix} Using and validated stored relative position: ${relativeCaretPos}.`);
  } else {
                  relativeCaretPos = reportedCol - cellStartCol;
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(relativeCaretPos, currentCellTextLength)); 
                 // log(`${logPrefix} Stored relativePos missing, calculated from reportedCol (${reportedCol}): ${relativeCaretPos}`);
              }
          } else {
             // log(`${logPrefix} Stored click info INVALID (Metadata missing/invalid or tblId mismatch). Clearing stored state.`);
              if (editor) editor.ep_data_tables_last_clicked = null;
          }
      } catch (e) {
           console.error(`${logPrefix} Error validating stored click info for line ${lastClick.lineNum}.`, e);
           if (editor) editor.ep_data_tables_last_clicked = null;
      }
  }

  if (!trustedLastClick) {
     // log(`${logPrefix} Fallback: Using reported caret position Line=${reportedLineNum}, Col=${reportedCol}.`);
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
                 // log(`${logPrefix} Fallback: Found table in DOM without attribute! TblId=${domTblId}, Row=${domRow}`);
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domTblId && domRow !== null && domCells.length > 0) {
                   // log(`${logPrefix} Fallback: Attempting to reconstruct metadata from DOM...`);
                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length
                    };
                    lineAttrString = JSON.stringify(reconstructedMetadata);
                    tableMetadata = reconstructedMetadata;
                   // log(`${logPrefix} Fallback: Reconstructed metadata: ${lineAttrString}`);
                  }
                }
              }
            } catch(e) {
             // log(`${logPrefix} Fallback: Error checking DOM for table:`, e);
            }
          }
      } catch(e) { tableMetadata = null; }

      if (!tableMetadata) {
         // log(`${logPrefix} Fallback: Reported line ${reportedLineNum} is not a valid table line. Allowing default.`);
           return false;
      }

      currentLineNum = reportedLineNum;
      metadataForTargetLine = tableMetadata;
     // log(`${logPrefix} Fallback: Processing based on reported line ${currentLineNum}.`);

      lineText = rep.lines.atIndex(currentLineNum)?.text || '';
      cellTexts = lineText.split(DELIMITER);
     // log(`${logPrefix} Fallback: Fetched text for reported line ${currentLineNum}: "${lineText}"`);

      if (cellTexts.length !== metadataForTargetLine.cols) {
         // log(`${logPrefix} WARNING (Fallback): Cell count mismatch for reported line ${currentLineNum}.`);
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
             // log(`${logPrefix} --> (Fallback Calc) Found target cell ${foundIndex}. RelativePos: ${relativeCaretPos}.`);
              break; 
          }
          if (i < cellTexts.length - 1 && reportedCol === cellEndCol + DELIMITER.length) {
              foundIndex = i + 1;
              relativeCaretPos = 0; 
              cellStartCol = currentOffset + cellLength + DELIMITER.length;
              precedingCellsOffset = cellStartCol;
             // log(`${logPrefix} --> (Fallback Calc) Caret at delimiter AFTER cell ${i}. Treating as start of cell ${foundIndex}.`);
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
               // log(`${logPrefix} --> (Fallback Calc) Caret detected at END of last cell (${foundIndex}).`);
          } else {
           // log(`${logPrefix} (Fallback Calc) FAILED to determine target cell for caret col ${reportedCol}. Allowing default handling.`);
            return false; 
          }
      }
      targetCellIndex = foundIndex;
  }

  if (currentLineNum < 0 || targetCellIndex < 0 || !metadataForTargetLine || targetCellIndex >= metadataForTargetLine.cols) {
      // log(`${logPrefix} FAILED final validation: Line=${currentLineNum}, Cell=${targetCellIndex}, Metadata=${!!metadataForTargetLine}. Allowing default.`);
    if (editor) editor.ep_data_tables_last_clicked = null;
    return false;
  }

 // log(`${logPrefix} --> Final Target: Line=${currentLineNum}, CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}`);

  const selStartActual = rep.selStart;
  const selEndActual = rep.selEnd;
  const hasSelection = selStartActual[0] !== selEndActual[0] || selStartActual[1] !== selEndActual[1];

  if (hasSelection) {
   // log(`${logPrefix} [selection] Active selection detected. Start:[${selStartActual[0]},${selStartActual[1]}], End:[${selEndActual[0]},${selEndActual[1]}]`);
   // log(`${logPrefix} [caretTrace] [selection] Initial rep.selStart: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

    if (selStartActual[0] !== currentLineNum || selEndActual[0] !== currentLineNum) {
     // log(`${logPrefix} [selection] Selection spans multiple lines (${selStartActual[0]}-${selEndActual[0]}) or is not on the current focused table line (${currentLineNum}). Preventing default action.`);
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

   // log(`${logPrefix} [selection] Cell context for selection: targetCellIndex=${targetCellIndex}, cellStartColInLine=${cellContentStartColInLine}, cellEndColInLine=${cellContentEndColInLine}, currentCellFullText='${currentCellFullText}'`);

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
     // log(`${logPrefix} [selection] Handling key='${evt.key}' (Type: ${evt.type}) for valid intra-cell selection.`);

      if (evt.type !== 'keydown') {
       // log(`${logPrefix} [selection] Ignoring non-keydown event type ('${evt.type}') for selection handling. Allowing default.`);
        return false; 
      }
      evt.preventDefault();

      const rangeStart = [currentLineNum, selectionStartColInLine];
      const rangeEnd = [currentLineNum, selectionEndColInLine];
      let replacementText = '';
      let newAbsoluteCaretCol = selectionStartColInLine;
      const repBeforeEdit = editorInfo.ace_getRep();
     // log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

      if (isCurrentKeyTyping) {
        replacementText = evt.key;
        newAbsoluteCaretCol = selectionStartColInLine + replacementText.length;
       // log(`${logPrefix} [selection] -> Replacing selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]] with text '${replacementText}'`);
      } else {
       // log(`${logPrefix} [selection] -> Deleting selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]]`);
        const isWholeCell = selectionStartColInLine <= cellContentStartColInLine && selectionEndColInLine >= cellContentEndColInLine;
        if (isWholeCell) {
          replacementText = ' ';
          newAbsoluteCaretCol = selectionStartColInLine + 1;
         // log(`${logPrefix} [selection] Whole cell cleared – inserting single space to preserve caret/author span.`);
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
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


       // log(`${logPrefix} [selection] -> Re-applying tbljson line attribute...`);
        const applyHelper = editorInfo.ep_data_tables_applyMeta;
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) {
          const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;
          applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
         // log(`${logPrefix} [selection] -> tbljson line attribute re-applied (using rep before edit).`);
        } else {
          console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
          const currentRepFallback = editorInfo.ace_getRep();
          if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
           // log(`${logPrefix} [selection] -> Retrying attribute application with current rep...`);
            applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
           // log(`${logPrefix} [selection] -> tbljson line attribute re-applied (using current rep fallback).`);
          } else {
            console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute even with fallback rep.`);
          }
        }

       // log(`${logPrefix} [selection] -> Setting selection/caret to: [${currentLineNum}, ${newAbsoluteCaretCol}]`);
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterSelectionChange = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);

    editorInfo.ace_fastIncorp(1);
        const repAfterFastIncorp = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
       // log(`${logPrefix} [selection] -> Requested sync hint (fastIncorp 1).`);

       // log(`${logPrefix} [caretTrace] [selection] Attempting to re-assert selection post-fastIncorp to [${currentLineNum}, ${newAbsoluteCaretCol}]`);
        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterReassert = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

        const newRelativePos = newAbsoluteCaretCol - cellStartCol;
        if (editor) {
            editor.ep_data_tables_last_clicked = {
                lineNum: currentLineNum,
                tblId: metadataForTargetLine.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos
            };
           // log(`${logPrefix} [selection] -> Updated stored click/caret info:`, editor.ep_data_tables_last_clicked);
        } else {
           // log(`${logPrefix} [selection] -> Editor instance not found, cannot update ep_data_tables_last_clicked.`);
        }

       // log(`${logPrefix} END [selection] (Handled highlight modification) Key='${evt.key}' Type='${evt.type}'. Duration: ${Date.now() - startLogTime}ms`);
      return true;
      } catch (error) {
       // log(`${logPrefix} [selection] ERROR during highlight modification:`, error);
        console.error('[ep_data_tables] Error processing highlight modification:', error);
        return true;
      }
    }
  }

  const isCutKey = (evt.ctrlKey || evt.metaKey) && (evt.key === 'x' || evt.key === 'X' || evt.keyCode === 88);
  if (isCutKey && hasSelection) {
   // log(`${logPrefix} Ctrl+X (Cut) detected with selection. Letting cut event handler manage this.`);
    return false;
  } else if (isCutKey && !hasSelection) {
   // log(`${logPrefix} Ctrl+X (Cut) detected but no selection. Allowing default.`);
    return false;
  }

  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete' || evt.keyCode === 46;
  const isBackspaceKey = evt.key === 'Backspace' || evt.keyCode === 8;
  const isNavigationKey = [33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode);
  const isTabKey = evt.key === 'Tab';
  const isEnterKey = evt.key === 'Enter';
 // log(`${logPrefix} Key classification: Typing=${isTypingKey}, Backspace=${isBackspaceKey}, Delete=${isDeleteKey}, Nav=${isNavigationKey}, Tab=${isTabKey}, Enter=${isEnterKey}, Cut=${isCutKey}`);

  const currentCellTextLengthEarly = cellTexts[targetCellIndex]?.length ?? 0;

  if (evt.type === 'keydown' && !evt.ctrlKey && !evt.metaKey && !evt.altKey) {
    if (evt.keyCode === 39 && relativeCaretPos >= currentCellTextLengthEarly && targetCellIndex < metadataForTargetLine.cols - 1) {
     // log(`${logPrefix} ArrowRight at cell boundary – navigating to next cell to avoid anchor zone.`);
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, false, editorInfo, docManager);
      return true;
    }

    if (evt.keyCode === 37 && relativeCaretPos === 0 && targetCellIndex > 0) {
     // log(`${logPrefix} ArrowLeft at cell boundary – navigating to previous cell to avoid anchor zone.`);
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, true, editorInfo, docManager);
      return true;
    }
  }


  if (isNavigationKey && !isTabKey) {
     // log(`${logPrefix} Allowing navigation key: ${evt.key}. Clearing click state.`);
      if (editor) editor.ep_data_tables_last_clicked = null;
      return false;
  }

  if (isTabKey) { 
    // log(`${logPrefix} Tab key pressed. Event type: ${evt.type}`);
    evt.preventDefault();

     if (evt.type !== 'keydown') {
      // log(`${logPrefix} Ignoring Tab ${evt.type} event to prevent double navigation.`);
    return true;
  }

    // log(`${logPrefix} Processing Tab keydown - implementing cell navigation.`);
     const success = navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, evt.shiftKey, editorInfo, docManager);
     if (!success) {
      // log(`${logPrefix} Tab navigation failed, cell navigation not possible.`);
     }
     return true;
  }

  if (isEnterKey) {
     // log(`${logPrefix} Enter key pressed. Event type: ${evt.type}`);
    evt.preventDefault();

      if (evt.type !== 'keydown') {
       // log(`${logPrefix} Ignoring Enter ${evt.type} event to prevent double navigation.`);
    return true;
  }

     // log(`${logPrefix} Processing Enter keydown - implementing cell navigation.`);
      const success = navigateToCellBelow(currentLineNum, targetCellIndex, metadataForTargetLine, editorInfo, docManager);
      if (!success) {
       // log(`${logPrefix} Enter navigation failed, cell navigation not possible.`);
      }
      return true; 
  }

      const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex > 0) {
     // log(`${logPrefix} Intercepted Backspace at start of cell ${targetCellIndex}. Preventing default.`);
    evt.preventDefault();
          return true;
      }
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex === 0) {
       // log(`${logPrefix} Intercepted Backspace at start of first cell (line boundary). Preventing merge.`);
        evt.preventDefault();
        return true;
      }
  if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex < metadataForTargetLine.cols - 1) {
     // log(`${logPrefix} Intercepted Delete at end of cell ${targetCellIndex}. Preventing default.`);
          evt.preventDefault();
          return true;
      }
      if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex === metadataForTargetLine.cols - 1) {
       // log(`${logPrefix} Intercepted Delete at end of last cell (line boundary). Preventing merge.`);
        evt.preventDefault();
        return true;
      }

  const isInternalBackspace = isBackspaceKey && relativeCaretPos > 0;
  const isInternalDelete = isDeleteKey && relativeCaretPos < currentCellTextLength;

  if ((isInternalBackspace && relativeCaretPos === 1 && targetCellIndex > 0) ||
      (isInternalDelete && relativeCaretPos === 0 && targetCellIndex > 0)) {
   // log(`${logPrefix} Attempt to erase protected delimiter – operation blocked.`);
    evt.preventDefault();
    return true;
  }

  if (isTypingKey || isInternalBackspace || isInternalDelete) {
    if (isTypingKey && relativeCaretPos === 0 && targetCellIndex > 0) {
     // log(`${logPrefix} Caret at forbidden position 0 (just after delimiter). Auto-advancing to position 1.`);
      const safePosAbs = cellStartCol + 1;
      editorInfo.ace_performSelectionChange([currentLineNum, safePosAbs], [currentLineNum, safePosAbs], false);
      editorInfo.ace_updateBrowserSelectionFromRep();
      relativeCaretPos = 1;
     // log(`${logPrefix} Caret moved to safe position. New relativeCaretPos=${relativeCaretPos}`);
    }
    const currentCol = cellStartCol + relativeCaretPos;
   // log(`${logPrefix} Handling INTERNAL key='${evt.key}' Type='${evt.type}' at Line=${currentLineNum}, Col=${currentCol} (CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}).`);
   // log(`${logPrefix} [caretTrace] Initial rep.selStart for internal edit: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

    if (evt.type !== 'keydown') {
       // log(`${logPrefix} Ignoring non-keydown event type ('${evt.type}') for handled key.`);
        return false; 
    }

   // log(`${logPrefix} Preventing default browser action for keydown event.`);
    evt.preventDefault();

    let newAbsoluteCaretCol = -1;
    let repBeforeEdit = null;

    try {
        repBeforeEdit = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

    if (isTypingKey) {
            const insertPos = [currentLineNum, currentCol];
           // log(`${logPrefix} -> Inserting text '${evt.key}' at [${insertPos}]`);
            editorInfo.ace_performDocumentReplaceRange(insertPos, insertPos, evt.key);
            newAbsoluteCaretCol = currentCol + 1;

        } else if (isInternalBackspace) {
            const delRangeStart = [currentLineNum, currentCol - 1];
            const delRangeEnd = [currentLineNum, currentCol];
           // log(`${logPrefix} -> Deleting (Backspace) range [${delRangeStart}]-[${delRangeEnd}]`);
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol - 1;

        } else if (isInternalDelete) {
            const delRangeStart = [currentLineNum, currentCol];
            const delRangeEnd = [currentLineNum, currentCol + 1];
           // log(`${logPrefix} -> Deleting (Delete) range [${delRangeStart}]-[${delRangeEnd}]`);
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol;
        }
        const repAfterReplace = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


       // log(`${logPrefix} -> Re-applying tbljson line attribute...`);

       // log(`${logPrefix} DEBUG: Before calculating attrStringToApply - trustedLastClick=${trustedLastClick}, reportedLineNum=${reportedLineNum}, currentLineNum=${currentLineNum}`);
       // log(`${logPrefix} DEBUG: lineAttrString value:`, lineAttrString ? `"${lineAttrString}"` : 'null/undefined');

        const applyHelper = editorInfo.ep_data_tables_applyMeta; 
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) { 
             const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;

            // log(`${logPrefix} DEBUG: Calculated attrStringToApply:`, attrStringToApply ? `"${attrStringToApply}"` : 'null/undefined');
            // log(`${logPrefix} DEBUG: Condition result: (${trustedLastClick} || ${reportedLineNum} === ${currentLineNum}) = ${trustedLastClick || reportedLineNum === currentLineNum}`);

             applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
            // log(`${logPrefix} -> tbljson line attribute re-applied (using rep before edit).`);
                } else {
             console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
             const currentRepFallback = editorInfo.ace_getRep();
             if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
                // log(`${logPrefix} -> Retrying attribute application with current rep...`);
                 applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
                // log(`${logPrefix} -> tbljson line attribute re-applied (using current rep fallback).`);
            } else {
                  console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute even with fallback rep.`);
             }
        }

        if (newAbsoluteCaretCol >= 0) {
             const newCaretPos = [currentLineNum, newAbsoluteCaretCol];
            // log(`${logPrefix} -> Setting selection immediately to:`, newCaretPos);
            // log(`${logPrefix} [caretTrace] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
             try {
                editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                const repAfterSelectionChange = editorInfo.ace_getRep();
               // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);
               // log(`${logPrefix} -> Selection set immediately.`);

                editorInfo.ace_fastIncorp(1); 
                const repAfterFastIncorp = editorInfo.ace_getRep();
               // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
               // log(`${logPrefix} -> Requested sync hint (fastIncorp 1).`);

                const targetCaretPosForReassert = [currentLineNum, newAbsoluteCaretCol];
               // log(`${logPrefix} [caretTrace] Attempting to re-assert selection post-fastIncorp to [${targetCaretPosForReassert[0]}, ${targetCaretPosForReassert[1]}]`);
                editorInfo.ace_performSelectionChange(targetCaretPosForReassert, targetCaretPosForReassert, false);
                const repAfterReassert = editorInfo.ace_getRep();
               // log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

                const newRelativePos = newAbsoluteCaretCol - cellStartCol;
                editor.ep_data_tables_last_clicked = {
                    lineNum: currentLineNum, 
                    tblId: metadataForTargetLine.tblId,
                    cellIndex: targetCellIndex,
                    relativePos: newRelativePos
                };
               // log(`${logPrefix} -> Updated stored click/caret info:`, editor.ep_data_tables_last_clicked);
               // log(`${logPrefix} [caretTrace] Updated ep_data_tables_last_clicked. Line=${editor.ep_data_tables_last_clicked.lineNum}, Cell=${editor.ep_data_tables_last_clicked.cellIndex}, RelPos=${editor.ep_data_tables_last_clicked.relativePos}`);


            } catch (selError) {
                 console.error(`${logPrefix} -> ERROR setting selection immediately:`, selError);
             }
        } else {
           // log(`${logPrefix} -> Warning: newAbsoluteCaretCol not set, skipping selection update.`);
            }

        } catch (error) {
       // log(`${logPrefix} ERROR during manual key handling:`, error);
            console.error('[ep_data_tables] Error processing key event update:', error);
    return true;
  }

    const endLogTime = Date.now();
   // log(`${logPrefix} END (Handled Internal Edit Manually) Key='${evt.key}' Type='${evt.type}' -> Returned true. Duration: ${endLogTime - startLogTime}ms`);
    return true;

  }


  const endLogTimeFinal = Date.now();
 // log(`${logPrefix} END (Fell Through / Unhandled Case) Key='${evt.key}' Type='${evt.type}'. Allowing default. Duration: ${endLogTimeFinal - startLogTime}ms`);
 // log(`${logPrefix} [caretTrace] Final rep.selStart at end of aceKeyEvent (if unhandled): Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);
  return false;
};
exports.aceInitialized = (h, ctx) => {
  const logPrefix = '[ep_data_tables:aceInitialized]';
 // log(`${logPrefix} START`, { hook_name: h, context: ctx });
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

 // log(`${logPrefix} Attaching ep_data_tables_applyMeta helper to editorInfo.`);
  ed.ep_data_tables_applyMeta = applyTableLineMetadataAttribute;
 // log(`${logPrefix}: Attached applyTableLineMetadataAttribute helper to ed.ep_data_tables_applyMeta successfully.`);

 // log(`${logPrefix} Storing documentAttributeManager reference on editorInfo.`);
  ed.ep_data_tables_docManager = docManager;
 // log(`${logPrefix}: Stored documentAttributeManager reference as ed.ep_data_tables_docManager.`);

 // log(`${logPrefix} Preparing to attach paste and resize listeners via ace_callWithAce.`);
  ed.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = '[ep_data_tables:aceInitialized:callWithAceForListeners]';
   // log(`${callWithAceLogPrefix} Entered ace_callWithAce callback for listeners.`);

    if (!ace || !ace.editor) {
      console.error(`${callWithAceLogPrefix} ERROR: ace or ace.editor is not available. Cannot attach listeners.`);
     // log(`${callWithAceLogPrefix} Aborting listener attachment due to missing ace.editor.`);
      return;
    }
    const editor = ace.editor;
   // log(`${callWithAceLogPrefix} ace.editor obtained successfully.`);

   // log(`${logPrefix} Storing editor reference on editorInfo.`);
    ed.ep_data_tables_editor = editor;
   // log(`${logPrefix}: Stored editor reference as ed.ep_data_tables_editor.`);

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
        // Only strip zero-width characters if not preserving them
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
      // This avoids drift caused by whitespace normalization during IME flows.
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

      // Get target cell index and identifiers from the actual DOM selection within the inner iframe.
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
            // Get all styled spans in this cell (skip delimiters, caret anchors, and nested image UI spans)
            const spans = td.querySelectorAll('span:not(.ep-data_tables-delim):not(.ep-data_tables-caret-anchor):not(.image-inner):not(.image-resize-handle)');
            let relPos = 0; // Position relative to cell start
            spans.forEach((span) => {
              totalSpansProcessed++;
              // Skip if this span is nested inside another span we're processing (e.g., image inner content)
              const parentSpan = span.parentElement?.closest?.('span[class*="image:"]');
              if (parentSpan && parentSpan !== span) {
                // This is a nested span inside an image - skip, we'll capture from the outer image span
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
              
              // For images, get direct text content only (exclude nested span content)
              // This prevents double-counting text from image-inner spans
              let text;
              if (isImageSpan) {
                // Get only direct text nodes, not nested span content
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
                  // Skip table-related classes
                  if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) continue;
                  if (cls === 'ace-line' || cls.startsWith('ep-data_tables-')) continue;
                  
                  // Parse class name to attribute key-value pair
                  // Etherpad uses: attributeName:value (colon) or attributeName-value (hyphen)
                  if (cls.includes(':')) {
                    // Colon separator: image:https%3A%2F%2F..., font-size:12, image-height:298px, etc.
                    const colonIdx = cls.indexOf(':');
                    stylingAttrs.push([cls.substring(0, colonIdx), cls.substring(colonIdx + 1)]);
                  } else if (cls.includes('-')) {
                    // Hyphen separator - need smart parsing based on known attribute prefixes
                    // Some are boolean flags that should be kept intact
                    const knownBooleanClasses = [
                      'inline-image', 'image-placeholder', 'image-inner', 'image-resize-handle'
                    ];
                    if (knownBooleanClasses.includes(cls)) {
                      stylingAttrs.push([cls, 'true']);
                    } else {
                      // Known multi-word attribute prefixes (split AFTER these)
                      const knownPrefixes = [
                        'image-id-',      // image-id-XXXX-XXXX-XXXX
                        'image-float-',   // image-float-none, image-float-left, etc.
                        'font-family-',   // font-family names can have hyphens
                      ];
                      let matched = false;
                      for (const prefix of knownPrefixes) {
                        if (cls.startsWith(prefix)) {
                          const attrName = prefix.slice(0, -1); // Remove trailing hyphen
                          const attrValue = cls.substring(prefix.length);
                          stylingAttrs.push([attrName, attrValue]);
                          matched = true;
                          console.debug('[ep_data_tables:extractStyling] prefix-match', { cls, prefix, attrName, attrValue });
                          break;
                        }
                      }
                      if (!matched) {
                        // Default: split at FIRST hyphen (standard Etherpad format)
                        // e.g., "author-abc123" → ["author", "abc123"]
                        const dashIdx = cls.indexOf('-');
                        const attrName = cls.substring(0, dashIdx);
                        const attrValue = cls.substring(dashIdx + 1);
                        stylingAttrs.push([attrName, attrValue]);
                        // Log if this looks like it might be wrong (multi-hyphen class)
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
                // Store DOM cell length to adjust for content length differences during reapplication
                // NOTE: td.textContent includes the delimiter span for cells 1+, so we subtract 1 for those
                let domCellLen = td.textContent?.replace(/\u00A0/g, ' ').length || 0;
                if (cellIdx > 0 && domCellLen > 0) {
                  domCellLen -= 1; // Exclude delimiter character from length
                }
                // Mark if this is an image for special handling during re-application
                const isImage = isImageSpan || stylingAttrs.some(([k]) => k === 'image' || k === 'inline');
                extractedStyling.push({ cellIdx, relStart: relPos, len: textLen, text, attrs: stylingAttrs, domCellLen, isImage });
              }
              relPos += textLen;
            });
          });
          // Log extraction summary for debugging (always log to help debug issues)
          const summary = extractedStyling.map(s => ({
            cell: s.cellIdx,
            text: s.text.slice(0, 12).replace(/[\u200B]/g, '⁰'), // Show zero-width as visible char
            attrs: s.attrs.map(a => a[0]).join(','),
            isImage: s.isImage,
          }));
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
        
        // Track successfully applied image positions per cell
        // This allows subsequent image styles to reuse the same position
        // (All image attributes belong to the same zero-width character)
        const imagePositionsPerCell = {};
        
        // Log cells array for debugging position calculation
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
            
            // STRATEGY: Use TEXT-BASED MATCHING for all cells
            // This isolates each cell from position changes in other cells.
            // For SHORT text that appears multiple times, use distance-from-end heuristic.
            // For LONG text, simple indexOf works (unlikely to have duplicates).
            
            // Helper: Find position using distance-from-end heuristic for ambiguous text
            const findWithOccurrenceHeuristic = (content, text, wasNearEnd) => {
              const first = content.indexOf(text);
              if (first === -1) return { pos: -1, type: 'not-found' };
              
              const second = content.indexOf(text, first + 1);
              if (second === -1) {
                // Unique occurrence
                return { pos: first, type: 'unique' };
              }
              
              // Multiple occurrences - choose based on original position
              if (wasNearEnd) {
                // Find last occurrence
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
            
            // Calculate if styled text was near the end of the original cell
            const domCellLen = style.domCellLen || originalCellContent.length;
            const distFromEnd = domCellLen - (originalRelStart + styledText.length);
            const wasNearEnd = distFromEnd <= 2;
            
            // Step 1: For SHORT text (< 4 chars), use occurrence-aware matching FIRST
            // This prevents indexOf from grabbing the wrong "22" in "222222222222"
            if (styledText.length < 4 && originalCellContent.length > 0) {
              const result = findWithOccurrenceHeuristic(originalCellContent, styledText, wasNearEnd);
              if (result.pos !== -1) {
                foundCellIdx = originalCellIdx;
                foundPos = result.pos;
                matchType = result.type === 'unique' ? 'unique-short' 
                          : result.type === 'last' ? 'last-occurrence' 
                          : 'first-occurrence';
              }
              
              // If not found, try case-insensitive for short text too
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
            
            // Step 2: For LONG text (4+ chars), use simple indexOf (duplicates unlikely)
            if (foundCellIdx === -1 && styledText.length >= 4 && originalCellContent) {
              foundPos = originalCellContent.indexOf(styledText);
              if (foundPos !== -1) {
                foundCellIdx = originalCellIdx;
                matchType = 'exact';
              }
              
              // Try case-insensitive for long text (handles Grammarly edits)
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
            
            // Step 3: If still not found, search all cells (content may have moved)
            if (foundCellIdx === -1) {
              for (let c = 0; c < cells.length; c++) {
                if (c === originalCellIdx) continue; // Already tried
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
              // Text not found - check if this is an image style
              // For image styles, reuse a previously found position in the same cell
              // (All image attributes belong to the same zero-width character)
              const isZeroWidthText = /^[\u200B\u200C\u200D\uFEFF]+$/.test(styledText);
              if (style.isImage || isZeroWidthText) {
                // Try to reuse a previously found image position in this cell
                if (imagePositionsPerCell[originalCellIdx] !== undefined) {
                  foundCellIdx = originalCellIdx;
                  foundPos = imagePositionsPerCell[originalCellIdx];
                  matchType = 'image-reuse';
                } else {
                  // Find the first zero-width character in the cell
                  const zwsMatch = originalCellContent.match(/[\u200B\u200C\u200D\uFEFF]/);
                  if (zwsMatch) {
                    foundCellIdx = originalCellIdx;
                    foundPos = originalCellContent.indexOf(zwsMatch[0]);
                    matchType = 'image-zws-fallback';
                    // Cache this position for other image styles
                    imagePositionsPerCell[originalCellIdx] = foundPos;
                  }
                }
              }
              
              // If still not found, skip this styling
              if (foundCellIdx === -1 || foundPos === -1) {
                console.debug('[ep_data_tables:reapplyStyling] skipped', {
                  originalCellIdx, text: styledText.slice(0, 10), reason: 'not-found',
                  isImage: style.isImage, isZeroWidth: isZeroWidthText,
                });
                continue;
              }
            }
            
            // Validate position is within bounds
            const foundCellContent = cells[foundCellIdx] || '';
            if (foundPos + styledText.length > foundCellContent.length) {
              console.debug('[ep_data_tables:reapplyStyling] skipped', {
                foundCellIdx, text: styledText.slice(0, 10), reason: 'out-of-bounds',
                foundPos, textLen: styledText.length, cellLen: foundCellContent.length,
              });
              continue;
            }
            
            // Calculate absolute position using ONLY canonical cell lengths
            // No DOM-based adjustments - each cell is independent!
            let absStart = 0;
            for (let c = 0; c < foundCellIdx; c++) {
              absStart += (cells[c]?.length || 0) + DELIMITER.length;
            }
            absStart += foundPos;
            
            const absEnd = absStart + styledText.length;
            
            // Calculate total line length to ensure we don't exceed bounds
            let totalLineLen = 0;
            for (let c = 0; c < cells.length; c++) {
              totalLineLen += (cells[c]?.length || 0);
              if (c < cells.length - 1) totalLineLen += DELIMITER.length;
            }
            
            // Skip if position would exceed line bounds - this can cause extra lines to be created
            if (absEnd > totalLineLen) {
              console.debug('[ep_data_tables:reapplyStyling] skipped', {
                foundCellIdx, text: styledText.slice(0, 10), reason: 'exceeds-line-bounds',
                absStart, absEnd, totalLineLen,
              });
              continue;
            }
            
            // Apply the styling
            if (style.attrs.length > 0) {
              // Cache this position for image styles so other image attrs can reuse it
              const isZeroWidthText = /^[\u200B\u200C\u200D\uFEFF]+$/.test(styledText);
              if ((style.isImage || isZeroWidthText) && imagePositionsPerCell[foundCellIdx] === undefined) {
                imagePositionsPerCell[foundCellIdx] = foundPos;
              }
              
              // Enhanced logging to trace position calculation
              const priorCellsDebug = [];
              let priorSum = 0;
              for (let c = 0; c < foundCellIdx; c++) {
                const cellLen = (cells[c]?.length || 0);
                priorCellsDebug.push(`c${c}:${cellLen}`);
                priorSum += cellLen + DELIMITER.length;
              }
              console.debug('[ep_data_tables:reapplyStyling] applying', {
                cell: foundCellIdx,
                foundPos,
                priorSum,
                absStart,
                absEnd,
                text: styledText.slice(0, 10),
                match: matchType,
                priorCells: priorCellsDebug.join('+'),
                attrs: style.attrs.map(a => `${a[0]}:${String(a[1]).slice(0,10)}`),
              });
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
          }
        } catch (err) {
          console.debug('[ep_data_tables:reapplyStyling] error (non-fatal)', err?.message);
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

            // Use pre-captured styling if available (from before collection), 
            // otherwise try to extract from current DOM (may be stale)
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

        // CRITICAL: Capture styling from DOM BEFORE any rep access or collection
        // After ace_getRep(), collection may run and wipe styling
        let preCapturedStyling = null;
        let preCapturedLineNum = -1;
        try {
          const innerDoc = ed.ep_data_tables_innerDoc || 
                          (typeof $inner !== 'undefined' && $inner[0]?.ownerDocument) || 
                          document;
          // Get current selection to determine line number before collection
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
                  if (preCapturedStyling.length > 0) {
                    // Log detailed info about what was captured
                    console.debug('[ep_data_tables:handleDesktopCommitInput] pre-captured styling', {
                      lineNum: i, 
                      count: preCapturedStyling.length,
                      details: preCapturedStyling.map(s => ({
                        cell: s.cellIdx,
                        text: s.text.slice(0, 20),
                        attrs: s.attrs.map(a => a[0]).join(','),
                      })),
                    });
                  }
                  break;
                }
              }
            }
          }
        } catch (preCapErr) {
          console.debug('[ep_data_tables:handleDesktopCommitInput] pre-capture failed', preCapErr?.message);
        }
        
        // Always log pre-capture result for debugging
        console.debug('[ep_data_tables:handleDesktopCommitInput] pre-capture result', {
          lineNum: preCapturedLineNum,
          stylingCount: preCapturedStyling?.length || 0,
          hasTable: preCapturedLineNum >= 0,
        });

        const rep = ed.ace_getRep();
        if (!rep || !rep.selStart) {
          console.debug('[ep_data_tables:handleDesktopCommitInput] early-exit: no rep or selStart');
          return;
        }
        let lineNum = rep.selStart[0];

        let tableMetadata = getTableMetadataForLine(lineNum);
        
        // If table metadata not found on current line, try pre-captured line
        // This handles cases where collection shifted line numbers
        if (!tableMetadata && preCapturedLineNum >= 0 && preCapturedLineNum !== lineNum) {
          console.debug('[ep_data_tables:handleDesktopCommitInput] trying pre-captured line', {
            currentLine: lineNum, preCapturedLine: preCapturedLineNum,
          });
          tableMetadata = getTableMetadataForLine(preCapturedLineNum);
          if (tableMetadata) {
            lineNum = preCapturedLineNum; // Use pre-captured line
          }
        }
        
        if (!tableMetadata) {
          console.debug('[ep_data_tables:handleDesktopCommitInput] early-exit: no table metadata', {
            lineNum, preCapturedLineNum, hasPreCapturedStyling: !!preCapturedStyling,
          });
          return;
        }

        const lineEntry = rep.lines.atIndex(lineNum);
        let sanitizedCells = collectSanitizedCells(lineEntry, tableMetadata, 'input-commit');
        const currentLineText = lineEntry?.text || '';
        let canonicalLine = sanitizedCells.join(DELIMITER);
        
        // Use pre-captured styling if available (line match check removed - we've already resolved lineNum)
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
          // Use PRE-CAPTURED styling from BEFORE collection ran (not stale lineEntry.lineNode)
          try {
            if (usablePreCapturedStyling && usablePreCapturedStyling.length > 0) {
              // CRITICAL: Get fresh line text INSIDE ace_callWithAce
              // The lineEntry.text we have is from BEFORE collection updated the document.
              // We need to re-read the line to get the actual content after collection.
              ed.ace_callWithAce((aceInstance) => {
                const freshRep = aceInstance.ace_getRep();
                const freshLineEntry = freshRep.lines.atIndex(lineNum);
                const freshLineText = freshLineEntry?.text || '';
                const actualCells = freshLineText.split(DELIMITER);
                reapplyStylingToLine(aceInstance, lineNum, actualCells, usablePreCapturedStyling);
                
                // Clean up any blank lines that may have been created after this table row
                // This can happen with Grammarly and other extensions
                try {
                  const repAfter = aceInstance.ace_getRep();
                  if (repAfter && repAfter.lines) {
                    const totalLines = repAfter.lines.length();
                    const blankLinesToRemove = [];
                    
                    // Scan for blank lines immediately after this row
                    for (let li = lineNum + 1; li < totalLines && li < lineNum + 10; li++) {
                      const checkEntry = repAfter.lines.atIndex(li);
                      const checkText = checkEntry?.text || '';
                      
                      // Stop at next table row
                      const isTableRow = docManager && typeof docManager.getAttributeOnLine === 'function'
                        && docManager.getAttributeOnLine(li, ATTR_TABLE_JSON);
                      if (isTableRow) break;
                      
                      // Check if blank (empty or just whitespace)
                      if (!checkText.trim() || checkText === '\n') {
                        blankLinesToRemove.push(li);
                      } else {
                        break; // Stop at first non-blank, non-table line
                      }
                    }
                    
                    // Remove blank lines bottom-up to preserve line numbers
                    if (blankLinesToRemove.length > 0) {
                      console.debug('[ep_data_tables:input-commit] cleaning up blank lines', {
                        lineNum, blankLines: blankLinesToRemove,
                      });
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
                } catch (cleanupErr) {
                  console.debug('[ep_data_tables:input-commit] blank line cleanup error (non-fatal)', cleanupErr?.message);
                }
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
        if (isAndroidUA() || isIOSUA()) return;
        try {
          const rep0 = ed.ace_getRep && ed.ace_getRep();
          if (rep0 && rep0.selStart) {
            let lineNum = rep0.selStart[0];
            let cellIndex = -1;
            let tableMetadata = null;
            let cellSnapshot = null;
            let originalTableLine = null; // Track where the table actually is
            
            try {
              // CRITICAL: First, try to find the table from DOM selection
              // This works even if the rep line number is wrong
              const domTarget = getDomCellTargetFromSelection();
              const domFoundTable = !!(domTarget && domTarget.tblId); // Track if DOM says we're in a table
              
              if (domFoundTable) {
                // Found a table from DOM selection - use its line number if available
                if (typeof domTarget.lineNum === 'number' && domTarget.lineNum >= 0) {
                  originalTableLine = domTarget.lineNum;
                  console.debug('[ep_data_tables:compositionstart] DOM selection found table', {
                    selLine: lineNum, domLine: domTarget.lineNum, tblId: domTarget.tblId, cellIdx: domTarget.idx,
                  });
                  // Use DOM-derived line number for metadata lookup
                  lineNum = domTarget.lineNum;
                  if (typeof domTarget.idx === 'number' && domTarget.idx >= 0) {
                    cellIndex = domTarget.idx;
                  }
                }
              }
              
              tableMetadata = getTableMetadataForLine(lineNum);
              
              // DOM fallback: if metadata not found via attribute, try extracting from table element
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
                        console.debug('[ep_data_tables:compositionstart] recovered metadata from DOM element', {
                          lineNum, tblId: domTblId, row: domRow, cols: domCols,
                        });
                      }
                    }
                  }
                } catch (domMetaErr) {
                  console.debug('[ep_data_tables:compositionstart] DOM metadata extraction failed', domMetaErr);
                }
              } else {
                originalTableLine = lineNum;
              }
              
              // NOTE: Nearby line scanning REMOVED - it was causing false positives
              // The previous logic would scan ±3 lines for tables even when NOT editing in a table,
              // which caused corruption when typing near (but not inside) a table.
              // 
              // If DOM selection indicates we're in a table (domFoundTable=true) but we couldn't
              // get metadata, that's a different issue that should be investigated separately.
              if (!tableMetadata && domFoundTable) {
                console.debug('[ep_data_tables:compositionstart] DOM found table but no metadata - unusual state', {
                  selLine: rep0.selStart[0], domLine: domTarget?.lineNum, tblId: domTarget?.tblId,
                });
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
                if (entry && tableMetadata.cols > 0) {
                  const lineNode = entry?.lineNode;
                  const tableEl = lineNode?.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
                  if (tableEl) {
                    const tr = tableEl.querySelector('tbody > tr');
                    if (tr && tr.children.length === tableMetadata.cols) {
                      cellSnapshot = new Array(tableMetadata.cols);
                      for (let i = 0; i < tableMetadata.cols; i++) {
                        const td = tr.children[i];
                        cellSnapshot[i] = sanitizeCellContent(td?.innerText || ' ');
                      }
                      console.debug('[ep_data_tables:compositionstart] snapshot captured from DOM', {
                        lineNum: originalTableLine ?? lineNum, cellIndex, cellSnapshot, tblId: tableMetadata.tblId,
                      });
                    }
                  }
                  // Fallback to text-based snapshot if DOM extraction failed
                  if (!cellSnapshot) {
                    const lineText = entry.text || '';
                    const cells = lineText.split(DELIMITER);
                    cellSnapshot = new Array(tableMetadata.cols);
                    for (let i = 0; i < tableMetadata.cols; i++) {
                      cellSnapshot[i] = sanitizeCellContent(cells[i] || ' ');
                    }
                    console.debug('[ep_data_tables:compositionstart] snapshot captured from text', {
                      lineNum: originalTableLine ?? lineNum, cellIndex, cellSnapshot, tblId: tableMetadata.tblId,
                    });
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
              lineNum,
              cellIndex,
              tblId: tableMetadata?.tblId ?? null,
              // Snapshot for corruption recovery
              snapshot: cellSnapshot,
              snapshotMeta: tableMetadata ? { ...tableMetadata } : null,
              // CRITICAL: Track the original table line for recovery
              originalTableLine: originalTableLine,
            };
            
            // CRITICAL: Set module-level variable for postWriteCanonicalize to access
            // This tells the canonicalizer where the table SHOULD be, not where the attribute ended up
            if (tableMetadata?.tblId && originalTableLine !== null) {
              __epDT_compositionOriginalLine = {
                tblId: tableMetadata.tblId,
                lineNum: originalTableLine,
                timestamp: Date.now(),
              };
            }
            
            // Log final state for debugging
            console.debug('[ep_data_tables:compositionstart] state captured', {
              selStart: rep0.selStart, lineNum, cellIndex,
              tblId: desktopComposition.tblId,
              originalTableLine,
              hasSnapshot: !!cellSnapshot,
              hasMeta: !!tableMetadata,
            });
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

      // IME/autocorrect diagnostics: capture-phase logging and newline soft-normalization for table lines
      // COMMENTED OUT FOR PRODUCTION - Uncomment for debugging IME/composition issues
      // const logIMEEvent = (rawEvt, tag) => {
      //   try {
      //     const e = rawEvt && (rawEvt.originalEvent || rawEvt);
      //     const rep = ed.ace_getRep && ed.ace_getRep();
      //     const selStart = rep && rep.selStart;
      //     const lineNum = selStart ? selStart[0] : -1;
      //     let isTableLine = false;
      //     if (lineNum >= 0) {
      //       let s = docManager && docManager.getAttributeOnLine ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON) : null;
      //       if (!s) {
      //         const meta = getTableLineMetadata(lineNum, ed, docManager);
      //         isTableLine = !!meta && typeof meta.cols === 'number';
      //       } else {
      //         isTableLine = true;
      //       }
      //     }
      //     if (!isTableLine) return;
      //     const payload = {
      //       tag,
      //       type: e && e.type,
      //       inputType: e && e.inputType,
      //       data: typeof (e && e.data) === 'string' ? e.data : null,
      //       isComposing: !!(e && e.isComposing),
      //       key: e && e.key,
      //       code: e && e.code,
      //       which: e && e.which,
      //       keyCode: e && e.keyCode,
      //     };
      //     console.debug('[ep_data_tables:ime-diag]', payload);
      //   } catch (_) {}
      // };

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

      try {
        const disableAuto = (el) => {
          if (!el) return;
          el.setAttribute('autocorrect', 'off');
          el.setAttribute('autocomplete', 'off');
          el.setAttribute('autocapitalize', 'off');
          el.setAttribute('spellcheck', 'false');
          // TEMPORARILY DISABLED FOR TESTING - uncomment to re-enable Grammarly blocking
          // el.setAttribute('data-gramm', 'false');
          // el.setAttribute('data-enable-grammarly', 'false');
        };
        disableAuto(innerDocBody[0] || innerDocBody);
        
        // TEMPORARILY DISABLED FOR TESTING - uncomment to re-enable Grammarly blocking on iframes
        // Also disable on the iframes themselves (Grammarly checks these too)
        // try {
        //   const outerFrame = document.querySelector('iframe[name="ace_outer"]');
        //   const innerFrame = outerFrame?.contentDocument?.querySelector('iframe[name="ace_inner"]');
        //   if (outerFrame) {
        //     outerFrame.setAttribute('data-gramm', 'false');
        //     outerFrame.setAttribute('data-enable-grammarly', 'false');
        //   }
        //   if (innerFrame) {
        //     innerFrame.setAttribute('data-gramm', 'false');
        //     innerFrame.setAttribute('data-enable-grammarly', 'false');
        //   }
        // } catch (_) {}
      } catch (_) {}
      
      if (!$inner || $inner.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: $inner is not valid. Cannot attach listeners.`);
        return;
      }

   // log(`${callWithAceLogPrefix} Attaching cut event listener to $inner (inner iframe body).`);
    $inner.on('cut', (evt) => {
      const cutLogPrefix = '[ep_data_tables:cutHandler]';
      console.log(`${cutLogPrefix} CUT EVENT TRIGGERED. Event object:`, evt);

      console.log(`${cutLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        console.warn(`${cutLogPrefix} WARNING: Could not get representation or selection. Allowing default cut.`);
        console.warn(`${cutLogPrefix} Could not get rep or selStart.`);
        return;
      }
      console.log(`${cutLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      console.log(`${cutLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);
      const hasSelectionInRep = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);
      if (!hasSelectionInRep) {
        console.log(`${cutLogPrefix} No selection detected in rep; deferring decision until table-line check.`);
      }

      if (selStart[0] !== selEnd[0]) {
        console.warn(`${cutLogPrefix} WARNING: Selection spans multiple lines. Preventing cut to protect table structure.`);
        evt.preventDefault();
        return;
      }

      console.log(`${cutLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
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
        console.log(`${cutLogPrefix} Line ${lineNum} is NOT a recognised table line. Allowing default cut.`);
        return;
      }

      console.log(`${cutLogPrefix} Line ${lineNum} IS a table line. Metadata:`, tableMetadata);

      if (!hasSelectionInRep) {
        console.log(`${cutLogPrefix} Preventing default CUT on table line with collapsed selection to protect delimiters.`);
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

      /* allow "…cell content + delimiter" selections */
      const wouldClampStart = targetCellIndex > 0 && selStart[1] === cellStartCol - DELIMITER.length;
      const wouldClampEnd = targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length;

      console.log(`[ep_data_tables:cut-handler] Cut selection analysis:`, {
        targetCellIndex,
        selStartCol: selStart[1],
        selEndCol: selEnd[1],
        cellStartCol,
        cellEndCol,
        delimiterLength: DELIMITER.length,
        expectedLeadingDelimiterPos: cellStartCol - DELIMITER.length,
        expectedTrailingDelimiterPos: cellEndCol + DELIMITER.length,
        wouldClampStart,
        wouldClampEnd
      });

      if (wouldClampStart) {
        console.log(`[ep_data_tables:cut-handler] CLAMPING cut selection start from ${selStart[1]} to ${cellStartCol}`);
        selStart[1] = cellStartCol;
      }

      if (wouldClampEnd) {
        console.log(`[ep_data_tables:cut-handler] CLAMPING cut selection end from ${selEnd[1]} to ${cellEndCol}`);
        selEnd[1] = cellEndCol;
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        console.warn(`${cutLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing cut to protect table structure.`);
        evt.preventDefault();
        return;
      }

      console.log(`${cutLogPrefix} Selection is entirely within cell ${targetCellIndex}. Intercepting cut to preserve table structure.`);
      evt.preventDefault();

      try {
        const selectedText = lineText.substring(selStart[1], selEnd[1]);
        console.log(`${cutLogPrefix} Selected text to cut: "${selectedText}"`);

        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(selectedText).then(() => {
            console.log(`${cutLogPrefix} Successfully copied to clipboard via Navigator API.`);
          }).catch((err) => {
            console.warn(`${cutLogPrefix} Failed to copy to clipboard via Navigator API:`, err);
          });
        } else {
          console.log(`${cutLogPrefix} Using fallback clipboard method.`);
          const textArea = document.createElement('textarea');
          textArea.value = selectedText;
          document.body.appendChild(textArea);
          textArea.select();
          try {
            document.execCommand('copy');
            console.log(`${cutLogPrefix} Successfully copied to clipboard via execCommand fallback.`);
          } catch (err) {
            console.warn(`${cutLogPrefix} Failed to copy to clipboard via fallback:`, err);
          }
          document.body.removeChild(textArea);
        }

        console.log(`${cutLogPrefix} Performing deletion via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${cutLogPrefix}[ace_callWithAceOps]`;
          console.log(`${callAceLogPrefix} Entered ace_callWithAce for cut operations. selStart:`, selStart, `selEnd:`, selEnd);

          console.log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to delete selected text.`);
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');
          console.log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
            console.log(`${callAceLogPrefix} Cell ${targetCellIndex} became empty after cut – inserting single space to preserve structure.`);
            const insertPos = [lineNum, selStart[1]];
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            const attrStart = insertPos;
            const attrEnd   = [insertPos[0], insertPos[1] + 1];
            aceInstance.ace_performDocumentApplyAttributesToRange(
              attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

          console.log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
          const repAfterCut = aceInstance.ace_getRep();
          console.log(`${callAceLogPrefix} Fetched rep after cut for applyMeta. Line ${lineNum} text now: "${repAfterCut.lines.atIndex(lineNum).text}"`);

          ed.ep_data_tables_applyMeta(
            lineNum,
            tableMetadata.tblId,
            tableMetadata.row,
            tableMetadata.cols,
            repAfterCut,
            ed,
            null,
            docManager
          );
          console.log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

          const newCaretPos = [lineNum, selStart[1]];
          console.log(`${callAceLogPrefix} Setting caret position to: [${newCaretPos}].`);
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
          console.log(`${callAceLogPrefix} Selection change successful.`);

          console.log(`${callAceLogPrefix} Cut operations within ace_callWithAce completed successfully.`);
        }, 'tableCutTextOperations', true);

        console.log(`${cutLogPrefix} Cut operation completed successfully.`);
      } catch (error) {
        console.error(`${cutLogPrefix} ERROR during cut operation:`, error);
        console.log(`${cutLogPrefix} Cut operation failed. Error details:`, { message: error.message, stack: error.stack });
      }
    });

   // log(`${callWithAceLogPrefix} Attaching beforeinput event listener to $inner (inner iframe body).`);
    $inner.on('beforeinput', (evt) => {
      const deleteLogPrefix = '[ep_data_tables:beforeinputDeleteHandler]';
     // log(`${deleteLogPrefix} BEFOREINPUT EVENT TRIGGERED. inputType: "${evt.originalEvent.inputType}", event object:`, evt);

      if (!evt.originalEvent.inputType || !evt.originalEvent.inputType.startsWith('delete')) {
       // log(`${deleteLogPrefix} Not a deletion event (inputType: "${evt.originalEvent.inputType}"). Allowing default.`);
        return;
      }

     // log(`${deleteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
       // log(`${deleteLogPrefix} WARNING: Could not get representation or selection. Allowing default delete.`);
        console.warn(`${deleteLogPrefix} Could not get rep or selStart.`);
        return;
      }
     // log(`${deleteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
     // log(`${deleteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

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
       // log(`${deleteLogPrefix} WARNING: Selection spans multiple lines. Preventing delete to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${deleteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
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
       // log(`${deleteLogPrefix} Line ${lineNum} is NOT a recognised table line. Allowing default delete.`);
        return;
      }

     // log(`${deleteLogPrefix} Line ${lineNum} IS a table line. Metadata:`, tableMetadata);

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
      const wouldClampStart = targetCellIndex > 0 && selStart[1] === cellStartCol - DELIMITER.length;
      const wouldClampEnd = targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length;

      console.log(`[ep_data_tables:beforeinput-delete] Delete selection analysis:`, {
        targetCellIndex,
        selStartCol: selStart[1],
        selEndCol: selEnd[1],
        cellStartCol,
        cellEndCol,
        delimiterLength: DELIMITER.length,
        expectedLeadingDelimiterPos: cellStartCol - DELIMITER.length,
        expectedTrailingDelimiterPos: cellEndCol + DELIMITER.length,
        wouldClampStart,
        wouldClampEnd
      });

      if (wouldClampStart) {
        console.log(`[ep_data_tables:beforeinput-delete] CLAMPING delete selection start from ${selStart[1]} to ${cellStartCol}`);
        selStart[1] = cellStartCol;
      }

      if (wouldClampEnd) {
        console.log(`[ep_data_tables:beforeinput-delete] CLAMPING delete selection end from ${selEnd[1]} to ${cellEndCol}`);
        selEnd[1] = cellEndCol;
      }

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
       // log(`${deleteLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing delete to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${deleteLogPrefix} Selection is entirely within cell ${targetCellIndex}. Intercepting delete to preserve table structure.`);
      evt.preventDefault();

      try {
       // log(`${deleteLogPrefix} Performing deletion via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${deleteLogPrefix}[ace_callWithAceOps]`;
         // log(`${callAceLogPrefix} Entered ace_callWithAce for delete operations. selStart:`, selStart, `selEnd:`, selEnd);

         // log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to delete selected text.`);
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');
         // log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
           // log(`${callAceLogPrefix} Cell ${targetCellIndex} became empty after delete – inserting single space to preserve structure.`);
            const insertPos = [lineNum, selStart[1]];
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            const attrStart = insertPos;
            const attrEnd   = [insertPos[0], insertPos[1] + 1];
            aceInstance.ace_performDocumentApplyAttributesToRange(
              attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

         // log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
          const repAfterDelete = aceInstance.ace_getRep();
         // log(`${callAceLogPrefix} Fetched rep after delete for applyMeta. Line ${lineNum} text now: "${repAfterDelete.lines.atIndex(lineNum).text}"`);

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
         // log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

          const newCaretAbsoluteCol = (cellTextAfterDeletion.length === 0) ? selStart[1] + 1 : selStart[1];
          const newCaretPos = [lineNum, newCaretAbsoluteCol];
         // log(`${callAceLogPrefix} Setting caret position to: [${newCaretPos}].`);
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
         // log(`${callAceLogPrefix} Selection change successful.`);

         // log(`${callAceLogPrefix} Delete operations within ace_callWithAce completed successfully.`);
        }, 'tableDeleteTextOperations', true);

       // log(`${deleteLogPrefix} Delete operation completed successfully.`);
      } catch (error) {
        console.error(`${deleteLogPrefix} ERROR during delete operation:`, error);
       // log(`${deleteLogPrefix} Delete operation failed. Error details:`, { message: error.message, stack: error.stack });
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
      if (isAndroidChromeComposition) {
        logCompositionEvent('compositionend-android-handler', evt);
        isAndroidChromeComposition = false;
        handledCurrentComposition = false;
        suppressBeforeInputInsertTextDuringComposition = false;
      }
    });

    $inner.on('compositionend', (evt) => {
      if (isAndroidUA() || isIOSUA()) return;
      const compLogPrefix = '[ep_data_tables:compositionEndDesktop]';
      const nativeEvt = evt.originalEvent || evt;
      const dataPreview = typeof nativeEvt?.data === 'string' ? nativeEvt.data : '';
      logCompositionEvent('compositionend-desktop-fired', evt, { data: dataPreview });

        // Prevent the immediate post-composition input commit from running; we pipeline instead
        suppressNextInputCommit = true;
      requestAnimationFrame(() => {
        try {
          ed.ace_callWithAce((aceInstance) => {
            // CRITICAL GUARD: Only run table-related pipeline if we're actually editing a table
            // Check if compositionstart found a table (snapshotMeta set) or if DOM selection is in a table
            const domTargetCheck = getDomCellTargetFromSelection();
            const domIsInTable = !!(domTargetCheck && domTargetCheck.tblId);
            const compositionFoundTable = !!(desktopComposition && desktopComposition.snapshotMeta);
            
            if (!compositionFoundTable && !domIsInTable) {
              // User is NOT editing in a table - let normal Etherpad handle this
              logCompositionEvent('compositionend-desktop-skipped-not-in-table', evt, {
                compositionFoundTable,
                domIsInTable,
                lineNum: desktopComposition?.lineNum,
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
              if (!caret && (desktopComposition && typeof desktopComposition.lineNum !== 'number')) return;
              // Prefer the line captured at compositionstart to avoid caret drift.
              const pipelineLineNum = (desktopComposition && typeof desktopComposition.lineNum === 'number')
                ? desktopComposition.lineNum
                : caret[0];
            let metadata = getTableMetadataForLine(pipelineLineNum);
              let entry = repNow.lines.atIndex(pipelineLineNum);
            if (!entry) {
              logCompositionEvent('compositionend-desktop-no-line-entry', evt, { lineNum: pipelineLineNum });
              return;
            }
              // If tblId was captured at start, and current metadata doesn't match, relocate the line by tblId.
              if (desktopComposition && desktopComposition.tblId) {
                const currentTblId = metadata && metadata.tblId ? metadata.tblId : null;
                if (currentTblId !== desktopComposition.tblId) {
                  const relocatedLine = findLineNumByTblId(desktopComposition.tblId);
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
              let idx = (desktopComposition && desktopComposition.cellIndex >= 0)
                ? desktopComposition.cellIndex
                : (() => {
                    const selCol = (desktopComposition && desktopComposition.start) ? desktopComposition.start[1] : (caret ? caret[1] : 0);
                    const rawMap = computeTargetCellIndexFromRaw(entry, selCol);
                    return rawMap.index;
                  })();
            if (idx < 0) idx = Math.min(metadata.cols - 1, 0);

            // Compute relative selection in cell
              let baseOffset = 0;
              for (let i = 0; i < idx; i++) baseOffset += (cellsNow[i]?.length ?? 0) + DELIMITER.length;
              const sColAbs = (desktopComposition && desktopComposition.start) ? desktopComposition.start[1] : (caret ? caret[1] : 0);
              const eColAbs = (desktopComposition && desktopComposition.end) ? desktopComposition.end[1] : sColAbs;
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

            // Post-composition orphan detection and repair using snapshot
            // This handles cases where the composition caused line fragmentation
            setTimeout(() => {
              // CRITICAL: Pre-check editor state before calling ace_callWithAce
              // The ace_callWithAce itself can crash if Etherpad's internal state is corrupted
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
                  
                  const snapshotMeta = desktopComposition.snapshotMeta;
                  const snapshotCells = desktopComposition.snapshot;
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
                          // This catches orphans where browser stripped tbljson-* but kept tblCell-*
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
                    
                    // CRITICAL: Find the primary line using LIVE DOM query
                    // The lineNode references from rep.lines might be stale - query the actual DOM
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
                    
                    // Delete orphan lines (bottom-up to preserve line numbers)
                    // Content has already been merged, so this is safe
                    // CRITICAL: Wrap each operation in try-catch to prevent cascade failures
                    orphans.sort((a, b) => b.lineNum - a.lineNum);
                    let orphansRemoved = 0;
                    for (const orphan of orphans) {
                      try {
                        // Re-fetch rep to get current line count (may have changed after previous deletions)
                        const repCheck = ace2.ace_getRep();
                        if (!repCheck || !repCheck.lines) {
                          console.warn('[ep_data_tables:compositionend-orphan-repair] rep invalid, aborting', { orphan: orphan.lineNum });
                          break;
                        }
                        const currentLineCount = repCheck.lines.length();
                        if (orphan.lineNum >= currentLineCount) {
                          console.debug('[ep_data_tables:compositionend-orphan-repair] skipping orphan - line no longer exists', {
                            orphanLine: orphan.lineNum, currentLineCount,
                          });
                          continue;
                        }
                        // Verify the line entry exists before operating on it
                        const orphanEntry = repCheck.lines.atIndex(orphan.lineNum);
                        if (!orphanEntry) {
                          console.debug('[ep_data_tables:compositionend-orphan-repair] skipping orphan - no entry', {
                            orphanLine: orphan.lineNum,
                          });
                          continue;
                        }
                        
                      try {
                        if (docManager && typeof docManager.removeAttributeOnLine === 'function') {
                          docManager.removeAttributeOnLine(orphan.lineNum, ATTR_TABLE_JSON);
                        }
                        } catch (attrErr) {
                          console.debug('[ep_data_tables:compositionend-orphan-repair] removeAttribute failed', attrErr);
                        }
                        
                      console.debug('[ep_data_tables:compositionend-orphan-repair] removing orphan (merged)', {
                        orphanLine: orphan.lineNum,
                      });
                      ace2.ace_performDocumentReplaceRange([orphan.lineNum, 0], [orphan.lineNum + 1, 0], '');
                        orphansRemoved++;
                      } catch (orphanDeleteErr) {
                        console.error('[ep_data_tables:compositionend-orphan-repair] error deleting orphan line', {
                          orphanLine: orphan.lineNum,
                          error: orphanDeleteErr?.message || orphanDeleteErr,
                        });
                        // Don't break - try to continue with remaining orphans
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

   // log(`${callWithAceLogPrefix} Attaching drag and drop event listeners to $inner (inner iframe body).`);

    $inner.on('drop', (evt) => {
      const dropLogPrefix = '[ep_data_tables:dropHandler]';
     // log(`${dropLogPrefix} DROP EVENT TRIGGERED. Event object:`, evt);

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

     // log(`${dropLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
       // log(`${dropLogPrefix} WARNING: Could not get representation or selection. Allowing default drop.`);
        return;
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];
     // log(`${dropLogPrefix} Current line number: ${lineNum}.`);

     // log(`${dropLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let isTableLine = !!lineAttrString;

      if (!isTableLine) {
        const metadataFallback = getTableLineMetadata(lineNum, ed, docManager);
        isTableLine = !!metadataFallback;
      }

      if (isTableLine) {
     // log(`${dropLogPrefix} Line ${lineNum} IS a table line. Preventing drop to protect table structure.`);
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

     // log(`${dragLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let isTableLine = !!lineAttrString;

      if (!isTableLine) {
        isTableLine = !!getTableLineMetadata(lineNum, ed, docManager);
      }

      if (isTableLine) {
       // log(`${dragLogPrefix} Preventing dragover on table line ${lineNum} to control drop handling.`);
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

   // log(`${callWithAceLogPrefix} Attaching paste event listener to $inner (inner iframe body).`);
    $inner.on('paste', (evt) => {
      const pasteLogPrefix = '[ep_data_tables:pasteHandler]';
     // log(`${pasteLogPrefix} PASTE EVENT TRIGGERED. Event object:`, evt);

     // log(`${pasteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
       // log(`${pasteLogPrefix} WARNING: Could not get representation or selection. Allowing default paste.`);
        console.warn(`${pasteLogPrefix} Could not get rep or selStart.`);
        return;
      }
     // log(`${pasteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
     // log(`${pasteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      if (selStart[0] !== selEnd[0]) {
       // log(`${pasteLogPrefix} WARNING: Selection spans multiple lines. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${pasteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (!lineAttrString) {
       // log(`${pasteLogPrefix} No '${ATTR_TABLE_JSON}' attribute found. Checking if this is a block-styled table row via DOM reconstruction.`);
        const fallbackMeta = getTableLineMetadata(lineNum, ed, docManager);
        if (fallbackMeta) {
          tableMetadata = fallbackMeta;
          lineAttrString = JSON.stringify(fallbackMeta);
         // log(`${pasteLogPrefix} Block-styled table row detected. Reconstructed metadata:`, fallbackMeta);
        }
      }

      if (!lineAttrString) {
       // log(`${pasteLogPrefix} Line ${lineNum} is NOT a table line (no '${ATTR_TABLE_JSON}' attribute found and no DOM reconstruction possible). Allowing default paste.`);
        return;
      }
     // log(`${pasteLogPrefix} Line ${lineNum} IS a table line. Attribute string: "${lineAttrString}".`);

      try {
       // log(`${pasteLogPrefix} Parsing table metadata from attribute string.`);
        if (!tableMetadata) {
          tableMetadata = JSON.parse(lineAttrString);
        }
       // log(`${pasteLogPrefix} Parsed table metadata:`, tableMetadata);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
         // log(`${pasteLogPrefix} WARNING: Invalid or incomplete table metadata on line ${lineNum}. Allowing default paste. Metadata:`, tableMetadata);
          console.warn(`${pasteLogPrefix} Invalid table metadata for line ${lineNum}.`);
          return;
        }
       // log(`${pasteLogPrefix} Table metadata validated successfully: tblId=${tableMetadata.tblId}, row=${tableMetadata.row}, cols=${tableMetadata.cols}.`);
      } catch(e) {
        console.error(`${pasteLogPrefix} ERROR parsing table metadata for line ${lineNum}:`, e);
       // log(`${pasteLogPrefix} Metadata parse error. Allowing default paste. Error details:`, { message: e.message, stack: e.stack });
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
       // log(`${pasteLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${pasteLogPrefix} Accessing clipboard data.`);
      const clipboardData = evt.originalEvent.clipboardData || window.clipboardData;
      if (!clipboardData) {
       // log(`${pasteLogPrefix} WARNING: No clipboard data found. Allowing default paste.`);
        return;
      }
     // log(`${pasteLogPrefix} Clipboard data object obtained:`, clipboardData);

      const types = clipboardData.types || [];
      if (types.includes('text/html') && clipboardData.getData('text/html')) {
       // log(`${pasteLogPrefix} Detected text/html in clipboard – deferring to other plugins and default paste.`);
        return;
      }

     // log(`${pasteLogPrefix} Getting 'text/plain' from clipboard.`);
      const pastedTextRaw = clipboardData.getData('text/plain');
     // log(`${pasteLogPrefix} Pasted text raw: "${pastedTextRaw}" (Type: ${typeof pastedTextRaw})`);

      let pastedText = pastedTextRaw
        .replace(/(\r\n|\n|\r)/gm, " ")
        .replace(new RegExp(DELIMITER, 'g'), ' ')
        .replace(/\t/g, " ")
        .replace(/\s+/g, " ")
        .trim();

     // log(`${pasteLogPrefix} Pasted text after sanitization: "${pastedText}"`);

      if (typeof pastedText !== 'string' || pastedText.length === 0) {
       // log(`${pasteLogPrefix} No plain text in clipboard or text is empty (after sanitization). Allowing default paste.`);
        const types = clipboardData.types;
       // log(`${pasteLogPrefix} Clipboard types available:`, types);
        if (types && types.includes('text/html')) {
           // log(`${pasteLogPrefix} Clipboard also contains HTML:`, clipboardData.getData('text/html'));
        }
        return;
      }
     // log(`${pasteLogPrefix} Plain text obtained from clipboard: "${pastedText}". Length: ${pastedText.length}.`);

      const currentCellText = cells[targetCellIndex] || '';
      const selectionLength = selEnd[1] - selStart[1];
      const newCellLength = currentCellText.length - selectionLength + pastedText.length;

      const MAX_CELL_LENGTH = 8000;
      if (newCellLength > MAX_CELL_LENGTH) {
       // log(`${pasteLogPrefix} WARNING: Paste would exceed maximum cell length (${newCellLength} > ${MAX_CELL_LENGTH}). Truncating paste.`);
        const truncatedPaste = pastedText.substring(0, MAX_CELL_LENGTH - (currentCellText.length - selectionLength));
        if (truncatedPaste.length === 0) {
         // log(`${pasteLogPrefix} Paste would be completely truncated. Preventing paste.`);
          evt.preventDefault();
          return;
        }
       // log(`${pasteLogPrefix} Using truncated paste: "${truncatedPaste}"`);
        pastedText = truncatedPaste;
      }

     // log(`${pasteLogPrefix} INTERCEPTING paste of plain text into table line ${lineNum}. PREVENTING DEFAULT browser action.`);
      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
       // log(`${pasteLogPrefix} Preparing to perform paste operations via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
            const callAceLogPrefix = `${pasteLogPrefix}[ace_callWithAceOps]`;
           // log(`${callAceLogPrefix} Entered ace_callWithAce for paste operations. selStart:`, selStart, `selEnd:`, selEnd);

           // log(`${callAceLogPrefix} Original line text from initial rep: "${rep.lines.atIndex(lineNum).text}". SelStartCol: ${selStart[1]}, SelEndCol: ${selEnd[1]}.`);

           // log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to insert text: "${pastedText}".`);
            aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, pastedText);
           // log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

           // log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
            const repAfterReplace = aceInstance.ace_getRep();
           // log(`${callAceLogPrefix} Fetched rep after replace for applyMeta. Line ${lineNum} text now: "${repAfterReplace.lines.atIndex(lineNum).text}"`);

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
           // log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

            const newCaretCol = selStart[1] + pastedText.length;
            const newCaretPos = [lineNum, newCaretCol];
           // log(`${callAceLogPrefix} New calculated caret position: [${newCaretPos}]. Setting selection.`);
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
           // log(`${callAceLogPrefix} Selection change successful.`);

           // log(`${callAceLogPrefix} Requesting fastIncorp(10) for sync.`);
            aceInstance.ace_fastIncorp(10);
           // log(`${callAceLogPrefix} fastIncorp requested.`);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
               const newRelativePos = newCaretCol - cellStartCol;
               editor.ep_data_tables_last_clicked = {
                  lineNum: lineNum,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
               };
              // log(`${callAceLogPrefix} Updated stored click/caret info:`, editor.ep_data_tables_last_clicked);
            }

           // log(`${callAceLogPrefix} Paste operations within ace_callWithAce completed successfully.`);
        }, 'tablePasteTextOperations', true);
       // log(`${pasteLogPrefix} ed.ace_callWithAce for paste operations was called.`);

      } catch (error) {
        console.error(`${pasteLogPrefix} CRITICAL ERROR during paste handling operation:`, error);
       // log(`${pasteLogPrefix} Error details:`, { message: error.message, stack: error.stack });
       // log(`${pasteLogPrefix} Paste handling FAILED. END OF HANDLER.`);
      }
    });
   // log(`${callWithAceLogPrefix} Paste event listener attached.`);

   // log(`${callWithAceLogPrefix} Attaching column resize listeners...`);

    const $iframeOuter = $('iframe[name="ace_outer"]');
    const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
    const innerDoc = $iframeInner.contents();
    const outerDoc = $iframeOuter.contents();

   // log(`${callWithAceLogPrefix} Found iframe documents: outer=${outerDoc.length}, inner=${innerDoc.length}`);

    $inner.on('mousedown', '.ep-data_tables-resize-handle', (evt) => {
      const resizeLogPrefix = '[ep_data_tables:resizeMousedown]';
     // log(`${resizeLogPrefix} Resize handle mousedown detected`);

      if (evt.button !== 0) {
       // log(`${resizeLogPrefix} Ignoring non-left mouse button: ${evt.button}`);
        return;
      }

      const target = evt.target;
      const $target = $(target);
      const isImageRelated = $target.closest('.inline-image, .image-placeholder, .image-inner').length > 0;
      const isImageResizeHandle = $target.hasClass('image-resize-handle') || $target.closest('.image-resize-handle').length > 0;

      if (isImageRelated || isImageResizeHandle) {
       // log(`${resizeLogPrefix} Click detected on image-related element or image resize handle, ignoring for table resize`);
        return;
      }

      evt.preventDefault();
      evt.stopPropagation();

      const handle = evt.target;
      const columnIndex = parseInt(handle.getAttribute('data-column'), 10);
      const table = handle.closest('table.dataTable');
      const lineNode = table.closest('div.ace-line');

     // log(`${resizeLogPrefix} Parsed resize target: columnIndex=${columnIndex}, table=${!!table}, lineNode=${!!lineNode}`);

      if (table && lineNode && !isNaN(columnIndex)) {
        const tblId = table.getAttribute('data-tblId');
        const rep = ed.ace_getRep();

        if (!rep || !rep.lines) {
          console.error(`${resizeLogPrefix} Cannot get editor representation`);
          return;
        }

        const lineNum = rep.lines.indexOfKey(lineNode.id);

       // log(`${resizeLogPrefix} Table info: tblId=${tblId}, lineNum=${lineNum}`);

        if (tblId && lineNum !== -1) {
          try {
            const lineAttrString = docManager.getAttributeOnLine(lineNum, 'tbljson');
            if (lineAttrString) {
              const metadata = JSON.parse(lineAttrString);
              if (metadata.tblId === tblId) {
               // log(`${resizeLogPrefix} Starting resize with metadata:`, metadata);
                startColumnResize(table, columnIndex, evt.clientX, metadata, lineNum);
               // log(`${resizeLogPrefix} Started resize for column ${columnIndex}`);

               // log(`${resizeLogPrefix} Global resize state: isResizing=${isResizing}, targetTable=${!!resizeTargetTable}, targetColumn=${resizeTargetColumn}`);
              } else {
               // log(`${resizeLogPrefix} Table ID mismatch: ${metadata.tblId} vs ${tblId}`);
              }
            } else {
             // log(`${resizeLogPrefix} No table metadata found for line ${lineNum}, trying DOM reconstruction...`);

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
                       // log(`${resizeLogPrefix} Reconstructed metadata from DOM:`, reconstructedMetadata);

                        startColumnResize(table, columnIndex, evt.clientX, reconstructedMetadata, lineNum);
                       // log(`${resizeLogPrefix} Started resize for column ${columnIndex} using reconstructed metadata`);

                       // log(`${resizeLogPrefix} Global resize state: isResizing=${isResizing}, targetTable=${!!resizeTargetTable}, targetColumn=${resizeTargetColumn}`);
    } else {
                       // log(`${resizeLogPrefix} DOM table found but no cells detected`);
                      }
                    } else {
                     // log(`${resizeLogPrefix} DOM table found but tblId mismatch or missing row: domTblId=${domTblId}, domRow=${domRow}`);
                    }
                  } else {
                   // log(`${resizeLogPrefix} No table found in DOM for line ${lineNum}`);
                  }
                } else {
                 // log(`${resizeLogPrefix} Could not get line entry or lineNode for line ${lineNum}`);
                }
              } else {
               // log(`${resizeLogPrefix} Could not get rep or rep.lines for DOM reconstruction`);
              }
            }
          } catch (e) {
            console.error(`${resizeLogPrefix} Error getting table metadata:`, e);
          }
        } else {
         // log(`${resizeLogPrefix} Invalid line number (${lineNum}) or table ID (${tblId})`);
        }
      } else {
       // log(`${resizeLogPrefix} Invalid resize target:`, { table: !!table, lineNode: !!lineNode, columnIndex });
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
       // log(`${mouseupLogPrefix} Mouseup detected on ${evt.target.tagName || 'unknown'}. isResizing: ${isResizing}`);

        if (isResizing) {
         // log(`${mouseupLogPrefix} Processing resize completion...`);
          evt.preventDefault();
          evt.stopPropagation();

          setTimeout(() => {
           // log(`${mouseupLogPrefix} Executing finishColumnResize after delay...`);
            finishColumnResize(ed, docManager);
           // log(`${mouseupLogPrefix} Resize completion finished.`);
          }, 50);
        } else {
         // log(`${mouseupLogPrefix} Not in resize mode, ignoring mouseup.`);
        }
      };

     // log(`${callWithAceLogPrefix} Attaching global mousemove/mouseup handlers to multiple contexts...`);

      $(document).on('mousemove', handleMousemove);
      $(document).on('mouseup', handleMouseup);
     // log(`${callWithAceLogPrefix} Attached to main document`);

      if (outerDoc.length > 0) {
        outerDoc.on('mousemove', handleMousemove);
        outerDoc.on('mouseup', handleMouseup);
       // log(`${callWithAceLogPrefix} Attached to outer iframe document`);
      }

      if (innerDoc.length > 0) {
        innerDoc.on('mousemove', handleMousemove);
        innerDoc.on('mouseup', handleMouseup);
       // log(`${callWithAceLogPrefix} Attached to inner iframe document`);
      }

      $inner.on('mousemove', handleMousemove);
      $inner.on('mouseup', handleMouseup);
     // log(`${callWithAceLogPrefix} Attached to inner iframe body`);

      const failsafeMouseup = (evt) => {
        if (isResizing) {
         // log(`${mouseupLogPrefix} FAILSAFE: Detected mouse event during resize: ${evt.type}`);
          if (evt.type === 'mouseup' || evt.type === 'mousedown' || evt.type === 'click') {
           // log(`${mouseupLogPrefix} FAILSAFE: Triggering resize completion due to ${evt.type}`);
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
     // log(`${callWithAceLogPrefix} Attached failsafe event handlers`);

      const preventTableDrag = (evt) => {
        const target = evt.target;
        const inTable = target && typeof target.closest === 'function' && target.closest('table.dataTable');
        if (inTable) {
         // log('[ep_data_tables:dragPrevention] Preventing drag operation originating from inside table');
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
     // log(`${callWithAceLogPrefix} Attached drag prevention handlers to inner body`);

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

   // log(`${callWithAceLogPrefix} Column resize listeners attached successfully.`);
      } catch (e) {
        console.error(`${callWithAceLogPrefix} ERROR: Exception while attaching listeners:`, e);
      }
    }; // End of attachListeners function

    // Start the retry process to access iframes and attach all listeners
    tryGetIframeBody(0);

  }, 'tablePasteAndResizeListeners', true);
 // log(`${logPrefix} ace_callWithAce for listeners setup completed.`);

  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo, attributeString = null, documentAttributeManager = null) {
    const funcName = 'applyTableLineMetadataAttribute';
   // log(`${logPrefix}:${funcName}: START - Applying METADATA attribute to line ${lineNum}`, {tblId, rowIndex, numCols});

    let finalMetadata;

    if (attributeString) {
      try {
        const providedMetadata = JSON.parse(attributeString);
        if (providedMetadata.columnWidths && Array.isArray(providedMetadata.columnWidths) && providedMetadata.columnWidths.length === numCols) {
          finalMetadata = providedMetadata;
         // log(`${logPrefix}:${funcName}: Using provided metadata with existing columnWidths`);
        } else {
          finalMetadata = providedMetadata;
         // log(`${logPrefix}:${funcName}: Provided metadata missing columnWidths, attempting DOM extraction`);
           }
         } catch (e) {
       // log(`${logPrefix}:${funcName}: Error parsing provided attributeString, will reconstruct:`, e);
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
               // log(`${logPrefix}:${funcName}: Extracted column widths from DOM: ${columnWidths.map(w => w.toFixed(1) + '%').join(', ')}`);
              }
            }
          }
             }
           } catch (e) {
       // log(`${logPrefix}:${funcName}: Error extracting column widths from DOM:`, e);
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
   // log(`${logPrefix}:${funcName}: Final metadata attribute string: ${finalAttributeString}`);

    try {
       const lineEntry = rep.lines.atIndex(lineNum);
       if (!lineEntry) {
          // log(`${logPrefix}:${funcName}: ERROR - Could not find line entry for line number ${lineNum}`);
           return;
       }
       const lineLength = Math.max(1, lineEntry.text.length);
      // log(`${logPrefix}:${funcName}: Line ${lineNum} text length: ${lineLength}`);

       const attributes = [[ATTR_TABLE_JSON, finalAttributeString]];
       const start = [lineNum, 0];
       const end = [lineNum, lineLength];

      // log(`${logPrefix}:${funcName}: Applying tbljson attribute to range [${start}]-[${end}]`);
       editorInfo.ace_performDocumentApplyAttributesToRange(start, end, attributes);
      // log(`${logPrefix}:${funcName}: Successfully applied tbljson attribute to line ${lineNum}`);

    } catch(e) {
        console.error(`[ep_data_tables] ${logPrefix}:${funcName}: Error applying metadata attribute on line ${lineNum}:`, e);
    }
  }

  /** Insert a fresh rows×cols blank table at the caret */
  ed.ace_createTableViaAttributes = (rows = 2, cols = 2) => {
    const funcName = 'ace_createTableViaAttributes';
   // log(`${funcName}: START - Refactored Phase 4 (Get Selection Fix)`, { rows, cols });
    rows = Math.max(1, rows); cols = Math.max(1, cols);
   // log(`${funcName}: Ensuring minimum 1 row, 1 col.`);

    const tblId   = rand();
   // log(`${funcName}: Generated table ID: ${tblId}`);
    const initialCellContent = ' ';
    const lineTxt = Array.from({ length: cols }).fill(initialCellContent).join(DELIMITER);
   // log(`${funcName}: Constructed initial line text for ${cols} cols: "${lineTxt}"`);
    const block = Array.from({ length: rows }).fill(lineTxt).join('\n') + '\n';
   // log(`${funcName}: Constructed block for ${rows} rows:\n${block}`);

   // log(`${funcName}: Getting current representation and selection...`);
    const currentRepInitial = ed.ace_getRep(); 
    if (!currentRepInitial || !currentRepInitial.selStart || !currentRepInitial.selEnd) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation or selection via ace_getRep(). Aborting.`);
       // log(`${funcName}: END - Error getting initial rep/selection`);
        return;
    }
    const start = currentRepInitial.selStart;
    const end = currentRepInitial.selEnd;
    const initialStartLine = start[0];
   // log(`${funcName}: Current selection from initial rep:`, { start, end });

   // log(`${funcName}: Phase 2 - Inserting text block...`);
    ed.ace_performDocumentReplaceRange(start, end, block);
   // log(`${funcName}: Inserted block of delimited text lines.`);
   // log(`${funcName}: Requesting text sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20);
   // log(`${funcName}: Text sync requested.`);

   // log(`${funcName}: Phase 3 - Applying metadata attributes to ${rows} inserted lines...`);
    const currentRep = ed.ace_getRep();
    if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get updated rep after text insertion. Cannot apply attributes reliably.`);
       // log(`${funcName}: END - Error getting updated rep`);
        return; 
    }
   // log(`${funcName}: Fetched updated rep for attribute application.`);

    for (let r = 0; r < rows; r++) {
      const lineNumToApply = initialStartLine + r;
     // log(`${funcName}: -> Processing row ${r} on line ${lineNumToApply}`);

      const lineEntry = currentRep.lines.atIndex(lineNumToApply);
      if (!lineEntry) {
       // log(`${funcName}: Could not find line entry for ${lineNumToApply}, skipping attribute application.`);
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
         // log(`${funcName}: Applying ${ATTR_CELL} attribute to Line ${lineNumToApply} Col ${c} Range ${offset}-${offset + cellContent.length}`);
          ed.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < cols - 1) {
          offset += DELIMITER.length;
        }
      }

      applyTableLineMetadataAttribute(lineNumToApply, tblId, r, cols, currentRep, ed, null, null); 
    }
   // log(`${funcName}: Finished applying metadata attributes.`);
   // log(`${funcName}: Requesting attribute sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20);
   // log(`${funcName}: Attribute sync requested.`);

   // log(`${funcName}: Phase 4 - Setting final caret position...`);
    const finalCaretLine = initialStartLine + rows;
    const finalCaretPos = [finalCaretLine, 0];
   // log(`${funcName}: Target caret position:`, finalCaretPos);
    try {
      ed.ace_performSelectionChange(finalCaretPos, finalCaretPos, false);
      // log(`${funcName}: Successfully set caret position.`);
    } catch(e) {
       console.error(`[ep_data_tables] ${funcName}: Error setting caret position after table creation:`, e);
      // log(`[ep_data_tables] ${funcName}: Error details:`, { message: e.message, stack: e.stack });
    }

   // log(`${funcName}: END - Refactored Phase 4`);
  };

  ed.ace_doDatatableOptions = (action) => {
    const funcName = 'ace_doDatatableOptions';
   // log(`${funcName}: START - Processing action: ${action}`);

    const editor = ed.ep_data_tables_editor;
    if (!editor) {
      console.error(`[ep_data_tables] ${funcName}: Could not get editor reference.`);
      return;
    }

    const lastClick = editor.ep_data_tables_last_clicked;
    if (!lastClick || !lastClick.tblId) {
     // log(`${funcName}: No table selected. Please click on a table cell first.`);
      console.warn('[ep_data_tables] No table selected. Please click on a table cell first.');
      return;
    }

   // log(`${funcName}: Operating on table ${lastClick.tblId}, clicked line ${lastClick.lineNum}, cell ${lastClick.cellIndex}`);

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

     // log(`${funcName}: Successfully obtained documentAttributeManager from stored reference.`);

      const tableLines = [];
      const totalLines = currentRep.lines.length();

      for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
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
                   // log(`${funcName}: Reconstructed metadata from DOM for line ${lineIndex}: ${lineAttrString}`);
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
                tableLines.push({
                  lineIndex,
                  row: lineMetadata.row,
                  cols: lineMetadata.cols,
                  lineText: lineEntry.text,
                  metadata: lineMetadata
                });
              }
            }
          }
        } catch (e) {
          continue;
        }
      }

      if (tableLines.length === 0) {
       // log(`${funcName}: No table lines found for table ${lastClick.tblId}`);
        return;
      }

      tableLines.sort((a, b) => a.row - b.row);
     // log(`${funcName}: Found ${tableLines.length} table lines`);

      const numRows = tableLines.length;
      const numCols = tableLines[0].cols;

      let targetRowIndex = -1;

      targetRowIndex = tableLines.findIndex(line => line.lineIndex === lastClick.lineNum);

      if (targetRowIndex === -1) {
       // log(`${funcName}: Direct line number match failed, searching by DOM structure...`);
        const clickedLineEntry = currentRep.lines.atIndex(lastClick.lineNum);
        if (clickedLineEntry && clickedLineEntry.lineNode) {
          const clickedTable = clickedLineEntry.lineNode.querySelector('table.dataTable[data-tblId="' + lastClick.tblId + '"]');
          if (clickedTable) {
            const clickedRowAttr = clickedTable.getAttribute('data-row');
            if (clickedRowAttr !== null) {
              const clickedRowNum = parseInt(clickedRowAttr, 10);
              targetRowIndex = tableLines.findIndex(line => line.row === clickedRowNum);
             // log(`${funcName}: Found target row by DOM attribute matching: row ${clickedRowNum}, index ${targetRowIndex}`);
            }
          }
        }
      }

      if (targetRowIndex === -1) {
       // log(`${funcName}: Warning: Could not find target row, defaulting to row 0`);
        targetRowIndex = 0;
      }

      const targetColIndex = lastClick.cellIndex || 0;

     // log(`${funcName}: Table dimensions: ${numRows} rows x ${numCols} cols. Target: row ${targetRowIndex}, col ${targetColIndex}`);

      let newNumCols = numCols;
      let success = false;

      switch (action) {
        case 'addTblRowA':
         // log(`${funcName}: Inserting row above row ${targetRowIndex}`);
          success = addTableRowAboveWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;

        case 'addTblRowB':
         // log(`${funcName}: Inserting row below row ${targetRowIndex}`);
          success = addTableRowBelowWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;

        case 'addTblColL':
         // log(`${funcName}: Inserting column left of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnLeftWithText(tableLines, targetColIndex, ed, docManager);
          break;

        case 'addTblColR':
         // log(`${funcName}: Inserting column right of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnRightWithText(tableLines, targetColIndex, ed, docManager);
          break;

        case 'delTblRow':
          const rowConfirmMessage = `Are you sure you want to delete Row ${targetRowIndex + 1} and all content within?`;
          if (!confirm(rowConfirmMessage)) {
           // log(`${funcName}: Row deletion cancelled by user`);
            return;
          }
         // log(`${funcName}: Deleting row ${targetRowIndex}`);
          success = deleteTableRowWithText(tableLines, targetRowIndex, ed, docManager);
          break;

        case 'delTblCol':
          const colConfirmMessage = `Are you sure you want to delete Column ${targetColIndex + 1} and all content within?`;
          if (!confirm(colConfirmMessage)) {
           // log(`${funcName}: Column deletion cancelled by user`);
            return;
          }
         // log(`${funcName}: Deleting column ${targetColIndex}`);
          newNumCols = numCols - 1;
          success = deleteTableColumnWithText(tableLines, targetColIndex, ed, docManager);
          break;

        default:
         // log(`${funcName}: Unknown action: ${action}`);
          return;
      }

      if (!success) {
        console.error(`[ep_data_tables] ${funcName}: Table operation failed for action: ${action}`);
        return;
      }

     // log(`${funcName}: Table operation completed successfully with text and metadata synchronization`);

    } catch (error) {
      console.error(`[ep_data_tables] ${funcName}: Error during table operation:`, error);
     // log(`${funcName}: Error details:`, { message: error.message, stack: error.stack });
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
               // log('[ep_data_tables] addTableRowAbove: Extracted column widths from DOM:', columnWidths);
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
                // log('[ep_data_tables] addTableRowBelow: Extracted column widths from DOM:', columnWidths);
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
       // log('[ep_data_tables] Deleting first row (row 0) - inserting blank line to prevent table from getting stuck');
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
                   // log('[ep_data_tables] deleteTableRow: Extracted column widths from DOM:', columnWidths);
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


 // log('aceInitialized: END - helpers defined.');
};

exports.aceStartLineAndCharForPoint = () => { return undefined; };
exports.aceEndLineAndCharForPoint   = () => { return undefined; };

exports.aceSetAuthorStyle = (hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceSetAuthorStyle]';
 // log(`${logPrefix} START`, { hook, ctx });

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd || !ctx.key) {
   // log(`${logPrefix} No selection or style key. Allowing default.`);
    return;
  }

  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];

  if (startLine !== endLine) {
   // log(`${logPrefix} Selection spans multiple lines. Preventing style application to protect table structure.`);
    return false;
  }

  const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(startLine, ATTR_TABLE_JSON);
  if (!lineAttrString) {
   // log(`${logPrefix} Line ${startLine} is not a table line. Allowing default style application.`);
    return;
  }

  const BLOCKED_STYLES = [
    'list', 'listType', 'indent', 'align', 'heading', 'code', 'quote',
    'horizontalrule', 'pagebreak', 'linebreak', 'clear'
  ];

  if (BLOCKED_STYLES.includes(ctx.key)) {
   // log(`${logPrefix} Blocked potentially harmful style '${ctx.key}' from being applied to table cell.`);
    return false;
  }

  try {
    const tableMetadata = JSON.parse(lineAttrString);
    if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
     // log(`${logPrefix} Invalid table metadata. Preventing style application.`);
      return false;
    }

    const lineText = ctx.rep.lines.atIndex(startLine)?.text || '';
    const cells = lineText.split(DELIMITER);
    let currentOffset = 0;
    let selectionStartCell = -1;
    let selectionEndCell = -1;

    for (let i = 0; i < cells.length; i++) {
      const cellLength = cells[i]?.length ?? 0;
      const cellEndCol = currentOffset + cellLength;

      if (ctx.rep.selStart[1] >= currentOffset && ctx.rep.selStart[1] <= cellEndCol) {
        selectionStartCell = i;
      }
      if (ctx.rep.selEnd[1] >= currentOffset && ctx.rep.selEnd[1] <= cellEndCol) {
        selectionEndCell = i;
      }
      currentOffset += cellLength + DELIMITER.length;
    }

    if (selectionStartCell !== selectionEndCell) {
     // log(`${logPrefix} Selection spans multiple cells. Preventing style application to protect table structure.`);
      return false;
    }

    const cellStartCol = cells.slice(0, selectionStartCell).reduce((acc, cell) => acc + cell.length + DELIMITER.length, 0);
    const cellEndCol = cellStartCol + cells[selectionStartCell].length;

    if (ctx.rep.selStart[1] <= cellStartCol || ctx.rep.selEnd[1] >= cellEndCol) {
     // log(`${logPrefix} Selection includes cell delimiters. Preventing style application to protect table structure.`);
      return false;
    }

   // log(`${logPrefix} Style '${ctx.key}' allowed within cell boundaries.`);
    return;
  } catch (e) {
    console.error(`${logPrefix} Error processing style application:`, e);
   // log(`${logPrefix} Error details:`, { message: e.message, stack: e.stack });
    return false;
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
    
    // CRITICAL: Check for orphan table-related spans that are NOT inside a table
    // These are corrupted fragments that would create extra lines if not suppressed
    // NOTE: Check BOTH tbljson-* AND tblCell-* classes - browser may strip tbljson- during IME
    if (!table) {
      // Check if parent or ancestors have tbljson-* OR tblCell-* class but no table ancestor
      // Use classList for reliable class detection across browsers
      const hasTableRelatedClass = (el) => {
        if (!el) return false;
        // Method 1: Check classList (most reliable)
        if (el.classList && el.classList.length > 0) {
          for (const cls of el.classList) {
            // Check BOTH tbljson-* and tblCell-* patterns
            if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) return true;
          }
        }
        // Method 2: Fallback to className string check
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
          // Capture which class triggered detection for logging
          try {
            for (const cls of (checkEl.classList || [])) {
              if (cls.startsWith('tbljson-') || cls.startsWith('tblCell-')) {
                detectedClass = cls;
                break;
              }
            }
          } catch (_) {}
          break;
        }
        checkEl = checkEl.parentElement;
      }
      
      if (foundOrphanTableSpan) {
        // Check if this line ALREADY has a rendered table
        // If NOT, this is likely fresh paste content awaiting table construction - DON'T suppress
        // Only suppress if there's an existing table and this span is outside it (true orphan)
        try {
          const lineDiv = parentEl.closest('div.ace-line');
          if (lineDiv) {
            const existingTable = lineDiv.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
            if (!existingTable) {
              // NO table on this line yet - this is fresh paste content, let it through
              // The tbljson span will trigger acePostWriteDomLineHTML to build the table
              console.debug('[ep_data_tables:collector] allowing fresh tbljson content (no table yet)', {
                lineId: lineDiv.id || null,
                detectedClass,
                originalText: node.nodeValue?.slice(0, 50),
              });
              // DON'T suppress - let normal collection proceed to capture the tbljson content
              return; // Exit without modifying context.text
            }
          }
        } catch (_) {}
        
        // There IS a table on this line, but this span is outside it - THAT's a true orphan, suppress
        context.text = '';
        try {
          const lineDiv = parentEl.closest('div.ace-line');
          const classStr = parentEl.className ? String(parentEl.className) : '';
          console.debug('[ep_data_tables:collector] SUPPRESSED orphan table span (table exists)', {
            lineId: lineDiv?.id || null,
            parentClass: classStr.slice(0, 100),
            detectedClass,
            originalText: node.nodeValue?.slice(0, 50),
          });
        } catch (_) {}
        return;
      }
      
      // CRITICAL FIX: Check if this ace-line contains a table even though this text node isn't in it
      // If so, suppress this text - it's stray content that shouldn't be part of the table row
      // This handles cases where IME corruption puts non-table content on the same line as a table
      try {
        const lineDiv = parentEl.closest('div.ace-line');
        if (lineDiv) {
          const tableInLine = lineDiv.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]');
          if (tableInLine) {
            // This ace-line has a table, but this text node is NOT inside the table
            // This is stray content that would corrupt the table - suppress it
            context.text = '';
            console.debug('[ep_data_tables:collector] SUPPRESSED non-table content in table line', {
              lineId: lineDiv.id || null,
              originalText: node.nodeValue?.slice(0, 50),
              tblId: tableInLine.getAttribute('data-tblId') || tableInLine.getAttribute('data-tblid'),
            });
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
      try {
        const lineDiv = parentEl.closest && parentEl.closest('div.ace-line');
        const rep = cc && cc.rep;
        const lineIndex = (lineDiv && rep && rep.lines && typeof rep.lines.indexOfKey === 'function')
          ? rep.lines.indexOfKey(lineDiv.id)
          : null;
        console.debug('[ep_data_tables:collector] suppress-extra-text', {
          lineId: lineDiv && lineDiv.id || null,
          lineIndex,
        });
      } catch (_) {}
      return; 
    }
    if (state) state._epDT_emittedCanonical = true;

    const tr = table.querySelector('tbody > tr');
    if (!tr || !tr.children || tr.children.length === 0) return;

    // Sanitize cell text, but PRESERVE zero-width characters for images
    // Images use U+200B as placeholder - if we strip it, images disappear
    const sanitize = (s, hasImage) => {
      let x = (s || '').replace(new RegExp(DELIMITER, 'g'), ' ');
      // Only strip zero-width chars if this cell doesn't contain images
      // Images rely on U+200B as their text placeholder
      if (!hasImage) {
        x = x.replace(/[\u200B\u200C\u200D\uFEFF]/g, '');
      }
      if (!x) x = ' ';
      return x;
    };
    
    // Detect if a cell contains image content by checking for image classes OR ZWS in content
    // We check content directly because after Grammarly edits, image spans may be corrupted
    // but the ZWS placeholder characters may still be in the text
    const cellHasImage = (td) => {
      // Check for image spans
      const imageSpan = td.querySelector('[class*="image:"], [class*="inline-image"], span.image-placeholder');
      if (imageSpan) return true;
      // Also preserve if content already contains ZWS (likely image placeholder)
      // This handles cases where the image DOM structure is temporarily corrupted
      const text = td.innerText || td.textContent || '';
      if (/[\u200B\u200C\u200D\uFEFF]/.test(text)) return true;
      return false;
    };
    
    const cells = Array.from(tr.children).map((td) => sanitize(td.innerText || '', cellHasImage(td)));
    const canonical = cells.join(DELIMITER);
    context.text = canonical;

    try {
      const lineDiv = parentEl.closest && parentEl.closest('div.ace-line');
      const rep = cc && cc.rep;
      const lineIndex = (lineDiv && rep && rep.lines && typeof rep.lines.indexOfKey === 'function')
        ? rep.lines.indexOfKey(lineDiv.id)
        : null;
      console.debug('[ep_data_tables:collector] emit-canonical', {
        lineId: lineDiv && lineDiv.id || null,
        lineIndex,
        cells: cells.length,
        textSample: canonical.slice(0, 100),
      });
    } catch (_) {}

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
 // log(`${funcName}: Starting resize for column ${columnIndex}`);

  isResizing = true;
  resizeStartX = startX;
  resizeCurrentX = startX;
  resizeTargetTable = table;
  resizeTargetColumn = columnIndex;
  resizeTableMetadata = metadata;
  resizeLineNum = lineNum;

  const numCols = metadata.cols;
  resizeOriginalWidths = metadata.columnWidths ? [...metadata.columnWidths] : Array(numCols).fill(100 / numCols);

 // log(`${funcName}: Original widths:`, resizeOriginalWidths);

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

 // log(`createResizeOverlay: Found ${allTableRows.length} table rows, total height: ${totalTableHeight}px`);

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

 // log('createResizeOverlay: Created Google Docs style blue line overlay spanning entire table height');
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
   // log('finishColumnResize: Not in resize mode');
                      return;
                  }

  const funcName = 'finishColumnResize';
 // log(`${funcName}: Finishing resize`);

  const tableRect = resizeTargetTable.getBoundingClientRect();
  const deltaX = resizeCurrentX - resizeStartX;
  const deltaPercent = (deltaX / tableRect.width) * 100;

 // log(`${funcName}: Mouse moved ${deltaX}px (${deltaPercent.toFixed(1)}%)`);

  const finalWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;

  if (nextColumn < finalWidths.length) {
    const transfer = Math.min(deltaPercent, finalWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(finalWidths[currentColumn] - 5));

    finalWidths[currentColumn] += actualTransfer;
    finalWidths[nextColumn] -= actualTransfer;

   // log(`${funcName}: Transferred ${actualTransfer.toFixed(1)}% from column ${nextColumn} to column ${currentColumn}`);
  }

  const totalWidth = finalWidths.reduce((sum, width) => sum + width, 0);
  if (totalWidth > 0) {
    finalWidths.forEach((width, index) => {
      finalWidths[index] = (width / totalWidth) * 100;
    });
  }

 // log(`${funcName}: Final normalized widths:`, finalWidths.map(w => w.toFixed(1) + '%'));

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
   // log(`${callWithAceLogPrefix}: Finding and updating all table rows with tblId: ${resizeTableMetadata.tblId}`);

    try {
      const rep = ace.ace_getRep();
      if (!rep || !rep.lines) {
        console.error(`${callWithAceLogPrefix}: Invalid rep`);
        return;
      }

      const tableLines = [];
      const totalLines = rep.lines.length();

      for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
        try {
          let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

          if (lineAttrString) {
            const lineMetadata = JSON.parse(lineAttrString);
            if (lineMetadata.tblId === resizeTableMetadata.tblId) {
              tableLines.push({
                lineIndex,
                metadata: lineMetadata
              });
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
                   // log(`${callWithAceLogPrefix}: Reconstructed metadata from DOM for line ${lineIndex}:`, reconstructedMetadata);
                    tableLines.push({
                      lineIndex,
                      metadata: reconstructedMetadata
                    });
                  }
                }
              }
            }
                  }
              } catch (e) {
          continue;
        }
      }

     // log(`${callWithAceLogPrefix}: Found ${tableLines.length} table lines to update`);

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

       // log(`${callWithAceLogPrefix}: Updating line ${tableLine.lineIndex} (row ${tableLine.metadata.row}) with new column widths`);

        ace.ace_performDocumentApplyAttributesToRange(rangeStart, rangeEnd, [
          [ATTR_TABLE_JSON, updatedMetadataString]
        ]);
      }

     // log(`${callWithAceLogPrefix}: Successfully applied updated column widths to all ${tableLines.length} table rows`);

    } catch (error) {
      console.error(`${callWithAceLogPrefix}: Error applying updated metadata:`, error);
     // log(`${callWithAceLogPrefix}: Error details:`, { message: error.message, stack: error.stack });
    }
  }, 'applyTableResizeToAllRows', true);

 // log(`${funcName}: Column width update initiated for all table rows via ace_callWithAce`);

  resizeStartX = 0;
  resizeCurrentX = 0;
  resizeTargetTable = null;
  resizeTargetColumn = -1;
  resizeOriginalWidths = [];
  resizeTableMetadata = null;
  resizeLineNum = -1;

 // log(`${funcName}: Resize complete - state reset`);
};

exports.aceUndoRedo = (hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceUndoRedo]';
 // log(`${logPrefix} START`, { hook, ctx });

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd) {
   // log(`${logPrefix} No selection or context. Allowing default.`);
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
   // log(`${logPrefix} No table lines affected. Allowing default undo/redo.`);
    return;
  }

 // log(`${logPrefix} Table lines affected:`, { tableLines });

  try {
    for (const line of tableLines) {
      const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
      if (!lineAttrString) continue;

      const tableMetadata = JSON.parse(lineAttrString);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
       // log(`${logPrefix} Invalid table metadata after undo/redo. Attempting recovery.`);
        const lineText = ctx.rep.lines.atIndex(line)?.text || '';
        const cells = lineText.split(DELIMITER);

        if (cells.length > 1) {
          const newMetadata = {
            cols: cells.length,
            rows: 1,
            cells: cells.map((_, i) => ({ col: i, row: 0 }))
          };

          ctx.documentAttributeManager.setAttributeOnLine(line, ATTR_TABLE_JSON, JSON.stringify(newMetadata));
         // log(`${logPrefix} Recovered table structure for line ${line}`);
        } else {
          ctx.documentAttributeManager.removeAttributeOnLine(line, ATTR_TABLE_JSON);
         // log(`${logPrefix} Removed invalid table attribute from line ${line}`);
        }
      }
    }
  } catch (e) {
    console.error(`${logPrefix} Error during undo/redo validation:`, e);
   // log(`${logPrefix} Error details:`, { message: e.message, stack: e.stack });
  }
};


