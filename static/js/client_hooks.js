/* ep_data_tables – attribute‑based tables (line‑class + PostWrite renderer)
 * -----------------------------------------------------------------
 * Strategy
 *   • One line attribute   tbljson = JSON({tblId,row,cells:[{txt:"…"},…]})
 *   • One char‑range attr  td      = column‑index (string)
 *   • `aceLineAttribsToClasses` puts class `tbl-line` on the **line div** so
 *     we can catch it once per line in `acePostWriteDomLineHTML`.
 *   • Renderer accumulates rows that share the same tblId in a buffer on
 *     innerdocbody, flushes to a single <table> when the run ends.
 *   • No raw JSON text is ever visible to the user.
 */

/* eslint-env browser */

// ────────────────────────────── constants ──────────────────────────────
const ATTR_TABLE_JSON = 'tbljson';
const ATTR_CELL       = 'td';
const ATTR_CLASS_PREFIX = 'tbljson-'; // For finding the class in DOM
const log             = (...m) => console.debug('[ep_data_tables:client_hooks]', ...m);
const DELIMITER       = '\u241F';   // Internal column delimiter (␟)
// Use the same rare character inside the hidden span so acePostWriteDomLineHTML can
// still find delimiters when it splits node.innerHTML.
// Users never see this because the span is contenteditable=false and styled away.
const HIDDEN_DELIM    = DELIMITER;

// helper for stable random ids
const rand = () => Math.random().toString(36).slice(2, 8);

// encode/decode so JSON can survive as a CSS class token if ever needed
const enc = s => btoa(s).replace(/\+/g, '-').replace(/\//g, '_');
const dec = s => {
    // Revert to simpler decode, assuming enc provides valid padding
    const str = s.replace(/-/g, '+').replace(/_/g, '/');
    try {
        if (typeof atob === 'function') {
            return atob(str); // Browser environment
        } else if (typeof Buffer === 'function') {
            // Node.js environment
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

// NEW: Module-level state for last clicked cell
let lastClickedCellInfo = null; // { lineNum: number, cellIndex: number, tblId: string }

// NEW: Module-level state for column resizing (similar to images plugin)
let isResizing = false;
let resizeStartX = 0;
let resizeCurrentX = 0; // Track current mouse position
let resizeTargetTable = null;
let resizeTargetColumn = -1;
let resizeOriginalWidths = [];
let resizeTableMetadata = null;
let resizeLineNum = -1;
let resizeOverlay = null; // Visual overlay element

// ─────────────────── Reusable Helper Functions ───────────────────

/**
 * Recursively search for an element with a 'tbljson-' class inside a given element.
 * This is used to find the metadata carrier when it's nested inside block elements.
 * @param {HTMLElement} element - The root element to start searching from.
 * @returns {HTMLElement|null} - The found element or null.
 */
function findTbljsonElement(element) {
  if (!element) return null;
  // Check if this element has the tbljson class
  if (element.classList) {
    for (const cls of element.classList) {
      if (cls.startsWith(ATTR_CLASS_PREFIX)) {
        return element;
      }
    }
  }
  // Recursively check children
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
    // First, try the fast path: getting the attribute directly from the line.
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

    // Fallback for block-styled lines.
    // log(`${funcName}: No valid attribute on line ${lineNum}, checking DOM.`);
    const rep = editorInfo.ace_getRep();
    
    // This is the fix: Get the lineNode directly from the rep. It's more reliable
    // than querying the DOM and avoids the ace_getOuterDoc() call which was failing.
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

// ─────────────────── Cell Navigation Helper Functions ───────────────────
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
      // Shift+Tab: Move to previous cell
      targetCol--;
  if (targetCol < 0) {
        // Move to last cell of previous row
    targetRow--;
    targetCol = tableMetadata.cols - 1;
      }
    } else {
      // Tab: Move to next cell
      targetCol++;
      if (targetCol >= tableMetadata.cols) {
        // Move to first cell of next row
    targetRow++;
    targetCol = 0;
      }
  }

    // log(`${funcName}: Target coordinates - Row=${targetRow}, Col=${targetCol}`);
    
    // Find the line number for the target row
  const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);
    if (targetLineNum === -1) {
      // log(`${funcName}: Could not find line for target row ${targetRow}`);
      return false;
    }

    // Navigate to the target cell
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

    // Find the line number for the target row
  const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);

  if (targetLineNum !== -1) {
      // Found the row below, navigate to it.
      // log(`${funcName}: Found line for target row ${targetRow}, navigating.`);
      return navigateToCell(targetLineNum, targetCol, editorInfo, docManager);
    } else {
      // Could not find the row below, we must be on the last line.
      // Create a new, empty line after the table.
      // log(`${funcName}: Could not find next row. Creating new line after table.`);
  const rep = editorInfo.ace_getRep();
      const lineTextLength = rep.lines.atIndex(currentLineNum).text.length;
      const endOfLinePos = [currentLineNum, lineTextLength];

      // Move caret to end of the current line...
      editorInfo.ace_performSelectionChange(endOfLinePos, endOfLinePos, false);
      // ...and insert a newline character. This creates a new line below.
  editorInfo.ace_performDocumentReplaceRange(endOfLinePos, endOfLinePos, '\n');

      // The caret is automatically moved to the new line by the operation above,
      // but we ensure the visual selection is synced and the editor is focused.
      editorInfo.ace_updateBrowserSelectionFromRep();
  editorInfo.ace_focus();

      // We've now exited the table, so clear the last-clicked state.
      const editor = editorInfo.editor;
      if (editor) editor.ep_data_tables_last_clicked = null;
      // log(`${funcName}: Cleared last click info as we have exited the table.`);

      return true; // We handled it.
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
        
        // If no attribute found directly, check DOM (same logic as acePostWriteDomLineHTML)
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
        continue; // Skip lines with invalid metadata
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

    // --- NEW: Update plugin state BEFORE performing the UI action ---
    try {
      const editor = editorInfo.ep_data_tables_editor;
      // Use the new robust helper to get metadata, which handles block-styled lines.
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
    
    // The previous attempts involving wrappers and poking the renderer have all
    // proven to be unstable. The correct approach is to directly update the
    // internal model and then tell the browser to sync its visual selection to
    // that model.
    try {
      // 1. Update the internal representation of the selection.
    editorInfo.ace_performSelectionChange(targetPos, targetPos, false);
      // log(`${funcName}: Updated internal selection to [${targetPos}]`);

      // 2. Explicitly tell the editor to update the browser's visual selection
      // to match the new internal representation. This is the correct way to
      // make the caret appear in the new location without causing a race condition.
    editorInfo.ace_updateBrowserSelectionFromRep();
      // log(`${funcName}: Called updateBrowserSelectionFromRep to sync visual caret.`);
      
      // 3. Ensure the editor has focus.
    editorInfo.ace_focus();
      // log(`${funcName}: Editor focused.`);

    } catch(e) {
      console.error(`[ep_data_tables] ${funcName}: Error during direct navigation update:`, e);
      return false;
    }
    
  } catch (e) {
    // This synchronous catch is a fallback, though the error was happening asynchronously.
    console.error(`[ep_data_tables] ${funcName}: Error during cell navigation:`, e);
    return false;
  }

  // log(`${funcName}: Navigation considered successful.`);
  return true;
}

// ────────────────────── collectContentPre (DOM → atext) ─────────────────────
exports.collectContentPre = (hook, ctx) => {
  const funcName = 'collectContentPre';
  const node = ctx.domNode; // Etherpad's collector uses ctx.domNode
  const state = ctx.state;
  const cc = ctx.cc; // ContentCollector instance

  // log(`${funcName}: *** ENTRY POINT *** Hook: ${hook}, Node: ${node?.tagName}.${node?.className}`);

  // ***** START Primary Path: Reconstruct from rendered table *****
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
                let cellHTMLSegments = Array.from(trNode.children).map((td, index) => {
        let segmentHTML = td.innerHTML || '';
                  // log(`${funcName}: Line ${lineNum} TD[${index}] raw innerHTML (first 100): "${segmentHTML.substring(0,100)}"`);
                  
                  const resizeHandleRegex = /<div class="ep-data_tables-resize-handle"[^>]*><\/div>/ig;
                  segmentHTML = segmentHTML.replace(resizeHandleRegex, '');
                  // NEW: Also remove any previously injected hidden delimiter span so we do
                  // not serialise it back into the atext. Leaving it in would duplicate the
                  // hidden span on every save-reload cycle and, more importantly, confuse the
                  // later HTML-to-table reconstruction because the delimiter that lives *inside*
                  // the span would be mistaken for a real cell boundary.
                  const hiddenDelimRegexPrimary = /<span class="ep-data_tables-delim"[^>]*>.*?<\/span>/ig;
                  segmentHTML = segmentHTML.replace(hiddenDelimRegexPrimary, '');
                  // Remove caret-anchor spans (invisible, non-semantic)
                  const caretAnchorRegex = /<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig;
                  segmentHTML = segmentHTML.replace(caretAnchorRegex, '');
                  // If, after stripping tags/entities, the content is empty, serialize as empty string
                  const textCheck = segmentHTML.replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
                  if (textCheck === '') segmentHTML = '';

        const hidden = index === 0 ? '' :
        /* keep the char in the DOM but make it visually disappear and non-editable */
        `<span class="ep-data_tables-delim" contenteditable="false">${HIDDEN_DELIM}</span>`;
                  // log(`${funcName}: Line ${lineNum} TD[${index}] cleaned innerHTML (first 100): "${segmentHTML.substring(0,100)}"`);
        return segmentHTML;
      });

      if (cellHTMLSegments.length !== existingMetadata.cols) {
                    // log(`${funcName}: WARNING Line ${lineNum}: Reconstructed cell count (${cellHTMLSegments.length}) does not match metadata cols (${existingMetadata.cols}). Padding/truncating.`);
        while (cellHTMLSegments.length < existingMetadata.cols) cellHTMLSegments.push('');
                    if (cellHTMLSegments.length > existingMetadata.cols) cellHTMLSegments.length = existingMetadata.cols;
                }

                const canonicalLineText = cellHTMLSegments.join(DELIMITER);
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
                
                let cellHTMLSegments = Array.from(trNode.children).map((td, index) => {
                  let segmentHTML = td.innerHTML || '';
                  const resizeHandleRegex = /<div class="ep-data_tables-resize-handle"[^>]*><\/div>/ig;
                  segmentHTML = segmentHTML.replace(resizeHandleRegex, '');
                  if (index > 0) {
                    const hiddenDelimRegex = new RegExp('^<span class="ep-data_tables-delim" contenteditable="false">' + DELIMITER + '(<\\/span>)?<\\/span>', 'i');
                    segmentHTML = segmentHTML.replace(hiddenDelimRegex, '');
                  }
                  // Remove caret-anchor spans (invisible, non-semantic)
                  const caretAnchorRegex = /<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig;
                  segmentHTML = segmentHTML.replace(caretAnchorRegex, '');
                  // If, after stripping tags/entities, the content is empty, serialize as empty string
                  const textCheck = segmentHTML.replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
                  if (textCheck === '') segmentHTML = '';
                  return segmentHTML;
                });
                
                if (cellHTMLSegments.length !== domCols) {
                     // log(`${funcName}: WARNING Line ${lineNum} (Fallback): Reconstructed cell count (${cellHTMLSegments.length}) does not match DOM cols (${domCols}).`);
                     while(cellHTMLSegments.length < domCols) cellHTMLSegments.push('');
                     if(cellHTMLSegments.length > domCols) cellHTMLSegments.length = domCols;
                }

                const canonicalLineText = cellHTMLSegments.join(DELIMITER);
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
  // ***** END Primary Path *****


  // ***** Secondary Path: Apply attributes from tbljson-* class on spans (for initial creation/pasting) *****
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

// ───────────── attribute → span‑class mapping (linestylefilter hook) ─────────
exports.aceAttribsToClasses = (hook, ctx) => {
  const funcName = 'aceAttribsToClasses';
  // log(`>>>> ${funcName}: Called with key: ${ctx.key}`); // log entry
  if (ctx.key === ATTR_TABLE_JSON) {
    // log(`${funcName}: Processing ATTR_TABLE_JSON.`);
    // ctx.value is the raw JSON string from Etherpad's attribute pool
    const rawJsonValue = ctx.value;
    // log(`${funcName}: Received raw attribute value (ctx.value):`, rawJsonValue);

    // Attempt to parse for logging purposes
    let parsedMetadataForLog = '[JSON Parse Error]';
    try {
        parsedMetadataForLog = JSON.parse(rawJsonValue);
        // log(`${funcName}: Value parsed for logging:`, parsedMetadataForLog);
    } catch(e) {
        // log(`${funcName}: Error parsing raw JSON value for logging:`, e);
        // Continue anyway, enc() might still work if it's just a string
    }

    // Generate the class name by base64 encoding the raw JSON string.
    // This ensures acePostWriteDomLineHTML receives the expected encoded format.
    const className = `tbljson-${enc(rawJsonValue)}`;
    // log(`${funcName}: Generated class name by encoding raw JSON: ${className}`);
    return [className];
  }
  if (ctx.key === ATTR_CELL) {
    // Keep this in case we want cell-specific styling later
    // // log(`${funcName}: Processing ATTR_CELL: ${ctx.value}`); // Optional: Uncomment if needed
    return [`tblCell-${ctx.value}`];
  }
  // // log(`${funcName}: Processing other key: ${ctx.key}`); // Optional: Uncomment if needed
  return [];
};

// ───────────── line‑class mapping (REMOVE - superseded by aceAttribsToClasses) ─────────
// exports.aceLineAttribsToClasses = ... (Removed as aceAttribsToClasses adds table-line-data now)

// ─────────────────── Create Initial DOM Structure ────────────────────
// REMOVED - This hook doesn't reliably trigger on attribute changes during creation.
// exports.aceCreateDomLine = (hookName, args, cb) => { ... };

// Helper function to escape HTML (Keep this helper)
function escapeHtml(text = '') {
  const strText = String(text);
  var map = {
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  };
  return strText.replace(/[&<>"'']/g, function(m) { return map[m]; });
}

// NEW Helper function to build table HTML from pre-rendered delimited content with resize handles
function buildTableFromDelimitedHTML(metadata, innerHTMLSegments) {
  const funcName = 'buildTableFromDelimitedHTML';
  // log(`${funcName}: START`, { metadata, innerHTMLSegments });

  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined') {
    console.error(`[ep_data_tables] ${funcName}: Invalid or missing metadata. Aborting.`);
    // log(`${funcName}: END - Error`);
    return '<table class="dataTable dataTable-error"><tbody><tr><td>Error: Missing table metadata</td></tr></tbody></table>'; // Return error table
  }

  // Get column widths from metadata, or use equal distribution if not set
  const numCols = innerHTMLSegments.length;
  const columnWidths = metadata.columnWidths || Array(numCols).fill(100 / numCols);
  
  // Ensure we have the right number of column widths
  while (columnWidths.length < numCols) {
    columnWidths.push(100 / numCols);
  }
  if (columnWidths.length > numCols) {
    columnWidths.splice(numCols);
  }

  // Basic styling - can be moved to CSS later
  const tdStyle = `padding: 5px 7px; word-wrap:break-word; vertical-align: top; border: 1px solid #000; position: relative;`; // Added position: relative

  // Precompute encoded tbljson class so empty cells can carry the same marker
  let encodedTbljsonClass = '';
  try {
    encodedTbljsonClass = `tbljson-${enc(JSON.stringify(metadata))}`;
  } catch (_) { encodedTbljsonClass = ''; }

  // Map the HTML segments directly into TD elements with column widths
  const cellsHtml = innerHTMLSegments.map((segment, index) => {
    // Build the hidden delimiter *inside* the first author span so the caret
    // cannot sit between delimiter and text. For empty cells, synthesize a span
    // that carries tbljson and tblCell-N so caret anchoring remains stable.
    const textOnly = (segment || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
    let modifiedSegment = segment || '';
    const isEmpty = !segment || textOnly === '';
    if (isEmpty) {
      const cellClass = encodedTbljsonClass ? `${encodedTbljsonClass} tblCell-${index}` : `tblCell-${index}`;
      modifiedSegment = `<span class="${cellClass}">&nbsp;</span>`;
    }
    if (index > 0) {
      const delimSpan = `<span class="ep-data_tables-delim" contenteditable="false">${HIDDEN_DELIM}</span>`;
      // If the rendered segment already starts with a <span …> (which will be
      // the usual author-colour wrapper) inject the delimiter right after that
      // opening tag; otherwise just prefix it.
      modifiedSegment = modifiedSegment.replace(/^(<span[^>]*>)/i, `$1${delimSpan}`);
      if (!/^<span[^>]*>/i.test(modifiedSegment)) modifiedSegment = `${delimSpan}${modifiedSegment}`;
    }

    // --- NEW: Always embed the invisible caret-anchor as *last* child *within* the first span ---
    const caretAnchorSpan = '<span class="ep-data_tables-caret-anchor" contenteditable="false"></span>';
    const anchorInjected = modifiedSegment.replace(/<\/span>\s*$/i, `${caretAnchorSpan}</span>`);
    modifiedSegment = (anchorInjected !== modifiedSegment)
      ? anchorInjected
      : (isEmpty
          ? `<span class="${encodedTbljsonClass ? `${encodedTbljsonClass} ` : ''}tblCell-${index}">${modifiedSegment}${caretAnchorSpan}</span>`
          : `${modifiedSegment}${caretAnchorSpan}`);

    // Width & other decorations remain unchanged
    const widthPercent = columnWidths[index] || (100 / numCols);
    const cellStyle = `${tdStyle} width: ${widthPercent}%;`;

    const isLastColumn = index === innerHTMLSegments.length - 1;
    const resizeHandle = !isLastColumn ? 
      `<div class="ep-data_tables-resize-handle" data-column="${index}" style="position: absolute; top: 0; right: -2px; width: 4px; height: 100%; cursor: col-resize; background: transparent; z-index: 10;"></div>` : '';

    const tdContent = `<td style="${cellStyle}" data-column="${index}" draggable="false">${modifiedSegment}${resizeHandle}</td>`;
    return tdContent;
  }).join('');
  // log(`${funcName}: Joined all cellsHtml:`, cellsHtml);

  // Add 'dataTable-first-row' class if it's the logical first row (row index 0)
  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';
  // log(`${funcName}: First row class applied: '${firstRowClass}'`);

  // Construct the final table HTML
  // Rely on CSS for border-collapse, width etc. Add data attributes from metadata.
  const tableHtml = `<table class="dataTable${firstRowClass}" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse: collapse; table-layout: fixed;" draggable="false"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
  // log(`${funcName}: Generated final table HTML:`, tableHtml);
  // log(`${funcName}: END - Success`);
  return tableHtml;
}

// ───────────────── Populate Table Cells / Render (PostWrite) ──────────────────
exports.acePostWriteDomLineHTML = function (hook_name, args, cb) {
  const funcName = 'acePostWriteDomLineHTML';
  const node = args?.node;
  const nodeId = node?.id;
  const lineNum = args?.lineNumber; // Etherpad >= 1.9 provides lineNumber
  const logPrefix = '[ep_data_tables:acePostWriteDomLineHTML]'; // Consistent prefix

  // *** STARTUP LOGGING ***
  // log(`${logPrefix} ----- START ----- NodeID: ${nodeId} LineNum: ${lineNum}`);
  if (!node || !nodeId) {
      // log(`${logPrefix} ERROR - Received invalid node or node without ID. Aborting.`);
      console.error(`[ep_data_tables] ${funcName}: Received invalid node or node without ID.`);
    return cb();
  }

  // *** ENHANCED DEBUG: Log complete DOM state ***
  // log(`${logPrefix} NodeID#${nodeId}: COMPLETE DOM STRUCTURE DEBUG:`);
  // log(`${logPrefix} NodeID#${nodeId}: Node tagName: ${node.tagName}`);
  // log(`${logPrefix} NodeID#${nodeId}: Node className: ${node.className}`);
  // log(`${logPrefix} NodeID#${nodeId}: Node innerHTML length: ${node.innerHTML?.length || 0}`);
  // log(`${logPrefix} NodeID#${nodeId}: Node innerHTML (first 500 chars): "${(node.innerHTML || '').substring(0, 500)}"`);
  // log(`${logPrefix} NodeID#${nodeId}: Node children count: ${node.children?.length || 0}`);
  
  // log all child elements and their classes
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

  // --- 1. Find and Parse Metadata Attribute --- 
  // log(`${logPrefix} NodeID#${nodeId}: Searching for tbljson-* class...`);
  
  // ENHANCED Helper function to recursively search for tbljson class in all descendants
  function findTbljsonClass(element, depth = 0, path = '') {
    const indent = '  '.repeat(depth);
    // log(`${logPrefix} NodeID#${nodeId}: ${indent}Searching element: ${element.tagName || 'unknown'}, path: ${path}`);
    
    // Check the element itself
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
    
    // Recursively check all descendants
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

  // Search for tbljson class starting from the node
  // log(`${logPrefix} NodeID#${nodeId}: Starting recursive search for tbljson class...`);
  encodedJsonString = findTbljsonClass(node, 0, 'ROOT');
  
  if (encodedJsonString) {
    // log(`${logPrefix} NodeID#${nodeId}: *** SUCCESS: Found encoded tbljson class: ${encodedJsonString} ***`);
  } else {
    // log(`${logPrefix} NodeID#${nodeId}: *** NO TBLJSON CLASS FOUND ***`);
  } 

  // If no attribute found, it's not a table line managed by us
  if (!encodedJsonString) {
      // log(`${logPrefix} NodeID#${nodeId}: No tbljson-* class found. Assuming not a table line. END.`);
      
      // DEBUG: Add detailed logging to understand why tbljson class is missing
      // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Node tag: ${node.tagName}, Node classes:`, Array.from(node.classList || []));
      // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Node innerHTML (first 200 chars): "${(node.innerHTML || '').substring(0, 200)}"`);
      
      // Check if there are any child elements with classes
      if (node.children && node.children.length > 0) {
        for (let i = 0; i < Math.min(node.children.length, 5); i++) {
          const child = node.children[i];
          // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Child ${i} tag: ${child.tagName}, classes:`, Array.from(child.classList || []));
        }
      }
      
      // Check if there's already a table in this node (orphaned table)
      const existingTable = node.querySelector('table.dataTable[data-tblId]');
      if (existingTable) {
        const existingTblId = existingTable.getAttribute('data-tblId');
        const existingRow = existingTable.getAttribute('data-row');
        // log(`${logPrefix} NodeID#${nodeId}: DEBUG - Found orphaned table! TblId: ${existingTblId}, Row: ${existingRow}`);
        
        // This suggests the table exists but the tbljson class was lost
        // Check if we're in a post-resize situation
        if (existingTblId && existingRow !== null) {
          // log(`${logPrefix} NodeID#${nodeId}: POTENTIAL ISSUE - Table exists but no tbljson class. This may be a post-resize issue.`);
          
          // Try to look up what the metadata should be based on the table attributes
          const tableCells = existingTable.querySelectorAll('td');
          // log(`${logPrefix} NodeID#${nodeId}: Table has ${tableCells.length} cells`);
          
          // log the current line's attribute state if we can get line number
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

  // *** NEW CHECK: If table already rendered, skip regeneration ***
  const existingTable = node.querySelector('table.dataTable[data-tblId]');
  if (existingTable) {
      // log(`${logPrefix} NodeID#${nodeId}: Table already exists in DOM. Skipping innerHTML replacement.`);
      // Optionally, verify tblId matches metadata? For now, assume it's correct.
      // const existingTblId = existingTable.getAttribute('data-tblId');
      // try {
      //    const decoded = dec(encodedJsonString); 
      //    const currentMetadata = JSON.parse(decoded);
      //    if (existingTblId === currentMetadata?.tblId) { ... } 
      // } catch(e) { /* ignore validation error */ }
      return cb(); // Do nothing further
  }

  // log(`${logPrefix} NodeID#${nodeId}: Decoding and parsing metadata...`);
  try {
    const decoded = dec(encodedJsonString);
      // log(`${logPrefix} NodeID#${nodeId}: Decoded string: ${decoded}`);
      if (!decoded) throw new Error('Decoded string is null or empty.');
    rowMetadata = JSON.parse(decoded);
      // log(`${logPrefix} NodeID#${nodeId}: Parsed rowMetadata:`, rowMetadata);

      // Validate essential metadata
      if (!rowMetadata || typeof rowMetadata.tblId === 'undefined' || typeof rowMetadata.row === 'undefined' || typeof rowMetadata.cols !== 'number') {
          throw new Error('Invalid or incomplete metadata (missing tblId, row, or cols).');
      }
      // log(`${logPrefix} NodeID#${nodeId}: Metadata validated successfully.`);

  } catch(e) { 
      // log(`${logPrefix} NodeID#${nodeId}: FATAL ERROR - Failed to decode/parse/validate tbljson metadata. Rendering cannot proceed.`, e);
      console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Failed to decode/parse/validate tbljson.`, encodedJsonString, e);
      // Optionally render an error state in the node?
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Invalid table metadata attribute found.</div>';
      // log(`${logPrefix} NodeID#${nodeId}: Rendered error message in node. END.`);
    return cb();
  }
  // --- End Metadata Parsing ---

  // --- 2. Get and Parse Line Content ---
  // ALWAYS get the innerHTML of the line div itself to preserve all styling spans and attributes.
  // This innerHTML is set by Etherpad based on the line's current text in atext and includes
  // all the span elements with author colors, bold, italic, and other styling.
  // For an imported line's first render, atext is "Cell1|Cell2", so node.innerHTML will be "Cell1|Cell2".
  // For a natively created line, node.innerHTML is also "Cell1|Cell2".
  // After an edit, aceKeyEvent updates atext, and node.innerHTML reflects that new "EditedCell1|Cell2" string.
  // When styling is applied, it will include spans like: <span class="author-xxx bold">Cell1</span>|<span class="author-yyy italic">Cell2</span>
  const delimitedTextFromLine = node.innerHTML;
  // log(`${logPrefix} NodeID#${nodeId}: Using node.innerHTML for delimited text to preserve styling.`);
  // log(`${logPrefix} NodeID#${nodeId}: Raw innerHTML length: ${delimitedTextFromLine?.length || 0}`);
  // log(`${logPrefix} NodeID#${nodeId}: Raw innerHTML (first 1000 chars): "${(delimitedTextFromLine || '').substring(0, 1000)}"`);
  
  // *** ENHANCED DEBUG: Analyze delimiter presence ***
  const delimiterCount = (delimitedTextFromLine || '').split(DELIMITER).length - 1;
  // log(`${logPrefix} NodeID#${nodeId}: Delimiter '${DELIMITER}' count in innerHTML: ${delimiterCount}`);
  // log(`${logPrefix} NodeID#${nodeId}: Expected delimiters for ${rowMetadata.cols} columns: ${rowMetadata.cols - 1}`);
  
  // log all delimiter positions
  let pos = -1;
  const delimiterPositions = [];
  while ((pos = delimitedTextFromLine.indexOf(DELIMITER, pos + 1)) !== -1) {
    delimiterPositions.push(pos);
    // log(`${logPrefix} NodeID#${nodeId}: Delimiter found at position ${pos}, context: "${delimitedTextFromLine.substring(Math.max(0, pos - 20), pos + 21)}"`);
  }
  // log(`${logPrefix} NodeID#${nodeId}: All delimiter positions: [${delimiterPositions.join(', ')}]`);
  
  // The DELIMITER const is defined at the top of this file.
  // NEW: Remove all hidden-delimiter <span> wrappers **before** we split so
  // the embedded delimiter character they carry doesn't inflate or shrink
  // the segment count.
  const spanDelimRegex = new RegExp('<span class="ep-data_tables-delim"[^>]*>' + DELIMITER + '<\\/span>', 'ig');
  const sanitizedHTMLForSplit = (delimitedTextFromLine || '')
    .replace(spanDelimRegex, '')
    // strip caret anchors from raw line html before split
    .replace(/<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig, '');
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
    // Check if segment contains image-related content
    if (segment.includes('image:') || segment.includes('image-placeholder') || segment.includes('currently-selected')) {
      // log(`${logPrefix} NodeID#${nodeId}: *** SEGMENT[${i}] CONTAINS IMAGE CONTENT ***`);
    }
  }

  // log(`${logPrefix} NodeID#${nodeId}: Parsed HTML segments (${htmlSegments.length}):`, htmlSegments.map(s => (s || '').substring(0,50) + (s && s.length > 50 ? '...' : '')));

  // --- Enhanced Validation with Automatic Structure Reconstruction --- 
  let finalHtmlSegments = htmlSegments;
  
  if (htmlSegments.length !== rowMetadata.cols) {
      // log(`${logPrefix} NodeID#${nodeId}: *** MISMATCH DETECTED *** - Attempting reconstruction.`);
      
      // Check if this is an image selection issue
      const hasImageSelected = delimitedTextFromLine.includes('currently-selected');
      const hasImageContent = delimitedTextFromLine.includes('image:');
      if (hasImageSelected) {
        // log(`${logPrefix} NodeID#${nodeId}: *** POTENTIAL CAUSE: Image selection state may be affecting segment parsing ***`);
      }

      // First attempt: reconstruct using DOM spans that carry tblCell-N classes
      let usedClassReconstruction = false;
      try {
        const cols = Math.max(0, Number(rowMetadata.cols) || 0);
        const grouped = Array.from({ length: cols }, () => '');
        const candidates = Array.from(node.querySelectorAll('[class*="tblCell-"]'));

        const classNum = (el) => {
          if (!el || !el.classList) return -1;
          for (const cls of el.classList) {
            const m = /^tblCell-(\d+)$/.exec(cls);
            if (m) return parseInt(m[1], 10);
          }
          return -1;
        };
        const hasAncestorWithSameCell = (el, n) => {
          let p = el?.parentElement;
          while (p) {
            if (p.classList && p.classList.contains(`tblCell-${n}`)) return true;
            p = p.parentElement;
          }
          return false;
        };

        for (const el of candidates) {
          const n = classNum(el);
          if (n >= 0 && n < cols) {
            if (!hasAncestorWithSameCell(el, n)) {
              grouped[n] += el.outerHTML || '';
            }
          }
        }
        const usable = grouped.some(s => s && s.trim() !== '');
        if (usable) {
          finalHtmlSegments = grouped.map(s => (s && s.trim() !== '') ? s : '&nbsp;');
          usedClassReconstruction = true;
          console.warn(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Reconstructed ${finalHtmlSegments.length} segments from tblCell-N classes.`);
        }
      } catch (e) {
        console.debug(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Class-based reconstruction error; falling back.`, e);
      }

      // Fallback: reconstruct from string segments
      if (!usedClassReconstruction) {
        const reconstructedSegments = [];
        if (htmlSegments.length === 1 && rowMetadata.cols > 1) {
          reconstructedSegments.push(htmlSegments[0]);
          for (let i = 1; i < rowMetadata.cols; i++) {
            reconstructedSegments.push('&nbsp;');
          }
        } else if (htmlSegments.length > rowMetadata.cols) {
          for (let i = 0; i < rowMetadata.cols - 1; i++) {
            reconstructedSegments.push(htmlSegments[i] || '&nbsp;');
          }
          const remainingSegments = htmlSegments.slice(rowMetadata.cols - 1);
          reconstructedSegments.push(remainingSegments.join('|') || '&nbsp;');
        } else {
          for (let i = 0; i < rowMetadata.cols; i++) {
            reconstructedSegments.push(htmlSegments[i] || '&nbsp;');
          }
        }
        finalHtmlSegments = reconstructedSegments;
      }

      // Only warn if we still don't have the right number of segments
      if (finalHtmlSegments.length !== rowMetadata.cols) {
        console.warn(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Could not reconstruct to expected ${rowMetadata.cols} segments. Got ${finalHtmlSegments.length}.`);
      }
  } else {
      // log(`${logPrefix} NodeID#${nodeId}: Segment count matches metadata cols (${rowMetadata.cols}). Using original segments.`);
  }

  // --- 3. Build and Render Table ---
  // log(`${logPrefix} NodeID#${nodeId}: Calling buildTableFromDelimitedHTML...`);
  try {
    const newTableHTML = buildTableFromDelimitedHTML(rowMetadata, finalHtmlSegments);
      // log(`${logPrefix} NodeID#${nodeId}: Received new table HTML from helper. Replacing content.`);
      
      // The old local findTbljsonElement is removed from here. We use the global one now.
      const tbljsonElement = findTbljsonElement(node);
      
      // If we found a tbljson element and it's nested in a block element, 
      // we need to preserve the block wrapper while replacing the content
      if (tbljsonElement && tbljsonElement.parentElement && tbljsonElement.parentElement !== node) {
        // Check if the parent is a block-level element that should be preserved
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
      // Replace the node's content entirely with the generated table
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
  // --- End Table Building ---

  // *** REMOVED CACHING LOGIC ***
  // The old logic based on tableRowNodes cache is completely removed.

  // log(`${logPrefix}: ----- END ----- NodeID: ${nodeId}`);
  return cb();
};

// NEW: Helper function to get line number (adapted from ep_image_insert)
// Ensure this is defined before it's used in postAceInit
function _getLineNumberOfElement(element) {
    // Implementation similar to ep_image_insert
    let currentElement = element;
    let count = 0;
    while (currentElement = currentElement.previousElementSibling) {
        count++;
    }
    return count;
}

// ───────────────────── Handle Key Events ─────────────────────
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

  // Get caret info from event context - may be stale
  const reportedLineNum = rep.selStart[0];
  const reportedCol = rep.selStart[1]; 
  // log(`${logPrefix} Reported caret from rep: Line=${reportedLineNum}, Col=${reportedCol}`);

  // --- Get Table Metadata for the reported line --- 
  let tableMetadata = null;
  let lineAttrString = null; // Store for potential use later
  try {
    // Add debugging to see what's happening with attribute retrieval
    // log(`${logPrefix} DEBUG: Attempting to get ${ATTR_TABLE_JSON} attribute from line ${reportedLineNum}`);
    lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
    // log(`${logPrefix} DEBUG: getAttributeOnLine returned: ${lineAttrString ? `"${lineAttrString}"` : 'null/undefined'}`);
    
    // Also check if there are any attributes on this line at all
    if (typeof docManager.getAttributesOnLine === 'function') {
      try {
        const allAttribs = docManager.getAttributesOnLine(reportedLineNum);
        // log(`${logPrefix} DEBUG: All attributes on line ${reportedLineNum}:`, allAttribs);
      } catch(e) {
        // log(`${logPrefix} DEBUG: Error getting all attributes:`, e);
      }
    }
    
    // NEW: Check if there's a table in the DOM even though attribute is missing
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
            // Try to reconstruct the metadata from DOM
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
             tableMetadata = null; // Ensure it's null if invalid
        }
    } else {
        // log(`${logPrefix} DEBUG: No ${ATTR_TABLE_JSON} attribute found on line ${reportedLineNum}`);
        // Not a table line based on reported caret line
    }
  } catch(e) {
    console.error(`${logPrefix} Error checking/parsing line attribute for line ${reportedLineNum}.`, e);
    tableMetadata = null; // Ensure it's null on error
  }

  // Get last known good state
  const editor = editorInfo.editor; // Get editor instance
  const lastClick = editor?.ep_data_tables_last_clicked; // Read shared state
  // log(`${logPrefix} Reading stored click/caret info:`, lastClick);

  // --- Determine the TRUE target line, cell, and caret position --- 
  let currentLineNum = -1;
  let targetCellIndex = -1;
  let relativeCaretPos = -1;
  let precedingCellsOffset = 0; 
  let cellStartCol = 0; 
  let lineText = '';
  let cellTexts = [];
  let metadataForTargetLine = null;
  let trustedLastClick = false; // Flag to indicate if we are using stored info

  // ** Scenario 1: Try to trust lastClick info **
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
          
          // Check if metadata is valid and tblId matches
          if (storedLineMetadata && typeof storedLineMetadata.cols === 'number' && storedLineMetadata.tblId === lastClick.tblId) {
              // log(`${logPrefix} Stored click info VALIDATED (Metadata OK and tblId matches). Trusting stored state.`);
              trustedLastClick = true;
              currentLineNum = lastClick.lineNum; 
              targetCellIndex = lastClick.cellIndex;
              metadataForTargetLine = storedLineMetadata; 
              lineAttrString = storedLineAttrString; // Use the validated attr string
              
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
                  relativeCaretPos = reportedCol - cellStartCol; // Use reportedCol for initial calc if relative is missing
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
           if (editor) editor.ep_data_tables_last_clicked = null; // Clear on error
      }
  }
  
  // ** Scenario 2: Fallback - Use reported line/col ONLY if stored info wasn't trusted **
  if (!trustedLastClick) {
      // log(`${logPrefix} Fallback: Using reported caret position Line=${reportedLineNum}, Col=${reportedCol}.`);
      // Fetch metadata for the reported line again, in case it wasn't fetched or was invalid earlier
      try {
          lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
          if (lineAttrString) tableMetadata = JSON.parse(lineAttrString);
          if (!tableMetadata || typeof tableMetadata.cols !== 'number') tableMetadata = null;
          
          // If no attribute found directly, check if there's a table in the DOM even though attribute is missing (block styles)
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
                  // Try to reconstruct the metadata from DOM
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
      } catch(e) { tableMetadata = null; } // Ignore errors here, handled below

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

      // Calculate target cell based on reportedCol
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

  // --- Final Validation --- 
  if (currentLineNum < 0 || targetCellIndex < 0 || !metadataForTargetLine || targetCellIndex >= metadataForTargetLine.cols) {
       // log(`${logPrefix} FAILED final validation: Line=${currentLineNum}, Cell=${targetCellIndex}, Metadata=${!!metadataForTargetLine}. Allowing default.`);
    if (editor) editor.ep_data_tables_last_clicked = null;
    return false;
  }

  // log(`${logPrefix} --> Final Target: Line=${currentLineNum}, CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}`);
  // --- End Cell/Position Determination ---

  // --- START NEW: Handle Highlight Deletion/Replacement ---
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

    let selectionStartColInLine = selStartActual[1];      // may be clamped
    let selectionEndColInLine = selEndActual[1];          // may be clamped

    const currentCellFullText = cellTexts[targetCellIndex] || '';
    // cellStartCol is already defined and calculated based on trustedLastClick or fallback
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

    // Allow selection even if it starts at the very first char, but be ready to restore

    if (isSelectionEntirelyWithinCell) {
      // Pure selection (no key pressed yet) – allow browser shortcuts such as
      // Ctrl-C / Ctrl-X / Cmd-C / Cmd-X to work.  We only take control for
      // real keydown events that would modify the cell (handled further below).

      // 1. Non-keydown events → let them bubble (copy/cut command happens on
      //    the subsequent "copy"/"cut" event).
      if (evt.type !== 'keydown') return false;

      // 2. Keydown that involves modifiers (Ctrl/Cmd/Alt) → we are not going
      //    to change the cell text, so let the browser handle it.
      if (evt.ctrlKey || evt.metaKey || evt.altKey) return false;

      // 3. For destructive or printable keys we fall through so the specialised
      //    highlight-deletion logic that follows can run.
    }

    const isCurrentKeyDelete = evt.key === 'Delete' || evt.keyCode === 46;
    const isCurrentKeyBackspace = evt.key === 'Backspace' || evt.keyCode === 8;
    // Check if it's a printable character, not a modifier
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
      const repBeforeEdit = editorInfo.ace_getRep(); // Get rep before edit for attribute helper
      // log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

      if (isCurrentKeyTyping) {
        replacementText = evt.key;
        newAbsoluteCaretCol = selectionStartColInLine + replacementText.length;
        // log(`${logPrefix} [selection] -> Replacing selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]] with text '${replacementText}'`);
      } else { // Delete or Backspace
        // log(`${logPrefix} [selection] -> Deleting selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]]`);
        // If whole cell is being wiped, keep a single space so cell isn't empty
        const isWholeCell = selectionStartColInLine <= cellContentStartColInLine && selectionEndColInLine >= cellContentEndColInLine;
        if (isWholeCell) {
          replacementText = ' ';
          newAbsoluteCaretCol = selectionStartColInLine + 1;
          // log(`${logPrefix} [selection] Whole cell cleared – inserting single space to preserve caret/author span.`);
        }
      }

      try {
        // const repBeforeEdit = editorInfo.ace_getRep(); // Get rep before edit for attribute helper - MOVED UP
        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, replacementText);

        // NEW: ensure the replacement text inherits the cell attribute so the
        //       author-span (& tblCell-N) comes back immediately
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
        
        // Add sync hint AFTER setting selection
    editorInfo.ace_fastIncorp(1);
        const repAfterFastIncorp = editorInfo.ace_getRep();
        // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
        // log(`${logPrefix} [selection] -> Requested sync hint (fastIncorp 1).`);

        // --- Re-assert selection --- 
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
        return true; // Still return true as we prevented default.
      }
    }
  }
  // --- END NEW: Handle Highlight Deletion/Replacement ---

  // --- Check for Ctrl+X (Cut) key combination ---
  const isCutKey = (evt.ctrlKey || evt.metaKey) && (evt.key === 'x' || evt.key === 'X' || evt.keyCode === 88);
  if (isCutKey && hasSelection) {
    // log(`${logPrefix} Ctrl+X (Cut) detected with selection. Letting cut event handler manage this.`);
    // Let the cut event handler handle this - we don't need to preventDefault here
    // as the cut event will handle the operation and prevent default
    return false; // Allow the cut event to be triggered
  } else if (isCutKey && !hasSelection) {
    // log(`${logPrefix} Ctrl+X (Cut) detected but no selection. Allowing default.`);
    return false; // Allow default - nothing to cut
  }

  // --- Define Key Types ---
  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete' || evt.keyCode === 46;
  const isBackspaceKey = evt.key === 'Backspace' || evt.keyCode === 8;
  const isNavigationKey = [33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode);
  const isTabKey = evt.key === 'Tab';
  const isEnterKey = evt.key === 'Enter';
  // log(`${logPrefix} Key classification: Typing=${isTypingKey}, Backspace=${isBackspaceKey}, Delete=${isDeleteKey}, Nav=${isNavigationKey}, Tab=${isTabKey}, Enter=${isEnterKey}, Cut=${isCutKey}`);

  /*
   * Prevent caret placement *after* the invisible caret-anchor.
   * – RIGHT (→) pressed at the end of a cell jumps to the next cell.
   * – LEFT  (←) pressed at the start of a cell jumps to the previous cell.
   * This avoids the narrow dead-zone that lives between the anchor and the
   * resize handle where typing previously caused content to drift into the
   * neighbouring column.
   */
  const currentCellTextLengthEarly = cellTexts[targetCellIndex]?.length ?? 0;

  if (evt.type === 'keydown' && !evt.ctrlKey && !evt.metaKey && !evt.altKey) {
    // Right-arrow – if at the very end of a cell, move to the next cell.
    if (evt.keyCode === 39 && relativeCaretPos >= currentCellTextLengthEarly && targetCellIndex < metadataForTargetLine.cols - 1) {
      // log(`${logPrefix} ArrowRight at cell boundary – navigating to next cell to avoid anchor zone.`);
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, false, editorInfo, docManager);
      return true;
    }

    // Left-arrow – if at the very start of a cell, move to the previous cell.
    if (evt.keyCode === 37 && relativeCaretPos === 0 && targetCellIndex > 0) {
      // log(`${logPrefix} ArrowLeft at cell boundary – navigating to previous cell to avoid anchor zone.`);
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, true, editorInfo, docManager);
      return true;
    }
  }

  // --- Handle Keys --- 

  // 1. Allow non-Tab navigation keys immediately
  if (isNavigationKey && !isTabKey) {
      // log(`${logPrefix} Allowing navigation key: ${evt.key}. Clearing click state.`);
      if (editor) editor.ep_data_tables_last_clicked = null; // Clear state on navigation
      return false;
  }

  // 2. Handle Tab - Navigate to next cell (only on keydown to avoid double navigation)
  if (isTabKey) { 
     // log(`${logPrefix} Tab key pressed. Event type: ${evt.type}`);
    evt.preventDefault();
     
     // Only process keydown events for navigation to avoid double navigation
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

  // 3. Handle Enter - Navigate to cell below (only on keydown to avoid double navigation)
  if (isEnterKey) {
      // log(`${logPrefix} Enter key pressed. Event type: ${evt.type}`);
    evt.preventDefault();
      
      // Only process keydown events for navigation to avoid double navigation
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

  // 4. Intercept destructive keys ONLY at cell boundaries to protect delimiters
      const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
  // Backspace at the very beginning of cell > 0
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex > 0) {
      // log(`${logPrefix} Intercepted Backspace at start of cell ${targetCellIndex}. Preventing default.`);
    evt.preventDefault();
          return true;
      }
  // NEW: Backspace at very beginning of first cell – would merge with previous line
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex === 0) {
        // log(`${logPrefix} Intercepted Backspace at start of first cell (line boundary). Preventing merge.`);
        evt.preventDefault();
        return true;
      }
  // Delete at the very end of cell < last cell
  if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex < metadataForTargetLine.cols - 1) {
      // log(`${logPrefix} Intercepted Delete at end of cell ${targetCellIndex}. Preventing default.`);
          evt.preventDefault();
          return true;
      }
  // NEW: Delete at very end of last cell – would merge with next line
      if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex === metadataForTargetLine.cols - 1) {
        // log(`${logPrefix} Intercepted Delete at end of last cell (line boundary). Preventing merge.`);
        evt.preventDefault();
        return true;
      }

  // 5. Handle Typing/Backspace/Delete WITHIN a cell via manual modification
  const isInternalBackspace = isBackspaceKey && relativeCaretPos > 0;
  const isInternalDelete = isDeleteKey && relativeCaretPos < currentCellTextLength;

  // Guard: internal Backspace at relativePos 1 (would delete delimiter) & Delete at relativePos 0
  if ((isInternalBackspace && relativeCaretPos === 1 && targetCellIndex > 0) ||
      (isInternalDelete && relativeCaretPos === 0 && targetCellIndex > 0)) {
    // log(`${logPrefix} Attempt to erase protected delimiter – operation blocked.`);
    evt.preventDefault();
    return true;
  }

  if (isTypingKey || isInternalBackspace || isInternalDelete) {
    // --- PREVENT TYPING DIRECTLY AFTER DELIMITER (relativeCaretPos===0) ---
    if (isTypingKey && relativeCaretPos === 0 && targetCellIndex > 0) {
      // log(`${logPrefix} Caret at forbidden position 0 (just after delimiter). Auto-advancing to position 1.`);
      const safePosAbs = cellStartCol + 1;
      editorInfo.ace_performSelectionChange([currentLineNum, safePosAbs], [currentLineNum, safePosAbs], false);
      editorInfo.ace_updateBrowserSelectionFromRep();
      relativeCaretPos = 1;
      // log(`${logPrefix} Caret moved to safe position. New relativeCaretPos=${relativeCaretPos}`);
    }
    // *** Use the validated currentLineNum and currentCol derived from relativeCaretPos ***
    const currentCol = cellStartCol + relativeCaretPos;
    // log(`${logPrefix} Handling INTERNAL key='${evt.key}' Type='${evt.type}' at Line=${currentLineNum}, Col=${currentCol} (CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}).`);
    // log(`${logPrefix} [caretTrace] Initial rep.selStart for internal edit: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

    // Only process keydown events for modifications
    if (evt.type !== 'keydown') {
        // log(`${logPrefix} Ignoring non-keydown event type ('${evt.type}') for handled key.`);
        return false; 
    }

    // log(`${logPrefix} Preventing default browser action for keydown event.`);
    evt.preventDefault();

    let newAbsoluteCaretCol = -1;
    let repBeforeEdit = null; // Store rep before edits for attribute helper

    try {
        repBeforeEdit = editorInfo.ace_getRep(); // Get rep *before* making changes
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
            newAbsoluteCaretCol = currentCol; // Caret stays at the same column for delete
        }
        const repAfterReplace = editorInfo.ace_getRep();
        // log(`${logPrefix} [caretTrace] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


        // *** CRITICAL: Re-apply the line attribute after ANY modification ***
        // log(`${logPrefix} -> Re-applying tbljson line attribute...`);
        
        // DEBUG: Log the values before calculating attrStringToApply
        // log(`${logPrefix} DEBUG: Before calculating attrStringToApply - trustedLastClick=${trustedLastClick}, reportedLineNum=${reportedLineNum}, currentLineNum=${currentLineNum}`);
        // log(`${logPrefix} DEBUG: lineAttrString value:`, lineAttrString ? `"${lineAttrString}"` : 'null/undefined');
        
        const applyHelper = editorInfo.ep_data_tables_applyMeta; 
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) { 
             // Pass the original lineAttrString if available AND if it belongs to the currentLineNum
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
                 applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager); // Cannot guarantee old attr string is valid here
                 // log(`${logPrefix} -> tbljson line attribute re-applied (using current rep fallback).`);
            } else {
                  console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute even with fallback rep.`);
             }
        }
        
        // Set caret position immediately
        if (newAbsoluteCaretCol >= 0) {
             const newCaretPos = [currentLineNum, newAbsoluteCaretCol]; // Use the trusted currentLineNum
             // log(`${logPrefix} -> Setting selection immediately to:`, newCaretPos);
             // log(`${logPrefix} [caretTrace] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
             try {
                editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                const repAfterSelectionChange = editorInfo.ace_getRep();
                // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);
                // log(`${logPrefix} -> Selection set immediately.`);

                // Add sync hint AFTER setting selection
                editorInfo.ace_fastIncorp(1); 
                const repAfterFastIncorp = editorInfo.ace_getRep();
                // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
                // log(`${logPrefix} -> Requested sync hint (fastIncorp 1).`);

                // --- Re-assert selection --- 
                const targetCaretPosForReassert = [currentLineNum, newAbsoluteCaretCol];
                // log(`${logPrefix} [caretTrace] Attempting to re-assert selection post-fastIncorp to [${targetCaretPosForReassert[0]}, ${targetCaretPosForReassert[1]}]`);
                editorInfo.ace_performSelectionChange(targetCaretPosForReassert, targetCaretPosForReassert, false);
                const repAfterReassert = editorInfo.ace_getRep();
                // log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

                // Store the updated caret info for the next event
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
        // Maybe return false to allow default as a fallback on error?
        // For now, return true as we prevented default.
    return true;
  }

    const endLogTime = Date.now();
    // log(`${logPrefix} END (Handled Internal Edit Manually) Key='${evt.key}' Type='${evt.type}' -> Returned true. Duration: ${endLogTime - startLogTime}ms`);
    return true; // We handled the key event

  } // End if(isTypingKey || isInternalBackspace || isInternalDelete)


  // Fallback for any other keys or edge cases not handled above
  const endLogTimeFinal = Date.now();
  // log(`${logPrefix} END (Fell Through / Unhandled Case) Key='${evt.key}' Type='${evt.type}'. Allowing default. Duration: ${endLogTimeFinal - startLogTime}ms`);
  // Clear click state if it wasn't handled?
  // if (editor?.ep_data_tables_last_clicked) editor.ep_data_tables_last_clicked = null;
  // log(`${logPrefix} [caretTrace] Final rep.selStart at end of aceKeyEvent (if unhandled): Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);
  return false; // Allow default browser/ACE handling
};

// ───────────────────── ace init + public helpers ─────────────────────
exports.aceInitialized = (h, ctx) => {
  const logPrefix = '[ep_data_tables:aceInitialized]';
  // log(`${logPrefix} START`, { hook_name: h, context: ctx });
  const ed = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

  // log(`${logPrefix} Attaching ep_data_tables_applyMeta helper to editorInfo.`);
  ed.ep_data_tables_applyMeta = applyTableLineMetadataAttribute;
  // log(`${logPrefix}: Attached applyTableLineMetadataAttribute helper to ed.ep_data_tables_applyMeta successfully.`);

  // Store the documentAttributeManager reference for later use
  // log(`${logPrefix} Storing documentAttributeManager reference on editorInfo.`);
  ed.ep_data_tables_docManager = docManager;
  // log(`${logPrefix}: Stored documentAttributeManager reference as ed.ep_data_tables_docManager.`);

  // *** ENHANCED: Paste event listener + Column resize listeners ***
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

    // Store editor reference for later use in table operations
    // log(`${logPrefix} Storing editor reference on editorInfo.`);
    ed.ep_data_tables_editor = editor;
    // log(`${logPrefix}: Stored editor reference as ed.ep_data_tables_editor.`);

    // Attempt to find the inner iframe body, similar to ep_image_insert
    let $inner;
    try {
      // log(`${callWithAceLogPrefix} Attempting to find inner iframe body for listener attachment.`);
      const $iframeOuter = $('iframe[name="ace_outer"]');
      if ($iframeOuter.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find outer iframe (ace_outer).`);
        // log(`${callWithAceLogPrefix} Failed to find ace_outer.`);
        return;
      }
      // log(`${callWithAceLogPrefix} Found ace_outer:`, $iframeOuter);

      const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
      if ($iframeInner.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find inner iframe (ace_inner).`);
        // log(`${callWithAceLogPrefix} Failed to find ace_inner within ace_outer.`);
        return;
      }
      // log(`${callWithAceLogPrefix} Found ace_inner:`, $iframeInner);

      const innerDocBody = $iframeInner.contents().find('body');
      if (innerDocBody.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find body element in inner iframe.`);
        // log(`${callWithAceLogPrefix} Failed to find body in ace_inner.`);
        return;
      }
      $inner = $(innerDocBody[0]); // Ensure it's a jQuery object of the body itself
      // log(`${callWithAceLogPrefix} Successfully found inner iframe body:`, $inner);
    } catch (e) {
      console.error(`${callWithAceLogPrefix} ERROR: Exception while trying to find inner iframe body:`, e);
      // log(`${callWithAceLogPrefix} Exception details:`, { message: e.message, stack: e.stack });
      return;
    }

    if (!$inner || $inner.length === 0) {
      console.error(`${callWithAceLogPrefix} ERROR: $inner is not valid after attempting to find iframe body. Cannot attach listeners.`);
      // log(`${callWithAceLogPrefix} $inner is invalid. Aborting.`);
      return;
    }

    // *** CUT EVENT LISTENER ***
    // log(`${callWithAceLogPrefix} Attaching cut event listener to $inner (inner iframe body).`);
    $inner.on('cut', (evt) => {
      const cutLogPrefix = '[ep_data_tables:cutHandler]';
      // log(`${cutLogPrefix} CUT EVENT TRIGGERED. Event object:`, evt);

      // log(`${cutLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        // log(`${cutLogPrefix} WARNING: Could not get representation or selection. Allowing default cut.`);
        console.warn(`${cutLogPrefix} Could not get rep or selStart.`);
        return; // Allow default
      }
      // log(`${cutLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      // log(`${cutLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      // Check if there's actually a selection to cut
      if (selStart[0] === selEnd[0] && selStart[1] === selEnd[1]) {
        // log(`${cutLogPrefix} No selection to cut. Allowing default cut.`);
        return; // Allow default - nothing to cut
      }

      // Check if selection spans multiple lines
      if (selStart[0] !== selEnd[0]) {
        // log(`${cutLogPrefix} WARNING: Selection spans multiple lines. Preventing cut to protect table structure.`);
        evt.preventDefault();
        return;
      }

      // log(`${cutLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (lineAttrString) {
        // Fast-path: attribute exists – parse it.
        try {
        tableMetadata = JSON.parse(lineAttrString);
        } catch {}
      }

      if (!tableMetadata) {
        // Fallback for block-styled rows – reconstruct via DOM helper.
        tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      }

      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        // log(`${cutLogPrefix} Line ${lineNum} is NOT a recognised table line. Allowing default cut.`);
        return; // Not a table line
      }

      // log(`${cutLogPrefix} Line ${lineNum} IS a table line. Metadata:`, tableMetadata);

      // Validate selection is within cell boundaries
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
        selStart[1] = cellStartCol;          // clamp
      }
      
      if (wouldClampEnd) {
        console.log(`[ep_data_tables:cut-handler] CLAMPING cut selection end from ${selEnd[1]} to ${cellEndCol}`);
        selEnd[1] = cellEndCol;              // clamp
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        // log(`${cutLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing cut to protect table structure.`);
        evt.preventDefault();
        return;
      }

      // If we reach here, the selection is entirely within a single cell - allow cut and preserve table structure
      // log(`${cutLogPrefix} Selection is entirely within cell ${targetCellIndex}. Intercepting cut to preserve table structure.`);
      evt.preventDefault();

      try {
        // Get the selected text to copy to clipboard
        const selectedText = lineText.substring(selStart[1], selEnd[1]);
        // log(`${cutLogPrefix} Selected text to cut: "${selectedText}"`);

        // Copy to clipboard manually
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(selectedText).then(() => {
            // log(`${cutLogPrefix} Successfully copied to clipboard via Navigator API.`);
          }).catch((err) => {
            console.warn(`${cutLogPrefix} Failed to copy to clipboard via Navigator API:`, err);
          });
        } else {
          // Fallback for older browsers
          // log(`${cutLogPrefix} Using fallback clipboard method.`);
          const textArea = document.createElement('textarea');
          textArea.value = selectedText;
          document.body.appendChild(textArea);
          textArea.select();
          try {
            document.execCommand('copy');
            // log(`${cutLogPrefix} Successfully copied to clipboard via execCommand fallback.`);
          } catch (err) {
            console.warn(`${cutLogPrefix} Failed to copy to clipboard via fallback:`, err);
          }
          document.body.removeChild(textArea);
        }

        // Now perform the deletion within the cell using ace operations
        // log(`${cutLogPrefix} Performing deletion via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${cutLogPrefix}[ace_callWithAceOps]`;
          // log(`${callAceLogPrefix} Entered ace_callWithAce for cut operations. selStart:`, selStart, `selEnd:`, selEnd);
          
          // log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to delete selected text.`);
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');
          // log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

          // --- Ensure cell is not left empty (zero-length) ---
          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
            // log(`${callAceLogPrefix} Cell ${targetCellIndex} became empty after cut – inserting single space to preserve structure.`);
            const insertPos = [lineNum, selStart[1]]; // Start of the now-empty cell
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            // NEW – re-apply td attribute to the freshly inserted space
            const attrStart = insertPos;
            const attrEnd   = [insertPos[0], insertPos[1] + 1];
            aceInstance.ace_performDocumentApplyAttributesToRange(
              attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

          // log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
          const repAfterCut = aceInstance.ace_getRep();
          // log(`${callAceLogPrefix} Fetched rep after cut for applyMeta. Line ${lineNum} text now: "${repAfterCut.lines.atIndex(lineNum).text}"`);
          
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
          // log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

          const newCaretPos = [lineNum, selStart[1]];
          // log(`${callAceLogPrefix} Setting caret position to: [${newCaretPos}].`);
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
          // log(`${callAceLogPrefix} Selection change successful.`);

          // log(`${callAceLogPrefix} Cut operations within ace_callWithAce completed successfully.`);
        }, 'tableCutTextOperations', true);

        // log(`${cutLogPrefix} Cut operation completed successfully.`);
      } catch (error) {
        console.error(`${cutLogPrefix} ERROR during cut operation:`, error);
        // log(`${cutLogPrefix} Cut operation failed. Error details:`, { message: error.message, stack: error.stack });
      }
    });

    // *** BEFOREINPUT EVENT LISTENER FOR CONTEXT-MENU DELETE ***
    // log(`${callWithAceLogPrefix} Attaching beforeinput event listener to $inner (inner iframe body).`);
    $inner.on('beforeinput', (evt) => {
      const deleteLogPrefix = '[ep_data_tables:beforeinputDeleteHandler]';
      // log(`${deleteLogPrefix} BEFOREINPUT EVENT TRIGGERED. inputType: "${evt.originalEvent.inputType}", event object:`, evt);

      // Only intercept deletion-related input events
      if (!evt.originalEvent.inputType || !evt.originalEvent.inputType.startsWith('delete')) {
        // log(`${deleteLogPrefix} Not a deletion event (inputType: "${evt.originalEvent.inputType}"). Allowing default.`);
        return; // Allow default for non-delete events
      }

      // log(`${deleteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        // log(`${deleteLogPrefix} WARNING: Could not get representation or selection. Allowing default delete.`);
        console.warn(`${deleteLogPrefix} Could not get rep or selStart.`);
        return; // Allow default
      }
      // log(`${deleteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      // log(`${deleteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      // Check if there's actually a selection to delete
      if (selStart[0] === selEnd[0] && selStart[1] === selEnd[1]) {
        // log(`${deleteLogPrefix} No selection to delete. Allowing default delete.`);
        return; // Allow default - nothing to delete
      }

      // Check if selection spans multiple lines
      if (selStart[0] !== selEnd[0]) {
        // log(`${deleteLogPrefix} WARNING: Selection spans multiple lines. Preventing delete to protect table structure.`);
        evt.preventDefault();
        return;
      }

      // log(`${deleteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (lineAttrString) {
        // Fast-path: attribute exists – parse it.
        try {
          tableMetadata = JSON.parse(lineAttrString);
        } catch {}
      }

      if (!tableMetadata) {
        // Fallback for block-styled rows – reconstruct via DOM helper.
        tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      }

      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        // log(`${deleteLogPrefix} Line ${lineNum} is NOT a recognised table line. Allowing default delete.`);
        return; // Not a table line
      }

      // log(`${deleteLogPrefix} Line ${lineNum} IS a table line. Metadata:`, tableMetadata);

      // Validate selection is within cell boundaries
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
        selStart[1] = cellStartCol;          // clamp
      }
      
      if (wouldClampEnd) {
        console.log(`[ep_data_tables:beforeinput-delete] CLAMPING delete selection end from ${selEnd[1]} to ${cellEndCol}`);
        selEnd[1] = cellEndCol;              // clamp
      }
      
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        // log(`${deleteLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing delete to protect table structure.`);
        evt.preventDefault();
        return;
      }

      // If we reach here, the selection is entirely within a single cell - intercept delete and preserve table structure
      // log(`${deleteLogPrefix} Selection is entirely within cell ${targetCellIndex}. Intercepting delete to preserve table structure.`);
      evt.preventDefault();

      try {
        // No clipboard operations needed for delete - just perform the deletion within the cell using ace operations
        // log(`${deleteLogPrefix} Performing deletion via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${deleteLogPrefix}[ace_callWithAceOps]`;
          // log(`${callAceLogPrefix} Entered ace_callWithAce for delete operations. selStart:`, selStart, `selEnd:`, selEnd);
          
          // log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to delete selected text.`);
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');
          // log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

          // --- Ensure cell is not left empty (zero-length) ---
          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
            // log(`${callAceLogPrefix} Cell ${targetCellIndex} became empty after delete – inserting single space to preserve structure.`);
            const insertPos = [lineNum, selStart[1]]; // Start of the now-empty cell
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            // NEW – give the placeholder its cell attribute back
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

          // Determine new caret position – one char forward if we inserted a space
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

    // *** DRAG AND DROP EVENT LISTENERS ***
    // log(`${callWithAceLogPrefix} Attaching drag and drop event listeners to $inner (inner iframe body).`);
    
    // Prevent drops that could damage table structure
    $inner.on('drop', (evt) => {
      const dropLogPrefix = '[ep_data_tables:dropHandler]';
      // log(`${dropLogPrefix} DROP EVENT TRIGGERED. Event object:`, evt);

      // log(`${dropLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        // log(`${dropLogPrefix} WARNING: Could not get representation or selection. Allowing default drop.`);
        return; // Allow default
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];
      // log(`${dropLogPrefix} Current line number: ${lineNum}.`);

      // Check if we're dropping onto a table line
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

    // Also prevent dragover to ensure drop events are properly handled
    $inner.on('dragover', (evt) => {
      const dragLogPrefix = '[ep_data_tables:dragoverHandler]';
      
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        return; // Allow default
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];

      // Check if we're dragging over a table line
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

    // *** EXISTING PASTE LISTENER ***
    // log(`${callWithAceLogPrefix} Attaching paste event listener to $inner (inner iframe body).`);
    $inner.on('paste', (evt) => {
      const pasteLogPrefix = '[ep_data_tables:pasteHandler]';
      // log(`${pasteLogPrefix} PASTE EVENT TRIGGERED. Event object:`, evt);

      // log(`${pasteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        // log(`${pasteLogPrefix} WARNING: Could not get representation or selection. Allowing default paste.`);
        console.warn(`${pasteLogPrefix} Could not get rep or selStart.`);
        return; // Allow default
      }
      // log(`${pasteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      // log(`${pasteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      // NEW: Check if selection spans multiple lines
      if (selStart[0] !== selEnd[0]) {
        // log(`${pasteLogPrefix} WARNING: Selection spans multiple lines. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

      // log(`${pasteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (!lineAttrString) {
        // Block-styled row? Reconstruct metadata from the DOM.
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
        return; // Not a table line
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
          return; // Allow default
        }
        // log(`${pasteLogPrefix} Table metadata validated successfully: tblId=${tableMetadata.tblId}, row=${tableMetadata.row}, cols=${tableMetadata.cols}.`);
      } catch(e) {
        console.error(`${pasteLogPrefix} ERROR parsing table metadata for line ${lineNum}:`, e);
        // log(`${pasteLogPrefix} Metadata parse error. Allowing default paste. Error details:`, { message: e.message, stack: e.stack });
        return; // Allow default
      }

      // NEW: Validate selection is within cell boundaries
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
        selEnd[1] = cellEndCol;              // clamp
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
        return; // Allow default
      }
      // log(`${pasteLogPrefix} Clipboard data object obtained:`, clipboardData);

      // Allow default handling (so ep_hyperlinked_text plugin can process) if rich HTML is present
      const types = clipboardData.types || [];
      if (types.includes('text/html') && clipboardData.getData('text/html')) {
        // log(`${pasteLogPrefix} Detected text/html in clipboard – deferring to other plugins and default paste.`);
        return; // Do not intercept
      }

      // log(`${pasteLogPrefix} Getting 'text/plain' from clipboard.`);
      const pastedTextRaw = clipboardData.getData('text/plain');
      // log(`${pasteLogPrefix} Pasted text raw: "${pastedTextRaw}" (Type: ${typeof pastedTextRaw})`);

      // ENHANCED: More thorough sanitization of pasted content
      let pastedText = pastedTextRaw
        .replace(/(\r\n|\n|\r)/gm, " ") // Replace newlines with space
        .replace(new RegExp(DELIMITER, 'g'), ' ') // Strip our internal delimiter
        .replace(/\t/g, " ") // Replace tabs with space
        .replace(/\s+/g, " ") // Normalize whitespace
        .trim(); // Trim leading/trailing whitespace

      // log(`${pasteLogPrefix} Pasted text after sanitization: "${pastedText}"`);

      if (typeof pastedText !== 'string' || pastedText.length === 0) {
        // log(`${pasteLogPrefix} No plain text in clipboard or text is empty (after sanitization). Allowing default paste.`);
        const types = clipboardData.types;
        // log(`${pasteLogPrefix} Clipboard types available:`, types);
        if (types && types.includes('text/html')) {
            // log(`${pasteLogPrefix} Clipboard also contains HTML:`, clipboardData.getData('text/html'));
        }
        return; // Allow default if no plain text
      }
      // log(`${pasteLogPrefix} Plain text obtained from clipboard: "${pastedText}". Length: ${pastedText.length}.`);

      // NEW: Check if paste would exceed cell boundaries
      const currentCellText = cells[targetCellIndex] || '';
      const selectionLength = selEnd[1] - selStart[1];
      const newCellLength = currentCellText.length - selectionLength + pastedText.length;
      
      // Soft safety-valve: Etherpad can technically handle very long lines but
      // extremely large cells slow down rendering.  8 000 chars ≈ five classic
      // 'Lorem Ipsum' paragraphs and feels like a reasonable upper bound while
      // still letting users paste substantive text.  Increase/decrease as you
      // see fit or set to `Infinity` to remove the cap entirely.
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
      // Prevent other plugins from handling the same paste event once we
      // have intercepted it inside a table cell.
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

            // Update stored click/caret info
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

    // *** NEW: Column resize listeners ***
    // log(`${callWithAceLogPrefix} Attaching column resize listeners...`);
    
    // Get the iframe documents for proper event delegation
    const $iframeOuter = $('iframe[name="ace_outer"]');
    const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
    const innerDoc = $iframeInner.contents();
    const outerDoc = $iframeOuter.contents();
    
    // log(`${callWithAceLogPrefix} Found iframe documents: outer=${outerDoc.length}, inner=${innerDoc.length}`);
    
    // Mousedown on resize handles
    $inner.on('mousedown', '.ep-data_tables-resize-handle', (evt) => {
      const resizeLogPrefix = '[ep_data_tables:resizeMousedown]';
      // log(`${resizeLogPrefix} Resize handle mousedown detected`);
      
      // Only handle left mouse button clicks
      if (evt.button !== 0) {
        // log(`${resizeLogPrefix} Ignoring non-left mouse button: ${evt.button}`);
        return;
      }
      
      // Check if this is related to an image element to avoid conflicts
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
        // Get table metadata
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
                
                // DEBUG: Verify global state is set
                // log(`${resizeLogPrefix} Global resize state: isResizing=${isResizing}, targetTable=${!!resizeTargetTable}, targetColumn=${resizeTargetColumn}`);
              } else {
                // log(`${resizeLogPrefix} Table ID mismatch: ${metadata.tblId} vs ${tblId}`);
              }
            } else {
              // log(`${resizeLogPrefix} No table metadata found for line ${lineNum}, trying DOM reconstruction...`);
              
              // Fallback: Reconstruct metadata from DOM (same logic as ace_doDatatableOptions)
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
                        // Extract column widths from DOM cells
                        const columnWidths = [];
                        domCells.forEach(cell => {
                          const style = cell.getAttribute('style') || '';
                          const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                          if (widthMatch) {
                            columnWidths.push(parseFloat(widthMatch[1]));
                          } else {
                            // Fallback to equal distribution if no width found
                            columnWidths.push(100 / domCells.length);
                          }
                        });
                        
                        // Reconstruct metadata from DOM with preserved column widths
                        const reconstructedMetadata = {
                          tblId: domTblId,
                          row: parseInt(domRow, 10),
                          cols: domCells.length,
                          columnWidths: columnWidths
                        };
                        // log(`${resizeLogPrefix} Reconstructed metadata from DOM:`, reconstructedMetadata);
                        
                        startColumnResize(table, columnIndex, evt.clientX, reconstructedMetadata, lineNum);
                        // log(`${resizeLogPrefix} Started resize for column ${columnIndex} using reconstructed metadata`);
                        
                        // DEBUG: Verify global state is set
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
    
    // Enhanced mousemove and mouseup handlers - attach to multiple contexts for better coverage
    const setupGlobalHandlers = () => {
      const mouseupLogPrefix = '[ep_data_tables:resizeMouseup]';
      const mousemoveLogPrefix = '[ep_data_tables:resizeMousemove]';
      
      // Mousemove handler
      const handleMousemove = (evt) => {
        if (isResizing) {
          evt.preventDefault();
          updateColumnResize(evt.clientX);
        }
      };
      
      // Mouseup handler with enhanced debugging
      const handleMouseup = (evt) => {
        // log(`${mouseupLogPrefix} Mouseup detected on ${evt.target.tagName || 'unknown'}. isResizing: ${isResizing}`);
        
        if (isResizing) {
          // log(`${mouseupLogPrefix} Processing resize completion...`);
          evt.preventDefault();
          evt.stopPropagation();
          
          // Add a small delay to ensure all DOM updates are complete
          setTimeout(() => {
            // log(`${mouseupLogPrefix} Executing finishColumnResize after delay...`);
            finishColumnResize(ed, docManager);
            // log(`${mouseupLogPrefix} Resize completion finished.`);
          }, 50);
        } else {
          // log(`${mouseupLogPrefix} Not in resize mode, ignoring mouseup.`);
        }
      };
      
      // Attach to multiple contexts to ensure we catch the event
      // log(`${callWithAceLogPrefix} Attaching global mousemove/mouseup handlers to multiple contexts...`);
      
      // 1. Main document (outside iframes)
      $(document).on('mousemove', handleMousemove);
      $(document).on('mouseup', handleMouseup);
      // log(`${callWithAceLogPrefix} Attached to main document`);
      
      // 2. Outer iframe document  
      if (outerDoc.length > 0) {
        outerDoc.on('mousemove', handleMousemove);
        outerDoc.on('mouseup', handleMouseup);
        // log(`${callWithAceLogPrefix} Attached to outer iframe document`);
      }
      
      // 3. Inner iframe document
      if (innerDoc.length > 0) {
        innerDoc.on('mousemove', handleMousemove);
        innerDoc.on('mouseup', handleMouseup);
        // log(`${callWithAceLogPrefix} Attached to inner iframe document`);
      }
      
      // 4. Inner iframe body (the editing area)
      $inner.on('mousemove', handleMousemove);
      $inner.on('mouseup', handleMouseup);
      // log(`${callWithAceLogPrefix} Attached to inner iframe body`);
      
      // 5. Add a failsafe - listen for any mouse events during resize
      const failsafeMouseup = (evt) => {
        if (isResizing) {
          // log(`${mouseupLogPrefix} FAILSAFE: Detected mouse event during resize: ${evt.type}`);
          if (evt.type === 'mouseup' || evt.type === 'mousedown' || evt.type === 'click') {
            // log(`${mouseupLogPrefix} FAILSAFE: Triggering resize completion due to ${evt.type}`);
            setTimeout(() => {
              if (isResizing) { // Double-check we're still resizing
                finishColumnResize(ed, docManager);
              }
            }, 100);
          }
        }
      };
      
      // Attach failsafe to main document with capture=true to catch events early
      document.addEventListener('mouseup', failsafeMouseup, true);
      document.addEventListener('mousedown', failsafeMouseup, true);
      document.addEventListener('click', failsafeMouseup, true);
      // log(`${callWithAceLogPrefix} Attached failsafe event handlers`);
      
      // *** DRAG PREVENTION FOR TABLE ELEMENTS ***
      const preventTableDrag = (evt) => {
        const target = evt.target;
        // Check if the target is a table element or inside a table
        if (target.tagName === 'TABLE' && target.classList.contains('dataTable') ||
            target.tagName === 'TD' && target.closest('table.dataTable') ||
            target.tagName === 'TR' && target.closest('table.dataTable') ||
            target.tagName === 'TBODY' && target.closest('table.dataTable')) {
          // log('[ep_data_tables:dragPrevention] Preventing drag operation on table element:', target.tagName);
          evt.preventDefault();
          evt.stopPropagation();
          return false;
        }
      };
      
      // Add drag event listeners to prevent table dragging
      $inner.on('dragstart', preventTableDrag);
      $inner.on('drag', preventTableDrag);
      $inner.on('dragend', preventTableDrag);
      // log(`${callWithAceLogPrefix} Attached drag prevention handlers to inner body`);
    };
    
    // Setup the global handlers
    setupGlobalHandlers();
    
    // log(`${callWithAceLogPrefix} Column resize listeners attached successfully.`);

  }, 'tablePasteAndResizeListeners', true);
  // log(`${logPrefix} ace_callWithAce for listeners setup completed.`);

  // Helper function to apply the metadata attribute to a line
  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo, attributeString = null, documentAttributeManager = null) {
    const funcName = 'applyTableLineMetadataAttribute';
    // log(`${logPrefix}:${funcName}: START - Applying METADATA attribute to line ${lineNum}`, {tblId, rowIndex, numCols});

    let finalMetadata;
    
    // If attributeString is provided, check if it contains columnWidths
    if (attributeString) {
      try {
        const providedMetadata = JSON.parse(attributeString);
        if (providedMetadata.columnWidths && Array.isArray(providedMetadata.columnWidths) && providedMetadata.columnWidths.length === numCols) {
          // Already has valid columnWidths, use as-is
          finalMetadata = providedMetadata;
          // log(`${logPrefix}:${funcName}: Using provided metadata with existing columnWidths`);
        } else {
          // Has metadata but missing/invalid columnWidths, extract from DOM
          finalMetadata = providedMetadata;
          // log(`${logPrefix}:${funcName}: Provided metadata missing columnWidths, attempting DOM extraction`);
           }
         } catch (e) {
        // log(`${logPrefix}:${funcName}: Error parsing provided attributeString, will reconstruct:`, e);
        finalMetadata = null;
      }
    }
    
    // If we don't have complete metadata or need to extract columnWidths
    if (!finalMetadata || !finalMetadata.columnWidths) {
      let columnWidths = null;
      
      // Try to extract existing column widths from DOM if available
      try {
        const lineEntry = rep.lines.atIndex(lineNum);
        if (lineEntry && lineEntry.lineNode) {
          const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
          if (tableInDOM) {
            const domTblId = tableInDOM.getAttribute('data-tblId');
            if (domTblId === tblId) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                // Extract column widths from DOM cells
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    // Fallback to equal distribution if no width found
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
      
      // Build final metadata
      finalMetadata = finalMetadata || {
        tblId: tblId,
        row: rowIndex,
        cols: numCols
      };
      
      // Add column widths if we found them
      if (columnWidths && columnWidths.length === numCols) {
        finalMetadata.columnWidths = columnWidths;
      }
    }

    const finalAttributeString = JSON.stringify(finalMetadata);
    // log(`${logPrefix}:${funcName}: Final metadata attribute string: ${finalAttributeString}`);
    
    try {
       // Get current line info
       const lineEntry = rep.lines.atIndex(lineNum);
       if (!lineEntry) {
           // log(`${logPrefix}:${funcName}: ERROR - Could not find line entry for line number ${lineNum}`);
           return;
       }
       const lineLength = Math.max(1, lineEntry.text.length);
       // log(`${logPrefix}:${funcName}: Line ${lineNum} text length: ${lineLength}`);

       // Simple attribute application - just add the tbljson attribute
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

    // --- Phase 1: Prepare Data --- 
    const tblId   = rand();
    // log(`${funcName}: Generated table ID: ${tblId}`);
    const initialCellContent = ' '; // Start with a single space per cell
    const lineTxt = Array.from({ length: cols }).fill(initialCellContent).join(DELIMITER);
    // log(`${funcName}: Constructed initial line text for ${cols} cols: "${lineTxt}"`);
    const block = Array.from({ length: rows }).fill(lineTxt).join('\n') + '\n';
    // log(`${funcName}: Constructed block for ${rows} rows:\n${block}`);

    // Get current selection BEFORE making changes using ace_getRep()
    // log(`${funcName}: Getting current representation and selection...`);
    const currentRepInitial = ed.ace_getRep(); 
    if (!currentRepInitial || !currentRepInitial.selStart || !currentRepInitial.selEnd) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation or selection via ace_getRep(). Aborting.`);
        // log(`${funcName}: END - Error getting initial rep/selection`);
        return;
    }
    const start = currentRepInitial.selStart;
    const end = currentRepInitial.selEnd;
    const initialStartLine = start[0]; // Store the starting line number
    // log(`${funcName}: Current selection from initial rep:`, { start, end });

    // --- Phase 2: Insert Text Block --- 
    // log(`${funcName}: Phase 2 - Inserting text block...`);
    ed.ace_performDocumentReplaceRange(start, end, block);
    // log(`${funcName}: Inserted block of delimited text lines.`);
    // log(`${funcName}: Requesting text sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20); // Sync text insertion
    // log(`${funcName}: Text sync requested.`);

    // --- Phase 3: Apply Metadata Attributes --- 
    // log(`${funcName}: Phase 3 - Applying metadata attributes to ${rows} inserted lines...`);
    // Need rep to be updated after text insertion to apply attributes correctly
    const currentRep = ed.ace_getRep(); // Get potentially updated rep
    if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get updated rep after text insertion. Cannot apply attributes reliably.`);
        // log(`${funcName}: END - Error getting updated rep`);
        // Maybe attempt to continue without rep? Risky.
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

      // Apply cell-specific attributes to trigger authorship span splitting
      for (let c = 0; c < cols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) { // Only apply to non-empty cells
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

      // Call the module-level helper, passing necessary context (currentRep, ed)
      // Note: documentAttributeManager not available in this context for new table creation
      applyTableLineMetadataAttribute(lineNumToApply, tblId, r, cols, currentRep, ed, null, null); 
    }
    // log(`${funcName}: Finished applying metadata attributes.`);
    // log(`${funcName}: Requesting attribute sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20); // Final sync after attributes
    // log(`${funcName}: Attribute sync requested.`);

    // --- Phase 4: Set Caret Position --- 
    // log(`${funcName}: Phase 4 - Setting final caret position...`);
    const finalCaretLine = initialStartLine + rows; // Line number after the last inserted row
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
    
    // Get the last clicked cell info to determine which table to operate on
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
      // Get current representation and document manager
      const currentRep = ed.ace_getRep();
      if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation.`);
        return;
      }
      
      // Use the stored documentAttributeManager reference
      const docManager = ed.ep_data_tables_docManager;
      if (!docManager) {
        console.error(`[ep_data_tables] ${funcName}: Could not get document attribute manager from stored reference.`);
        return;
      }
      
      // log(`${funcName}: Successfully obtained documentAttributeManager from stored reference.`);
      
      // Find all lines that belong to this table
      const tableLines = [];
      const totalLines = currentRep.lines.length();
      
      for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
        try {
          // Use the same robust approach as acePostWriteDomLineHTML to find nested tables
          let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);
          
          // If no attribute found directly, check if there's a table in the DOM even though attribute is missing
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
                    // Reconstruct metadata from DOM
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
      
      // Sort by row number to ensure correct order
      tableLines.sort((a, b) => a.row - b.row);
      // log(`${funcName}: Found ${tableLines.length} table lines`);
      
      // Determine table dimensions and target indices with robust matching
      const numRows = tableLines.length;
      const numCols = tableLines[0].cols;
      
      // More robust way to find the target row - match by both line number AND row metadata
      let targetRowIndex = -1;
      
      // First try to match by line number
      targetRowIndex = tableLines.findIndex(line => line.lineIndex === lastClick.lineNum);
      
      // If that fails, try to match by finding the row that contains the clicked table
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
      
      // If still not found, default to first row but log the issue
      if (targetRowIndex === -1) {
        // log(`${funcName}: Warning: Could not find target row, defaulting to row 0`);
        targetRowIndex = 0;
      }
      
      const targetColIndex = lastClick.cellIndex || 0;
      
      // log(`${funcName}: Table dimensions: ${numRows} rows x ${numCols} cols. Target: row ${targetRowIndex}, col ${targetColIndex}`);
      
      // Perform table operations with both text and metadata updates
      let newNumCols = numCols;
      let success = false;
      
      switch (action) {
        case 'addTblRowA': // Insert row above
          // log(`${funcName}: Inserting row above row ${targetRowIndex}`);
          success = addTableRowAboveWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;
          
        case 'addTblRowB': // Insert row below
          // log(`${funcName}: Inserting row below row ${targetRowIndex}`);
          success = addTableRowBelowWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;
          
        case 'addTblColL': // Insert column left
          // log(`${funcName}: Inserting column left of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnLeftWithText(tableLines, targetColIndex, ed, docManager);
          break;
          
        case 'addTblColR': // Insert column right
          // log(`${funcName}: Inserting column right of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnRightWithText(tableLines, targetColIndex, ed, docManager);
          break;
          
        case 'delTblRow': // Delete row
          // Show confirmation prompt for row deletion
          const rowConfirmMessage = `Are you sure you want to delete Row ${targetRowIndex + 1} and all content within?`;
          if (!confirm(rowConfirmMessage)) {
            // log(`${funcName}: Row deletion cancelled by user`);
            return;
          }
          // log(`${funcName}: Deleting row ${targetRowIndex}`);
          success = deleteTableRowWithText(tableLines, targetRowIndex, ed, docManager);
          break;
          
        case 'delTblCol': // Delete column
          // Show confirmation prompt for column deletion
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

  // Helper functions for table operations with text updates
  function addTableRowAboveWithText(tableLines, targetRowIndex, numCols, tblId, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      const newLineText = Array.from({ length: numCols }).fill(' ').join(DELIMITER);
      const insertLineIndex = targetLine.lineIndex;
      
      // Insert new line in text
      editorInfo.ace_performDocumentReplaceRange([insertLineIndex, 0], [insertLineIndex, 0], newLineText + '\n');
      
      // Apply cell-specific attributes to the new row for authorship
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
      
      // Preserve column widths from existing metadata or extract from DOM
      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        // Extract from DOM for block-styled rows
        try {
          const rep = editorInfo.ace_getRep();
          const lineEntry = rep.lines.atIndex(targetLine.lineIndex + 1); // +1 because we inserted a line
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
      
      // Update metadata for all subsequent rows
      for (let i = targetRowIndex; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1; // +1 because we inserted a line
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };
        
        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Apply metadata to the new row
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
      
      // Insert new line in text
      editorInfo.ace_performDocumentReplaceRange([insertLineIndex, 0], [insertLineIndex, 0], newLineText + '\n');
      
      // Apply cell-specific attributes to the new row for authorship
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
      
      // Preserve column widths from existing metadata or extract from DOM
      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        // Extract from DOM for block-styled rows
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
      
      // Update metadata for all subsequent rows
      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1; // +1 because we inserted a line
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };
        
        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Apply metadata to the new row
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
    try {
      // Process each line individually like table creation does
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);
        
        // Calculate the exact insertion position - stop BEFORE the target column's delimiter
        let insertPos = 0;
        for (let i = 0; i < targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0) + DELIMITER.length;
        }
        
        // Insert blank cell then delimiter (BLANK + separator)
        const textToInsert = ' ' + DELIMITER;
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];
        
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);
        
        // Immediately apply authorship attributes like table creation does
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (lineEntry) {
          const newLineText = lineEntry.text || '';
          const newCells = newLineText.split(DELIMITER);
          let offset = 0;
          
          // Apply cell-specific attributes to ALL cells (like table creation)
          for (let c = 0; c < tableLine.cols + 1; c++) { // +1 for the new column
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) { // Only apply to non-empty cells
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
              // log(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL} attribute to Line ${tableLine.lineIndex} Col ${c} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }
        
        // Reset all column widths to equal distribution when adding a column
        // This avoids complex width calculations and ensures robust behavior
        const newColCount = tableLine.cols + 1;
        const equalWidth = 100 / newColCount;
        const normalizedWidths = Array(newColCount).fill(equalWidth);
       // log(`[ep_data_tables] addTableColumnLeft: Reset all column widths to equal distribution: ${newColCount} columns at ${equalWidth.toFixed(1)}% each`);
        
        // Apply updated metadata
        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols + 1, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols + 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Final sync
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding column left with text:', e);
      return false;
    }
  }
  
  function addTableColumnRightWithText(tableLines, targetColIndex, editorInfo, docManager) {
    const funcName = 'addTableColumnRightWithText';
    try {
      // Process each line individually like table creation does
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);
        
        // Calculate the exact insertion position - stop BEFORE the target column's trailing delimiter
        let insertPos = 0;
        for (let i = 0; i <= targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0);
          if (i < targetColIndex) insertPos += DELIMITER.length;
        }
        
        // Insert delimiter then blank cell (separator + BLANK)
        const textToInsert = DELIMITER + ' ';
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];
        
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);
        
        // Immediately apply authorship attributes like table creation does
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (lineEntry) {
          const newLineText = lineEntry.text || '';
          const newCells = newLineText.split(DELIMITER);
          let offset = 0;
          
          // Apply cell-specific attributes to ALL cells (like table creation)
          for (let c = 0; c < tableLine.cols + 1; c++) { // +1 for the new column
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) { // Only apply to non-empty cells
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
              // log(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL} attribute to Line ${tableLine.lineIndex} Col ${c} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }
        
        // Reset all column widths to equal distribution when adding a column
        // This avoids complex width calculations and ensures robust behavior
        const newColCount = tableLine.cols + 1;
        const equalWidth = 100 / newColCount;
        const normalizedWidths = Array(newColCount).fill(equalWidth);
        // log(`[ep_data_tables] addTableColumnRight: Reset all column widths to equal distribution: ${newColCount} columns at ${equalWidth.toFixed(1)}% each`);
        
        // Apply updated metadata
        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols + 1, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols + 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Final sync
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding column right with text:', e);
      return false;
    }
  }
  
  function deleteTableRowWithText(tableLines, targetRowIndex, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      
      // Special handling for deleting the first row (row index 0)
      // Insert a blank line to prevent the table from getting stuck at line 1
      if (targetRowIndex === 0) {
        // log('[ep_data_tables] Deleting first row (row 0) - inserting blank line to prevent table from getting stuck');
        const insertStart = [targetLine.lineIndex, 0];
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertStart, '\n');
        
        // Now delete the table line (which is now at lineIndex + 1)
        const deleteStart = [targetLine.lineIndex + 1, 0];
        const deleteEnd = [targetLine.lineIndex + 2, 0];
        editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      } else {
        // Delete the entire line normally
      const deleteStart = [targetLine.lineIndex, 0];
      const deleteEnd = [targetLine.lineIndex + 1, 0];
      editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      }
      
      // Extract column widths from target line before deletion for preserving in remaining rows
      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        // Extract from DOM for block-styled rows
        try {
          const rep = editorInfo.ace_getRep();
          // Check any remaining table line for column widths
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
      
      // Update metadata for all subsequent rows
      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex - 1; // -1 because we deleted a line
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
    try {
      // Update text content for all table lines using precise character deletion
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);
        
        if (targetColIndex >= cells.length) {
          // log(`[ep_data_tables] Warning: Target column ${targetColIndex} doesn't exist in line with ${cells.length} columns`);
          continue;
        }
        
        // Calculate the exact character range to delete
        let deleteStart = 0;
        let deleteEnd = 0;
        
        // Calculate start position
        for (let i = 0; i < targetColIndex; i++) {
          deleteStart += (cells[i]?.length ?? 0) + DELIMITER.length;
        }
        
        // Calculate end position
        deleteEnd = deleteStart + (cells[targetColIndex]?.length ?? 0);
        
        // Include the delimiter in deletion
        if (targetColIndex === 0 && cells.length > 1) {
          // If deleting first column, include the delimiter after it
          deleteEnd += DELIMITER.length;
        } else if (targetColIndex > 0) {
          // If deleting any other column, include the delimiter before it
          deleteStart -= DELIMITER.length;
        }
        
        // log(`[ep_data_tables] Deleting column ${targetColIndex} from line ${tableLine.lineIndex}: chars ${deleteStart}-${deleteEnd} from "${lineText}"`);
        
        // Perform the precise deletion
        const rangeStart = [tableLine.lineIndex, deleteStart];
        const rangeEnd = [tableLine.lineIndex, deleteEnd];
        
        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, '');
        
        // Reset all column widths to equal distribution when deleting a column
        // This avoids complex width calculations and ensures robust behavior
        const newColCount = tableLine.cols - 1;
        if (newColCount > 0) {
          const equalWidth = 100 / newColCount;
          const normalizedWidths = Array(newColCount).fill(equalWidth);
          // log(`[ep_data_tables] deleteTableColumn: Reset all column widths to equal distribution: ${newColCount} columns at ${equalWidth.toFixed(1)}% each`);
        
        // Update metadata
          const newMetadata = { ...tableLine.metadata, cols: newColCount, columnWidths: normalizedWidths };
          applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, newColCount, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
        }
      }
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error deleting column with text:', e);
      return false;
    }
  }
  
  // ... existing code ...

  // log('aceInitialized: END - helpers defined.');
};

// ───────────────────── required no‑op stubs ─────────────────────
exports.aceStartLineAndCharForPoint = () => { return undefined; };
exports.aceEndLineAndCharForPoint   = () => { return undefined; };

// NEW: Style protection for table cells
exports.aceSetAuthorStyle = (hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceSetAuthorStyle]';
  // log(`${logPrefix} START`, { hook, ctx });

  // If no selection or no style to apply, allow default
  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd || !ctx.key) {
    // log(`${logPrefix} No selection or style key. Allowing default.`);
    return;
  }

  // Check if selection is within a table
  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];
  
  // If selection spans multiple lines, prevent style application
  if (startLine !== endLine) {
    // log(`${logPrefix} Selection spans multiple lines. Preventing style application to protect table structure.`);
    return false;
  }

  // Check if the line is a table line
  const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(startLine, ATTR_TABLE_JSON);
  if (!lineAttrString) {
    // log(`${logPrefix} Line ${startLine} is not a table line. Allowing default style application.`);
    return;
  }

  // List of styles that could break table structure
  const BLOCKED_STYLES = [
    'list', 'listType', 'indent', 'align', 'heading', 'code', 'quote',
    'horizontalrule', 'pagebreak', 'linebreak', 'clear'
  ];

  if (BLOCKED_STYLES.includes(ctx.key)) {
    // log(`${logPrefix} Blocked potentially harmful style '${ctx.key}' from being applied to table cell.`);
    return false;
  }

  // For allowed styles, ensure they only apply within cell boundaries
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

    // Find which cells the selection spans
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

    // If selection spans multiple cells, prevent style application
    if (selectionStartCell !== selectionEndCell) {
      // log(`${logPrefix} Selection spans multiple cells. Preventing style application to protect table structure.`);
      return false;
    }

    // If selection includes cell delimiters, prevent style application
    const cellStartCol = cells.slice(0, selectionStartCell).reduce((acc, cell) => acc + cell.length + DELIMITER.length, 0);
    const cellEndCol = cellStartCol + cells[selectionStartCell].length;
    
    if (ctx.rep.selStart[1] <= cellStartCol || ctx.rep.selEnd[1] >= cellEndCol) {
      // log(`${logPrefix} Selection includes cell delimiters. Preventing style application to protect table structure.`);
      return false;
    }

    // log(`${logPrefix} Style '${ctx.key}' allowed within cell boundaries.`);
    return; // Allow the style to be applied
  } catch (e) {
    console.error(`${logPrefix} Error processing style application:`, e);
    // log(`${logPrefix} Error details:`, { message: e.message, stack: e.stack });
    return false; // Prevent style application on error
  }
};

exports.aceEditorCSS                = () => { 
  // Path relative to Etherpad's static/plugins/ directory
  // Format should be: pluginName/path/to/file.css
  return ['ep_data_tables/static/css/datatables-editor.css', 'ep_data_tables/static/css/caret.css'];
};

// Register TABLE as a block element, hoping it influences rendering behavior
exports.aceRegisterBlockElements = () => ['table'];

// NEW: Column resize helper functions (adapted from images plugin)
const startColumnResize = (table, columnIndex, startX, metadata, lineNum) => {
  const funcName = 'startColumnResize';
  // log(`${funcName}: Starting resize for column ${columnIndex}`);
  
  isResizing = true;
  resizeStartX = startX;
  resizeCurrentX = startX; // Initialize current position
  resizeTargetTable = table;
  resizeTargetColumn = columnIndex;
  resizeTableMetadata = metadata;
  resizeLineNum = lineNum;
  
  // Get current column widths
  const numCols = metadata.cols;
  resizeOriginalWidths = metadata.columnWidths ? [...metadata.columnWidths] : Array(numCols).fill(100 / numCols);
  
  // log(`${funcName}: Original widths:`, resizeOriginalWidths);
  
  // Create visual overlay instead of modifying table directly
  createResizeOverlay(table, columnIndex);
  
  // Prevent text selection during resize
  document.body.style.userSelect = 'none';
  document.body.style.webkitUserSelect = 'none';
  document.body.style.mozUserSelect = 'none';
  document.body.style.msUserSelect = 'none';
};

const createResizeOverlay = (table, columnIndex) => {
  // Create a visual overlay that shows resize feedback using the same positioning logic as image plugin
  if (resizeOverlay) {
    resizeOverlay.remove();
  }
  
  // Get all the necessary container references like image plugin
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

  // Find all tables that belong to the same table (same tblId)
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
  
  // Calculate the bounding box that encompasses all table rows
  let minTop = Infinity;
  let maxBottom = -Infinity;
  let tableLeft = 0;
  let tableWidth = 0;
  
  Array.from(allTableRows).forEach((tableRow, index) => {
    const rect = tableRow.getBoundingClientRect();
    minTop = Math.min(minTop, rect.top);
    maxBottom = Math.max(maxBottom, rect.bottom);
    
    // Use the first table row for horizontal positioning
    if (index === 0) {
      tableLeft = rect.left;
      tableWidth = rect.width;
    }
  });
  
  const totalTableHeight = maxBottom - minTop;
  
  // log(`createResizeOverlay: Found ${allTableRows.length} table rows, total height: ${totalTableHeight}px`);
  
  // Calculate positioning using the same method as image plugin
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
  
  // Get table position relative to inner body using the full table bounds
  const tableTopRelInner = minTop - innerBodyRect.top + scrollTopInner;
  const tableLeftRelInner = tableLeft - innerBodyRect.left + scrollLeftInner;
  
  // Calculate position in outer body coordinates (like image plugin)
  const innerFrameTopRelOuter = innerIframeRect.top - outerBodyRect.top + scrollTopOuter;
  const innerFrameLeftRelOuter = innerIframeRect.left - outerBodyRect.left + scrollLeftOuter;
  
  const overlayTopOuter = innerFrameTopRelOuter + tableTopRelInner;
  const overlayLeftOuter = innerFrameLeftRelOuter + tableLeftRelInner;
  
  // Apply padding and manual offsets like image plugin
  const outerPadding = window.getComputedStyle(padOuter[0]);
  const outerPaddingTop = parseFloat(outerPadding.paddingTop) || 0;
  const outerPaddingLeft = parseFloat(outerPadding.paddingLeft) || 0;
  
  // Use the same manual offsets as image plugin
  const MANUAL_OFFSET_TOP = 6;
  const MANUAL_OFFSET_LEFT = 39;
  
  const finalOverlayTop = overlayTopOuter + outerPaddingTop + MANUAL_OFFSET_TOP;
  const finalOverlayLeft = overlayLeftOuter + outerPaddingLeft + MANUAL_OFFSET_LEFT;
  
  // Calculate the position for the blue line at the right edge of the target column
  const tds = table.querySelectorAll('td');
  const tds_array = Array.from(tds);
  let linePosition = 0;
  
  if (columnIndex < tds_array.length) {
    const currentTd = tds_array[columnIndex];
    const currentTdRect = currentTd.getBoundingClientRect();
    const currentRelativeLeft = currentTdRect.left - tableLeft; // Use tableLeft instead of tableRect.left
    const currentWidth = currentTdRect.width;
    linePosition = currentRelativeLeft + currentWidth;
  }
  
  // Create overlay container (invisible background) that spans the entire table height
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
  
  // Create the blue vertical line (Google Docs style) spanning the full table height
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
  
  // Append to outer body like image plugin does with its outline
  padOuter.append(resizeOverlay);
  
  // log('createResizeOverlay: Created Google Docs style blue line overlay spanning entire table height');
};

const updateColumnResize = (currentX) => {
  if (!isResizing || !resizeTargetTable || !resizeOverlay) return;
  
  resizeCurrentX = currentX; // Store current position for finishColumnResize
  const deltaX = currentX - resizeStartX;
  
  // Get the table width from the first row for percentage calculation
  const tblId = resizeTargetTable.getAttribute('data-tblId');
  if (!tblId) return;
  
  // Find the first table row to get consistent width measurements
  const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
  const innerDocBody = $innerIframe.contents().find('body')[0];
  const firstTableRow = innerDocBody.querySelector(`table.dataTable[data-tblId="${tblId}"]`);
  
  if (!firstTableRow) return;
  
  const tableRect = firstTableRow.getBoundingClientRect();
  const deltaPercent = (deltaX / tableRect.width) * 100;
  
  // Calculate new widths for final application
  const newWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;
  
  if (nextColumn < newWidths.length) {
    const transfer = Math.min(deltaPercent, newWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(newWidths[currentColumn] - 5));
    
    newWidths[currentColumn] += actualTransfer;
    newWidths[nextColumn] -= actualTransfer;
    
    // Update the blue line position to show the new column boundary
    const resizeLine = resizeOverlay.querySelector('.resize-line');
    if (resizeLine) {
      // Calculate new position based on the updated column width
      const newColumnWidth = (newWidths[currentColumn] / 100) * tableRect.width;
      
      // Find the original left position relative to the first table row
      const tds = firstTableRow.querySelectorAll('td');
      const tds_array = Array.from(tds);
      
      if (currentColumn < tds_array.length) {
        const currentTd = tds_array[currentColumn];
        const currentTdRect = currentTd.getBoundingClientRect();
        const currentRelativeLeft = currentTdRect.left - tableRect.left;
        
        // New line position is the original left position plus the new width
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
  
  // Calculate final widths from actual mouse movement
  const tableRect = resizeTargetTable.getBoundingClientRect();
  const deltaX = resizeCurrentX - resizeStartX;
  const deltaPercent = (deltaX / tableRect.width) * 100;
  
  // log(`${funcName}: Mouse moved ${deltaX}px (${deltaPercent.toFixed(1)}%)`);
  
  const finalWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;
  
  if (nextColumn < finalWidths.length) {
    // Transfer width between columns with minimum constraints
    const transfer = Math.min(deltaPercent, finalWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(finalWidths[currentColumn] - 5));
    
    finalWidths[currentColumn] += actualTransfer;
    finalWidths[nextColumn] -= actualTransfer;
    
    // log(`${funcName}: Transferred ${actualTransfer.toFixed(1)}% from column ${nextColumn} to column ${currentColumn}`);
  }
  
  // Normalize widths
  const totalWidth = finalWidths.reduce((sum, width) => sum + width, 0);
  if (totalWidth > 0) {
    finalWidths.forEach((width, index) => {
      finalWidths[index] = (width / totalWidth) * 100;
    });
  }
  
  // log(`${funcName}: Final normalized widths:`, finalWidths.map(w => w.toFixed(1) + '%'));
  
  // Clean up overlay
  if (resizeOverlay) {
    resizeOverlay.remove();
    resizeOverlay = null;
  }
  
  // Clean up global styles
  document.body.style.userSelect = '';
  document.body.style.webkitUserSelect = '';
  document.body.style.mozUserSelect = '';
  document.body.style.msUserSelect = '';
  
  // Set isResizing to false BEFORE making changes
  isResizing = false;
  
  // Apply updated metadata to ALL rows in the table (not just the resized row)
  editorInfo.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = `${funcName}[ace_callWithAce]`;
    // log(`${callWithAceLogPrefix}: Finding and updating all table rows with tblId: ${resizeTableMetadata.tblId}`);
    
    try {
      const rep = ace.ace_getRep();
      if (!rep || !rep.lines) {
        console.error(`${callWithAceLogPrefix}: Invalid rep`);
        return;
      }
      
      // Find all lines that belong to this table
      const tableLines = [];
      const totalLines = rep.lines.length();
      
      for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
        try {
          // Get line metadata to check if it belongs to our table
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
            // Fallback: Check if there's a table in the DOM even though attribute is missing (block styles)
            const lineEntry = rep.lines.atIndex(lineIndex);
            if (lineEntry && lineEntry.lineNode) {
              const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
              if (tableInDOM) {
                const domTblId = tableInDOM.getAttribute('data-tblId');
                const domRow = tableInDOM.getAttribute('data-row');
                if (domTblId === resizeTableMetadata.tblId && domRow !== null) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length > 0) {
                    // Extract column widths from DOM cells
                    const columnWidths = [];
                    domCells.forEach(cell => {
                      const style = cell.getAttribute('style') || '';
                      const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                      if (widthMatch) {
                        columnWidths.push(parseFloat(widthMatch[1]));
                      } else {
                        // Fallback to equal distribution if no width found
                        columnWidths.push(100 / domCells.length);
                      }
                    });
                    
                    // Reconstruct metadata from DOM with preserved column widths
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
          continue; // Skip lines with invalid metadata
        }
      }
      
      // log(`${callWithAceLogPrefix}: Found ${tableLines.length} table lines to update`);
      
      // Update all table lines with new column widths
      for (const tableLine of tableLines) {
        const updatedMetadata = { ...tableLine.metadata, columnWidths: finalWidths };
        const updatedMetadataString = JSON.stringify(updatedMetadata);
        
        // Get the full line range for this table line
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (!lineEntry) {
          console.error(`${callWithAceLogPrefix}: Could not get line entry for line ${tableLine.lineIndex}`);
          continue;
        }
        
        const lineLength = Math.max(1, lineEntry.text.length);
        const rangeStart = [tableLine.lineIndex, 0];
        const rangeEnd = [tableLine.lineIndex, lineLength];
        
        // log(`${callWithAceLogPrefix}: Updating line ${tableLine.lineIndex} (row ${tableLine.metadata.row}) with new column widths`);
        
        // Apply the updated metadata attribute directly
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
  
  // Reset state
  resizeStartX = 0;
  resizeCurrentX = 0;
  resizeTargetTable = null;
  resizeTargetColumn = -1;
  resizeOriginalWidths = [];
  resizeTableMetadata = null;
  resizeLineNum = -1;
  
  // log(`${funcName}: Resize complete - state reset`);
};

// NEW: Undo/Redo protection
exports.aceUndoRedo = (hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceUndoRedo]';
  // log(`${logPrefix} START`, { hook, ctx });

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd) {
    // log(`${logPrefix} No selection or context. Allowing default.`);
    return;
  }

  // Get the affected line range
  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];

  // Check if any affected lines are table lines
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

  // Validate table structure after undo/redo
  try {
    for (const line of tableLines) {
      const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
      if (!lineAttrString) continue;

      const tableMetadata = JSON.parse(lineAttrString);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
        // log(`${logPrefix} Invalid table metadata after undo/redo. Attempting recovery.`);
        // Attempt to recover table structure
        const lineText = ctx.rep.lines.atIndex(line)?.text || '';
        const cells = lineText.split(DELIMITER);
        
        // If we have valid cells, try to reconstruct the table metadata
        if (cells.length > 1) {
          const newMetadata = {
            cols: cells.length,
            rows: 1,
            cells: cells.map((_, i) => ({ col: i, row: 0 }))
          };
          
          // Apply the recovered metadata
          ctx.documentAttributeManager.setAttributeOnLine(line, ATTR_TABLE_JSON, JSON.stringify(newMetadata));
          // log(`${logPrefix} Recovered table structure for line ${line}`);
        } else {
          // If we can't recover, remove the table attribute
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