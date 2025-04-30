/* ep_tables5 – attribute‑based tables (line‑class + PostWrite renderer)
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
const log             = (...m) => console.debug('[ep_tables5:client_hooks]', ...m);
const ZWSP            = '\u200B'; // Zero-Width Space used as placeholder in aceCreateDomLine

// helper for stable random ids
const rand = () => Math.random().toString(36).slice(2, 8);

// encode/decode so JSON can survive as a CSS class token if ever needed
const enc = s => btoa(s).replace(/\+/g, '-').replace(/\//g, '_');
const dec = s => {
    // Revert to simpler decode, assuming enc provides valid padding
    const str = s.replace(/-/g, '+').replace(/_/g, '/');
    try {
        return atob(str);
    } catch (e) {
        console.error('[ep_tables5] Error decoding base64 string:', s, e);
        return null;
    }
};

// Module-level cache for rendered table row nodes, keyed by tblId:rowIndex
const tableRowNodes = {};

// NEW: Module-level state for last clicked cell
let lastClickedCellInfo = null; // { lineNum: number, cellIndex: number, tblId: string }

// Helper function to check if a node is still in the main document
function isNodeInDocument(node) {
    return node && node.isConnected; // Use standard isConnected property
}

// ────────────────────── collectContentPre (DOM → atext) ─────────────────────
exports.collectContentPre = (hook, ctx) => {
  // Log caret position at the start of this hook
  log(`[CaretTrace] collectContentPre: START - Current rep.selStart:`, ctx?.rep?.selStart);

  const node = ctx.dom;
  // Only run for the line DIV itself (or top-level elements passed to collector)
  if (!(node?.classList?.contains('ace-line'))) {
      return; // Let default handlers process children like spans, tds etc.
  }

  log('Client collectContentPre: START for Line Div', { nodeClass: node?.className, nodeId: node?.id });
  const tableNode = node.querySelector('table.dataTable');

  if (tableNode) {
    log('Client collectContentPre: Found table.dataTable within Line Div', node.id);

    // --- Get Table Metadata (tblId, rowIdx) --- 
    let existingRowData = null;
    const lineNum = ctx.rep.lines.indexOfKey(node.id); // Use node ID directly
    log('Client collectContentPre: Line number for attribute lookup:', lineNum);
    if (lineNum !== -1 && ctx.documentAttributeManager) {
         try { 
             const existingAttr = ctx.documentAttributeManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
             if (existingAttr) { existingRowData = JSON.parse(existingAttr); log('Client collectContentPre: Found existing row data from attribute'); }
         } catch (e) { console.error('[ep_tables5] Client collectContentPre: Error parsing existing tbljson attribute', e); }
    }
    let tblId = existingRowData?.tblId || tableNode.getAttribute('data-tblId') || rand();
    let rowIdx = existingRowData?.row;
    if (rowIdx === undefined || rowIdx === null) {
        // Fallback: try to calculate row index based on previous siblings (less reliable)
        let count = 0;
        let sibling = node.previousElementSibling;
        while(sibling) {
            if (sibling.querySelector('table.dataTable')) {
                 count++;
            }
            sibling = sibling.previousElementSibling;
        }
        rowIdx = count; 
        log('Client collectContentPre: Warning - Row index calculated from siblings (fallback):', rowIdx);
    } else {
         log('Client collectContentPre: Using row index from existing attribute');
    }
    tableNode.setAttribute('data-tblId', tblId); 
    // --- End Metadata --- 

    const trNode = tableNode.querySelector('tbody > tr');
    if (!trNode) {
        log('Client collectContentPre: ERROR - Could not find <tr>. Aborting attribute set.');
        return undefined; // Prevent default for line div, but don't set attribute
    }

    // *** Read cell text from DOM SPANS (should reflect collected state) ***
    const cellsData = Array.from(trNode.children).map((td, index) => {
      const innerSpan = td.querySelector('.table-cell-content');
      const text = innerSpan ? innerSpan.textContent : (td.textContent || ' '); // DEBUG: Log intermediate text
      log(`Client collectContentPre (${node.id}): Reading cell ${index}. Span found: ${!!innerSpan}. Text: "${text}". TD innerHTML:`, td.innerHTML); // DEBUG: Detailed cell read log
      return { txt: text };
    });

    const rowObj = { tblId, row: rowIdx, cells: cellsData };
    const attributeString = JSON.stringify(rowObj);
    log('Client collectContentPre: Built rowObj from DOM spans:', rowObj);

    // *** Use lineAttributes to set the attribute for the line ***
    ctx.lineAttributes.push([ATTR_TABLE_JSON, attributeString]);
    log('Client collectContentPre: Pushed tbljson to lineAttributes.');

    // *** SET CANONICAL LINE TEXT: Concatenate cell text ***
    // This ensures ACE's internal representation matches our logical content.
    const canonicalLineText = cellsData.map(cell => cell.txt).join('');
    ctx.state.line = canonicalLineText;
    log(`Client collectContentPre (${node.id}): Set state.line to "${canonicalLineText}"`);

    // Prevent default processing for the line div's children (the table)
    // This is still needed so Etherpad doesn't try to re-collect text from the rendered table.
    log('Client collectContentPre: END - Preventing default collection for line children.');
    return undefined; 
  }

  // log('Client collectContentPre: Not a table line div, allowing default.');
  return; // Allow default collection for non-table lines
};

// ───────────── attribute → span‑class mapping (linestylefilter hook) ─────────
exports.aceAttribsToClasses = (hook, ctx) => {
  if (ctx.key === ATTR_TABLE_JSON) {
    log(`aceAttribsToClasses: Mapping ATTR_TABLE_JSON. Length: ${ctx.value.length}`);
    // Only return the data class
    return [`tbljson-${enc(ctx.value)}`];
  }
  if (ctx.key === ATTR_CELL) {
    // Keep this in case we want cell-specific styling later
    return [`tblCell-${ctx.value}`];
  }
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

// Helper function to build HTML for a single table row (used for initial render)
function buildHtmlTableForRow(rowData) {
  log('buildHtmlTableForRow: START', { rowData });
  // const borderStyle = `border: 1px solid #ccc`; // REMOVED - Handled by CSS
  // Use padding + word-wrap + vertical-align inline for now. Border and min-height are handled by CSS.
  const tdStyle = `padding: 5px 7px; word-wrap:break-word; vertical-align: top;`;

  const cells = Array.isArray(rowData.cells) ? rowData.cells : [];
  const cellsHtml = cells.map(cell =>
    `<td style="${tdStyle}"><span class="table-cell-content">${escapeHtml(cell.txt)}</span></td>`
  ).join('');

  // Add 'dataTable-first-row' class if it's the logical first row (row index 0)
  const firstRowClass = rowData.row === 0 ? ' dataTable-first-row' : '';

  // Removed border-collapse from inline style, rely on CSS
  const tableHtml = `<table class="dataTable${firstRowClass}" data-tblId="${rowData.tblId || ''}" style="width:100%;"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
  log('buildHtmlTableForRow: Generated HTML', { tableHtml }); // Log generated HTML
  return tableHtml;
}

// ───────────────── Populate Table Cells / Render (PostWrite) ──────────────────
exports.acePostWriteDomLineHTML = function (hook_name, args, cb) {
  // Log caret position at the start of this hook
  log(`[CaretTrace] acePostWriteDomLineHTML: START - Node ID: ${args?.node?.id} - Current rep.selStart:`, args?.rep?.selStart);

  // *** Comprehensive Logging at Start *** 
  console.groupCollapsed(`acePostWriteDomLineHTML: START - Node ID: ${args?.node?.id}`);
  log('acePostWriteDomLineHTML: Full args object:', args);
  log('acePostWriteDomLineHTML: Passed node object:', args?.node);
  log('acePostWriteDomLineHTML: Passed node outerHTML:', args?.node?.outerHTML);
  console.groupEnd();
  // ************************************* 

  const node = args.node;
  const nodeId = node?.id;

  // Abort if we don't have a node or node ID
  if (!node || !nodeId) {
      log('acePostWriteDomLineHTML: ERROR - Received invalid node or node without ID. Aborting.');
      return cb();
  }

  let rowData = null;
  let encodedJsonString = null;

  // Find the tbljson-* class (Primary identifier)
  if (node && node.classList) { 
      for (const cls of node.classList) { 
          if (cls.startsWith('tbljson-')) { encodedJsonString = cls.substring(8); break; }
      }
  } 
  if (!encodedJsonString && node && node.children) { 
        for (const child of node.children) {
             if (child.classList) {
                for (const cls of child.classList) {
                  if (cls.startsWith('tbljson-')) { encodedJsonString = cls.substring(8); break; }
              }
          } 
          if (encodedJsonString) break;
      }
  } 
  if (!encodedJsonString) { return cb(); } 

  log(`acePostWriteDomLineHTML NodeID#${nodeId}: Found encoded tbljson string: ${encodedJsonString}`); // Log encoded string
  try { 
      const decoded = dec(encodedJsonString); 
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: Decoded string: ${decoded}`); // Log decoded string
      if (decoded) rowData = JSON.parse(decoded); 
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: Parsed rowData:`, rowData); // Log parsed object
  } catch(e) { 
      console.error(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Failed to decode/parse tbljson.`, e); 
      return cb(); 
  }
  if (!rowData || !Array.isArray(rowData.cells) || typeof rowData.tblId === 'undefined' || typeof rowData.row === 'undefined') { 
      log(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Invalid or incomplete rowData (missing tblId or row).`); 
      return cb(); 
  }

  // *** Use tblId:row as the cache key ***
  const cacheKey = `${rowData.tblId}:${rowData.row}`;
  log(`acePostWriteDomLineHTML NodeID#${nodeId}: Using cacheKey: ${cacheKey}`);

  // *** Rendering / Update Logic - Cache First using cacheKey ***
  log(`acePostWriteDomLineHTML NodeID#${nodeId}: --- Cache Check START ---`);
  log(`acePostWriteDomLineHTML NodeID#${nodeId}: Checking cache for key: "${cacheKey}`);
  log(`acePostWriteDomLineHTML NodeID#${nodeId}: Current cache keys:`, Object.keys(tableRowNodes));
  let targetRowNode = tableRowNodes[cacheKey];
  log(`acePostWriteDomLineHTML NodeID#${nodeId}: Result of direct cache lookup (tableRowNodes[cacheKey]):`, targetRowNode);

  // Validate cached node
  let isCachedNodeValid = false;
  if (targetRowNode) {
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: Node found in cache object. Validating...`);
      if (!isNodeInDocument(targetRowNode)) {
          log(`acePostWriteDomLineHTML NodeID#${nodeId}: Cached node invalid (not connected). Clearing cache for key ${cacheKey}.`);
          delete tableRowNodes[cacheKey];
          targetRowNode = null; // Ensure it's null after deletion
      } else {
          log(`acePostWriteDomLineHTML NodeID#${nodeId}: Cached node IS connected to document.`);
          isCachedNodeValid = true;
      }
  } else {
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: No node found in cache object for key "${cacheKey}".`);
  }
  log(`acePostWriteDomLineHTML NodeID#${nodeId}: --- Cache Check END --- Valid Cached Node Found: ${isCachedNodeValid}`);

  // Decision Point: Use Cache or Render
  if (isCachedNodeValid) { // Use the validated flag
    // --- Cache Hit: Update via CACHED node --- 
    log(`acePostWriteDomLineHTML NodeID#${nodeId}: Cache HIT for key ${cacheKey}. Updating cells in cached node.`);
    const domCells = targetRowNode.querySelectorAll('td'); 
    if (domCells.length === rowData.cells.length) {
      domCells.forEach((td, index) => {
        const cellData = rowData.cells[index]; const newText = cellData.txt || ''; const innerSpan = td.querySelector('.table-cell-content');
        if (innerSpan) { 
            if (innerSpan.textContent !== newText) { 
                log(`acePostWriteDomLineHTML NodeID#${nodeId}: Updating cached TD #${index} span from "${innerSpan.textContent}" to "${newText}"`); 
                innerSpan.textContent = newText; 
            } 
        } else { 
            console.error(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Could not find inner span in cached TD #${index}.`); 
        }
      });
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: Finished updating cells in CACHED node for key ${cacheKey}.`);

      // *** ADDED: Force the passed node's content to match the updated cached table ***
      const cachedTable = targetRowNode.closest('table.dataTable');
      if (cachedTable) {
          if (node.innerHTML !== cachedTable.outerHTML) { // Avoid unnecessary updates
               log(`acePostWriteDomLineHTML NodeID#${nodeId}: Forcing passed node innerHTML to match updated cached table.`);
               node.innerHTML = cachedTable.outerHTML;
          } else {
               log(`acePostWriteDomLineHTML NodeID#${nodeId}: Passed node innerHTML already matches cached table.`);
          }
      } else {
           console.error(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Could not find parent table for cached row node!`);
           delete tableRowNodes[cacheKey]; // Invalidate cache if structure is broken
      }
    } else {
        console.error(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Mismatch cell count in cached node (${domCells.length} vs ${rowData.cells.length}). Invalidating cache for key ${cacheKey}.`);
        delete tableRowNodes[cacheKey]; // Invalidate cache on mismatch
        targetRowNode = null; // Force execution of the block below
    }
  } 

  // --- Cache Miss or Cache Invalidated: Perform INITIAL RENDER into passed node --- 
  // If targetRowNode is null here, it means cache miss OR cache was invalidated above.
  // We trust the rowData parsed from the attribute and render the table.
  if (!targetRowNode) { 
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: Cache MISS or invalid. Performing initial render via innerHTML into PASSED node (key: ${cacheKey}).`);
      log(`acePostWriteDomLineHTML NodeID#${nodeId}: Passed node outerHTML before render:`, node.outerHTML);
      try {
        const tableHtml = buildHtmlTableForRow(rowData);
        node.innerHTML = tableHtml;
          log(`acePostWriteDomLineHTML NodeID#${nodeId}: Replaced passed node innerHTML.`);
          // Cache the newly rendered row node using cacheKey
          const newRow = node.querySelector('table.dataTable tbody tr');
          if (newRow) {
              log(`acePostWriteDomLineHTML NodeID#${nodeId}: Caching newly rendered row node with key ${cacheKey}.`);
              tableRowNodes[cacheKey] = newRow;
              log(`acePostWriteDomLineHTML NodeID#${nodeId}: Confirmed cache set for ${cacheKey}. Current keys:`, Object.keys(tableRowNodes)); // Log after setting
    } else {
              console.error(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Could not find TR element to cache after initial render.`);
          }
      } catch (renderError) {
          console.error(`[ep_tables5] acePostWriteDomLineHTML NodeID#${nodeId}: Error during initial render NodeID#${nodeId}.`, renderError);
      }
  }

  // Log caret position at the end of this hook
  log(`[CaretTrace] acePostWriteDomLineHTML: END - Node ID: ${args?.node?.id} - Current rep.selStart:`, args?.rep?.selStart);

  log(`acePostWriteDomLineHTML NodeID#${nodeId}: === END ===`);
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
  const evt = ctx.evt;
  const rep = ctx.rep;
  const editorInfo = ctx.editorInfo;

  const startLogTime = Date.now();
  // LOG: Start of aceKeyEvent
  log(`[CaretTrace] aceKeyEvent: START Key='${evt?.key}' Code=${evt?.keyCode} Type=${evt?.type} Modifiers={ctrl:${evt?.ctrlKey},alt:${evt?.altKey},meta:${evt?.metaKey},shift:${evt?.shiftKey}}`, { selStart: rep?.selStart, selEnd: rep?.selEnd });

  if (!rep || !rep.selStart || !editorInfo || !evt) {
    log('[CaretTrace] aceKeyEvent: Skipping - Missing critical context (rep, selStart, editorInfo, or evt).');
    return false;
  }

  const currentLineNum = rep.selStart[0];
  // LOG: Current line number
  log(`[CaretTrace] aceKeyEvent: Caret is on line number: ${currentLineNum}`);
  const lineNode = rep.lines.atIndex(currentLineNum)?.lineNode;
  // LOG: Line node
  // log('[CaretTrace] aceKeyEvent: Corresponding line node:', lineNode); // Potentially verbose

  // --- Determine if inside table cell using lastClickedCellInfo stored on ace editor ---
  let isInsideTableCell = false;
  let targetCellIndex = -1;
  const editor = editorInfo.editor; // Get reference to the ace editor object
  let lastClickedCellInfo = editor?.ep_tables5_last_clicked; // Read shared state
  // LOG: Reading shared state
  log('[CaretTrace] aceKeyEvent: Reading shared click state (editor.ep_tables5_last_clicked):', lastClickedCellInfo);

  if (lastClickedCellInfo && lastClickedCellInfo.lineNum === currentLineNum) {
      // Optional: Verify tblId matches current line attribute? Seems complex, skipping for now.
      // const currentLineAttrib = ctx.documentAttributeManager?.getAttributeOnLine(currentLineNum, ATTR_TABLE_JSON);
      // let currentTblId = null;
      // try { currentTblId = currentLineAttrib ? JSON.parse(currentLineAttrib).tblId : null; } catch(e){}
      // if (currentTblId === lastClickedCellInfo.tblId) { ... }

      isInsideTableCell = true;
      targetCellIndex = lastClickedCellInfo.cellIndex;
      // LOG: Using shared state
      log(`[CaretTrace] aceKeyEvent: SUCCESS - Using last clicked cell info: Line=${currentLineNum}, Index=${targetCellIndex}, TblId=${lastClickedCellInfo.tblId}`);
  } else {
      // LOG: Mismatch or no shared state
      log(`[CaretTrace] aceKeyEvent: INFO - Caret on line ${currentLineNum}, but no matching clicked cell info found or line mismatch.`, {lastClickedCellInfo});
      // Optional: Clear selection visual if caret moved away and info exists
      // This clearing is now primarily handled by the mousedown listener in initialisation.js
      // However, we should ensure the shared state is cleared if we detect a mismatch here.
      if (lastClickedCellInfo) {
           log('[CaretTrace] aceKeyEvent: Clearing potentially stale cell info because caret moved to a different line.');
           if(editor) editor.ep_tables5_last_clicked = null; // Clear shared state
           // Visual clearing is handled by initialisation.js mousedown handler
           lastClickedCellInfo = null; // Clear local copy too
      }
  }
  // LOG: Final decision on cell status
  log(`[CaretTrace] aceKeyEvent: Determination -> isInsideTableCell = ${isInsideTableCell}, targetCellIndex = ${targetCellIndex}`);

  // --- Calculate Relative Caret Position --- 
  // Needed for applying edits correctly within the cell's text.
  let relativeCaretPos = -1;
  if (isInsideTableCell) {
      log('[CaretTrace] aceKeyEvent: Calculating relative caret position...');
      let originalTblJsonString = null;
      try {
          originalTblJsonString = ctx.documentAttributeManager?.getAttributeOnLine(currentLineNum, ATTR_TABLE_JSON);
          // LOG: Original attribute string for calculation
          log('[CaretTrace] aceKeyEvent: Fetched original tbljson for relative caret calc:', originalTblJsonString);
          if (originalTblJsonString) {
              const rowData = JSON.parse(originalTblJsonString);
              // LOG: Parsed rowData for calculation
              log('[CaretTrace] aceKeyEvent: Parsed rowData for relative caret calc:', JSON.parse(JSON.stringify(rowData))); // Deep log
              if (rowData && Array.isArray(rowData.cells) && targetCellIndex >= 0 && targetCellIndex < rowData.cells.length) {
                  // Calculate the character offset caused by preceding cells in the logical model
                  let precedingCellsOffset = 0;
                  for (let i = 0; i < targetCellIndex; i++) {
                      const cellTextLength = rowData.cells[i]?.txt?.length || 0;
                      precedingCellsOffset += cellTextLength;
                      // LOG: Preceding cell offset calculation step
                      log(`[CaretTrace] aceKeyEvent: Relative caret calc - Cell ${i} length: ${cellTextLength}, New precedingCellsOffset: ${precedingCellsOffset}`);
                      // NOTE: Assuming NO delimiters are counted in rep.selStart[1] now
                  }
                  
                  // Etherpad's column count (rep.selStart[1]) based on raw line text
                  const rawCaretCol = rep.selStart[1];
                  // The relative position is the raw column minus the offset from preceding cells
                  relativeCaretPos = rawCaretCol - precedingCellsOffset; // NEW: No initial ZWSP offset

                  // Basic bounds check within the current cell's text length
                  const currentCellTextLength = rowData.cells[targetCellIndex]?.txt?.length || 0;
                  // LOG: Bounds check values
                  log(`[CaretTrace] aceKeyEvent: Relative caret calc - Raw RelativePos: ${relativeCaretPos}, CellTextLength: ${currentCellTextLength}`);
                  relativeCaretPos = Math.max(0, Math.min(relativeCaretPos, currentCellTextLength));

                  // LOG: Final relative caret position
                  log(`[CaretTrace] aceKeyEvent: SUCCESS - Calculated relativeCaretPos = ${relativeCaretPos} (RawCol=${rawCaretCol}, PrecedingOffset=${precedingCellsOffset})`);
              } else {
                  log('[CaretTrace] aceKeyEvent: WARNING - Invalid rowData or targetCellIndex for relativeCaretPos calc.');
              }
          } else {
               log('[CaretTrace] aceKeyEvent: WARNING - Could not get original tbljson string for relativeCaretPos calc.');
          }
      } catch (e) {
           console.error('[CaretTrace] aceKeyEvent: ERROR calculating relativeCaretPos:', e, {originalTblJsonString});
           log('[CaretTrace] aceKeyEvent: Error details:', { error: e.message, stack: e.stack });
      }
       if (relativeCaretPos === -1) {
            log('[CaretTrace] aceKeyEvent: FAILED - Failed to calculate relativeCaretPos. Aborting key handling.');
            isInsideTableCell = false; // Prevent further processing if pos is unknown
       }
  }
  // --- End Relative Caret Position --- 

  // 1. Allow navigation keys immediately
    if ([33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode)) {
      // LOG: Allowing navigation key
      log('[CaretTrace] aceKeyEvent: Allowing navigation key:', evt.key);
    return false; // Allow default browser/ACE handling
  }

  // 2. Handle Tab within table cell
  if (isInsideTableCell && evt.key === 'Tab') {
     // LOG: Handling Tab
     log('[CaretTrace] aceKeyEvent: Tab key pressed in cell - Preventing default for now. Needs implementation.');
     evt.preventDefault(); 
     // TODO: Implement cell navigation logic
     return true; // We handled it (by preventing default)
  }

  // 3. Define keys we want to handle with attribute updates
  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete';
  const isBackspaceKey = evt.key === 'Backspace';
  const isEnterKey = evt.key === 'Enter'; // Let's handle Enter preventively too
  const shouldHandle = isTypingKey || isDeleteKey || isBackspaceKey || isEnterKey;
  // LOG: Key classification
  log(`[CaretTrace] aceKeyEvent: Key classification: isTypingKey=${isTypingKey}, isDeleteKey=${isDeleteKey}, isBackspaceKey=${isBackspaceKey}, isEnterKey=${isEnterKey}, shouldHandle=${shouldHandle}`);

  // 4. If inside cell AND it's a key we handle, proceed
  if (isInsideTableCell && targetCellIndex !== -1 && shouldHandle) {
    // LOG: Entering handling block
    log(`[CaretTrace] aceKeyEvent: HANDLED KEY - Entering handling block for key='${evt.key}' Type='${evt.type}' CellIndex=${targetCellIndex}`);

    // *** ADDED: Process only keydown for manual handling ***
    if (evt.type !== 'keydown') {
        // LOG: Ignoring non-keydown
        log(`[CaretTrace] aceKeyEvent: Ignoring non-keydown event type ('${evt.type}') for handled key.`);
        // Return false allows other handlers (like keyup) to proceed if needed,
        // but we primarily care about preventing default on keydown.
        return false; 
    }

    // Prevent default browser action for these keys (only on keydown)
    // LOG: Preventing default
    log('[CaretTrace] aceKeyEvent: Preventing default browser action for keydown event.');
    evt.preventDefault();

    let textModified = false;
    let newRelativeCaretPos = relativeCaretPos; // Will be updated if text modified
    let newRowData = null; 

    // Try to get current attribute data again (needed for modification)
    let originalRowData = null;
    try {
        const originalTblJsonString = ctx.documentAttributeManager?.getAttributeOnLine(currentLineNum, ATTR_TABLE_JSON);
        // LOG: Fetching original attribute for modification
        log('[CaretTrace] aceKeyEvent: Fetching original tbljson for modification:', originalTblJsonString);
        if (originalTblJsonString) originalRowData = JSON.parse(originalTblJsonString);
        else { throw new Error('Original tbljson attribute not found for modification'); }
    } catch(e) {
        console.error('[CaretTrace] aceKeyEvent: ERROR getting/parsing original attribute for modification:', e);
        log('[CaretTrace] aceKeyEvent: Error details:', { error: e.message, stack: e.stack });
        return true; // Abort but return true as we prevented default
    }
    
    // Calculate the new rowData based on the key event
    // LOG: Preparing for calculation
    log('[CaretTrace] aceKeyEvent: Calculating newRowData based on key event...', { originalRowData: JSON.parse(JSON.stringify(originalRowData)) }); // Deep log
    const currentCellText = originalRowData.cells[targetCellIndex]?.txt || '';
    // LOG: Current cell text and relative position used
    log(`[CaretTrace] aceKeyEvent: -> currentCellText (from originalRowData[${targetCellIndex}]): "${currentCellText}"`);
    log(`[CaretTrace] aceKeyEvent: -> relativeCaretPos used for slice/modification: ${relativeCaretPos}`);

    newRowData = JSON.parse(JSON.stringify(originalRowData)); // Deep clone

    // --- Apply modification based on key ---
    if (isTypingKey) {
        log(`[CaretTrace] aceKeyEvent: -> Applying typing modification: key='${evt.key}'`);
        const newText = currentCellText.slice(0, relativeCaretPos) + evt.key + currentCellText.slice(relativeCaretPos);
        log(`[CaretTrace] aceKeyEvent: -> Calculated newText: "${newText}"`);
        newRowData.cells[targetCellIndex].txt = newText;
        newRelativeCaretPos++;
        textModified = true;
    } else if (isBackspaceKey) {
        log(`[CaretTrace] aceKeyEvent: -> Applying backspace modification.`);
        if (relativeCaretPos > 0) {
            const newText = currentCellText.slice(0, relativeCaretPos - 1) + currentCellText.slice(relativeCaretPos);
            log(`[CaretTrace] aceKeyEvent: -> Calculated newText: "${newText}"`);
            newRowData.cells[targetCellIndex].txt = newText;
            newRelativeCaretPos--;
            textModified = true;
        } else {
             log(`[CaretTrace] aceKeyEvent: -> Backspace at start of cell (relativeCaretPos=${relativeCaretPos}), no text change.`);
        }
    } else if (isDeleteKey) {
        log(`[CaretTrace] aceKeyEvent: -> Applying delete modification.`);
        if (relativeCaretPos < currentCellText.length) {
            const newText = currentCellText.slice(0, relativeCaretPos) + currentCellText.slice(relativeCaretPos + 1);
            log(`[CaretTrace] aceKeyEvent: -> Calculated newText: "${newText}"`);
            newRowData.cells[targetCellIndex].txt = newText;
            textModified = true;
        } else {
             log(`[CaretTrace] aceKeyEvent: -> Delete at end of cell (relativeCaretPos=${relativeCaretPos}, length=${currentCellText.length}), no text change.`);
        }
    } else if (isEnterKey) {
        // LOG: Enter key handling (currently ignored)
        log('[CaretTrace] aceKeyEvent: -> Enter key pressed, currently ignored. Preventing default only.');
        // textModified remains false, so attribute won't be updated.
    } 
    // --- End modification ---

    // LOG: Result of modification attempt
    log(`[CaretTrace] aceKeyEvent: Modification check -> textModified = ${textModified}`);
    if (textModified) {
        log('[CaretTrace] aceKeyEvent: Text modified. Preparing to update attribute.');
        try {
            const newTblJsonAttribString = JSON.stringify(newRowData);
            const newCanonicalLineText = newRowData.cells.map(cell => cell.txt).join(''); // Calculate new canonical text
    
            // ** NEW: Attempt to update line text directly BEFORE applying attribute **
            const oldCanonicalLineText = ctx.rep.lines.atIndex(currentLineNum).text;
            if (newCanonicalLineText !== oldCanonicalLineText) {
                log(`[CaretTrace] aceKeyEvent: -> Manually updating canonical line text from "${oldCanonicalLineText}" to "${newCanonicalLineText}"`);
                // Use replaceRange to replace the entire old text with the new canonical text
                editorInfo.ace_performDocumentReplaceRange(
                    [currentLineNum, 0],                          // Start of line
                    [currentLineNum, oldCanonicalLineText.length], // End of OLD canonical text
                    newCanonicalLineText                           // Replace with NEW canonical text
                );
                 // Sync this text change immediately
                 editorInfo.ace_fastIncorp(10); // Increased hint value slightly
            }
            // ** END NEW BLOCK **
    
            let attributeAppliedSuccessfully = false; // Flag reset before attribute attempt
    
            // Now apply the attribute, hopefully Etherpad's internal state is more consistent
            log('[CaretTrace] aceKeyEvent: Applying tbljson attribute update:', newTblJsonAttribString);
            log('[CaretTrace] aceKeyEvent: -> Calling ace_performDocumentApplyAttributesToRange...');
            const lineLength = ctx.rep.lines.atIndex(currentLineNum).text.length; // Use potentially updated length
            log('[CaretTrace] aceKeyEvent: -> Current canonical line length:', lineLength);
            const attribRangeStart = [currentLineNum, 0];
            // ** MODIFIED: Apply attribute over potentially updated line length **
            const attribRangeEnd = [currentLineNum, Math.max(1, lineLength)]; // Cover new text, min 1
            log('[CaretTrace] aceKeyEvent: -> Applying attribute to range:', {start: attribRangeStart, end: attribRangeEnd});
            log('[CaretTrace] aceKeyEvent: -> Canonical line text before attrib apply:', ctx.rep.lines.atIndex(currentLineNum).text); // Log the now hopefully updated text
    
            // Inner try specifically for the attribute application
            try {
                editorInfo.ace_performDocumentApplyAttributesToRange(
                    attribRangeStart,
                    attribRangeEnd,
                    [[ATTR_TABLE_JSON, newTblJsonAttribString]]
                );
                attributeAppliedSuccessfully = true; // Set flag ONLY if no error thrown
                log('[CaretTrace] aceKeyEvent: -> ace_performDocumentApplyAttributesToRange call finished successfully.');
            } catch (attributeError) {
                // Log the specific error from attribute application
                log('[CaretTrace] aceKeyEvent: ERROR during ace_performDocumentApplyAttributesToRange:', attributeError);
                console.error('[ep_tables5] Error applying attribute:', attributeError);
                log('[CaretTrace] aceKeyEvent: Error details:', { error: attributeError.message, stack: attributeError.stack });
                // Ensure subsequent caret positioning is skipped
                attributeAppliedSuccessfully = false;
            }
    
    
            // Only call fastIncorp and positioning if attribute seemed successful
            if (attributeAppliedSuccessfully) {
                 log('[CaretTrace] aceKeyEvent: -> Calling ace_fastIncorp(5) for synchronization hint (after successful attrib)...');
                 editorInfo.ace_fastIncorp(5);
                 log('[CaretTrace] aceKeyEvent: -> ace_fastIncorp(5) call finished.');
    
    
                // *** Manual Caret Positioning ***
                log('[CaretTrace] aceKeyEvent: Starting manual caret positioning (Attribute applied successfully)...');
                // Calculation should still be based on newRowData's cell lengths, as this reflects the *logical* desired state.
                let absoluteCaretPos = 0;
                let calculated = false;
                try {
                    log(`[CaretTrace] aceKeyEvent: -> Calculating absolute caret pos based on newRowData (targetCellIndex=${targetCellIndex}):`, JSON.parse(JSON.stringify(newRowData)));
                    if (newRowData && newRowData.cells && targetCellIndex >= 0) {
                         for(let i = 0; i < targetCellIndex; i++) {
                             const cellTextLength = newRowData.cells[i]?.txt?.length || 0;
                             absoluteCaretPos += cellTextLength;
                             log(`[CaretTrace] aceKeyEvent: -> After cell ${i} content ('${newRowData.cells[i]?.txt}'): absoluteCaretPos=${absoluteCaretPos}`);
                         }
                         absoluteCaretPos += newRelativeCaretPos;
                         log(`[CaretTrace] aceKeyEvent: -> After adding newRelativeCaretPos (${newRelativeCaretPos}): absoluteCaretPos=${absoluteCaretPos}`);
                         calculated = true;
                     } else {
                          log('[CaretTrace] aceKeyEvent: -> WARNING: Cannot calculate absolute caret pos - invalid newRowData or targetCellIndex.');
                     }
    
                     if (calculated) {
                        log(`[CaretTrace] aceKeyEvent: -> Final calculated absolute caret pos for canonical text: ${absoluteCaretPos}`);
                        const targetSelection = [currentLineNum, absoluteCaretPos];
                         log(`[CaretTrace] aceKeyEvent: -> Attempting to set selection using ace_performSelectionChange:`, targetSelection);
                         editorInfo.ace_performSelectionChange(targetSelection, targetSelection, false);
                         log(`[CaretTrace] aceKeyEvent: -> ace_performSelectionChange call finished.`);
                     } else {
                         log('[CaretTrace] aceKeyEvent: -> Skipping ace_performSelectionChange because absoluteCaretPos calculation failed.');
                     }
                 } catch (e) {
                      console.error(`[CaretTrace] aceKeyEvent: ERROR during manual caret positioning:`, e);
                      log('[CaretTrace] aceKeyEvent: Error details:', { error: e.message, stack: e.stack });
                 }
             } else {
                  log('[CaretTrace] aceKeyEvent: Skipping fastIncorp hint and manual caret positioning because attribute application failed.');
             }
    
        } catch (error) {
            // Catch errors from the outer try (e.g., from replaceRange or initial checks)
            log('[CaretTrace] aceKeyEvent: ERROR during text replacement or attribute application process:', error);
            console.error('[ep_tables5] Error processing key event update:', error);
            log('[CaretTrace] aceKeyEvent: Error details:', { error: error.message, stack: error.stack });
        }
    } else {
        // LOG: No text modification occurred
        log(`[CaretTrace] aceKeyEvent: No text modification calculated for key '${evt.key}', only prevented default.`);
    }
       
    const endLogTime = Date.now();
    // LOG: End of handled key path
    log(`[CaretTrace] aceKeyEvent: END (Handled Key) Key='${evt.key}' Type='${evt.type}' -> Returned true. Duration: ${endLogTime - startLogTime}ms`);
    return true; // Indicate we handled the keydown event (by preventing default, potentially modifying)

  } // End if(isInsideTableCell && shouldHandle)

  // Allow default for navigation keys handled above, or any other keys not intercepted
  const endLogTimeDefault = Date.now();
  // LOG: End of default path
  log(`[CaretTrace] aceKeyEvent: END (Default Allowed) Key='${evt.key}' Type='${evt.type}' -> Returned false. Duration: ${endLogTimeDefault - startLogTime}ms`);
  return false; // Allow default browser/ACE handling
};

// ───────────────────── ace init + public helpers ─────────────────────
exports.aceInitialized = (h, ctx) => {
  log('aceInitialized: START', { h, ctx });
  const ed  = ctx.editorInfo;
  const rep = ctx.rep;
  const docManager = ctx.documentAttributeManager;
  const ZWSP_DELIMITER = '\u200B|\u200B'; // Delimiter used in collectContentPre

  // helper to apply attributes after inserting table lines
  // This needs to set the tbljson attribute which triggers aceCreateDomLine/acePostWriteDomLineHTML
  function applyTableLineAttribute (lineNum, tblId, rowIndex, cellContents) {
    log(`applyTableLineAttribute: Applying attribute to line ${lineNum}`, {tblId, rowIndex, cellContents});
    const rowData = {
        tblId: tblId,
        row: rowIndex,
        cells: cellContents.map(txt => ({ txt: txt || ' ' })) // Ensure cells have txt property
    };
    const attributeString = JSON.stringify(rowData);
    log(`applyTableLineAttribute: Attribute String: ${attributeString}`);
    try {
       // REVERT TO: Use ace_performDocumentApplyAttributesToRange for potentially better hook triggering
       const lineLength = rep.lines.atIndex(lineNum).text.length; // Get current line length
       const start = [lineNum, 0];
       const end = [lineNum, lineLength]; // Apply to entire line text range
       log(`applyTableLineAttribute: Applying attribute via ace_performDocumentApplyAttributesToRange to [${start}]-[${end}]`);
       ed.ace_performDocumentApplyAttributesToRange(start, end, [[ATTR_TABLE_JSON, attributeString]]);
       
       log(`applyTableLineAttribute: Applied attribute to line ${lineNum}`);
    } catch(e) {
        console.error(`[ep_tables5] Error applying attribute on line ${lineNum}:`, e);
    }
  }

  /** Insert a fresh rows×cols blank table at the caret */
  ed.ace_createTableViaAttributes = (rows = 2, cols = 2) => {
    log('ace_createTableViaAttributes: START', { rows, cols });
    rows = Math.max(1, rows); cols = Math.max(1, cols);
    ed.ace_fastIncorp(20); // Sync changes before insertion

    const tblId   = rand();
    const initialCellText = ' '; // Start with a single space
    const initialCellContents = Array.from({ length: cols }).fill(initialCellText);
    // Revert to inserting ZWSP delimited text
    const lineTxt = initialCellContents.join(''); 
    const block = Array.from({ length: rows }).fill(lineTxt).join('\n') + '\n'; 

    const start=[...rep.selStart], end=[...rep.selEnd];
    log('ace_createTableViaAttributes: Current selection:', { start, end });

    // Replace selection with placeholder text lines
    ed.ace_performDocumentReplaceRange(start, end, block);
    log('ace_createTableViaAttributes: Inserted block of placeholder text lines.');
    ed.ace_fastIncorp(20); // Sync text insertion

    // Apply attribute to each inserted line
    for (let r = 0; r < rows; r++) {
      const lineNumToApply = start[0] + r;
      log(`ace_createTableViaAttributes: Applying attribute for row ${r} on line ${lineNumToApply}`);
      const rowData = { 
          tblId: tblId,
          row: r,
          cells: initialCellContents.map(txt => ({ txt: txt }))
      };
      const attributeString = JSON.stringify(rowData);
      ed.ace_performDocumentApplyAttributesToRange(
            [lineNumToApply, 0], 
            [lineNumToApply, 1], // Apply to line marker range
            [[ATTR_TABLE_JSON, attributeString]]
      );
    }
    
    ed.ace_fastIncorp(20); // Final sync
    
    // REMOVED: Setting caret position after table creation caused errors.
    // Etherpad will place caret based on default behavior.
    // const lastRowLine = start[0] + rows - 1;\n    // const finalCaretPos = [lastRowLine + 1, 0]; \n    // log(\'ace_createTableViaAttributes: Setting final caret position to line after table:\', finalCaretPos);\n    // try {\n    //   ed.ace_setSelection(finalCaretPos, finalCaretPos);\n    // } catch(e) {\n    //    console.error(\'[ep_tables5] Error setting caret position after table creation:\', e);\n    // }\n\n    log(\'ace_createTableViaAttributes: END\');\n  };

    log('ace_createTableViaAttributes: END');
  };

  ed.ace_doDatatableOptions = () => {
    log('ace_doDatatableOptions: CALLED (Not Implemented)');
    // TODO: Implement row/column add/delete operations
    // These will involve manipulating the `tbljson` attribute on affected lines
    // and potentially inserting/deleting lines using ace_performDocumentReplaceRange
  }
  log('aceInitialized: END - helpers defined.');
};

// ───────────────────── required no‑op stubs ─────────────────────
exports.aceStartLineAndCharForPoint = () => { return undefined; };
exports.aceEndLineAndCharForPoint   = () => { return undefined; };
exports.aceSetAuthorStyle           = () => {};
// Return the relative path to the CSS file needed within the editor iframe
exports.aceEditorCSS                = () => { 
  // Path relative to Etherpad's static/plugins/ directory
  // Format should be: pluginName/path/to/file.css
  return ['ep_tables5/static/css/datatables-editor.css'];
};

// Register TABLE as a block element, hoping it influences rendering behavior
exports.aceRegisterBlockElements = () => ['table'];

// END OF FILE
