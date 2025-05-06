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
const DELIMITER       = '|'; // SIMPLIFIED DELIMITER
const HIDDEN_DELIM    = '|';        // Keep the real char for DOM alignment

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

// NEW: Module-level state for last clicked cell
let lastClickedCellInfo = null; // { lineNum: number, cellIndex: number, tblId: string }

// ────────────────────── collectContentPre (DOM → atext) ─────────────────────
exports.collectContentPre = (hook, ctx) => {
  const funcName = 'collectContentPre';
  const node = ctx.dom;
  const rep = ctx?.rep;
  const docAttrManager = ctx?.documentAttributeManager;

  // Log entry point
  log(`${funcName}: START - Processing node ID: ${node?.id}, Class: ${node?.className}`);

  // --- Check if it's a line DIV that WE rendered as a table --- 
  // We need to reliably identify lines managed by this plugin.
  // Checking for the presence of table.dataTable seems reasonable.
  if (!(node?.classList?.contains('ace-line'))) {
    log(`${funcName} (${node?.id}): Not an ace-line div. Allowing default.`);
    return; // Let default handlers process children like spans
  }
  const tableNode = node.querySelector('table.dataTable[data-tblId]'); // Be specific
  if (!tableNode) {
    log(`${funcName} (${node?.id}): No table.dataTable[data-tblId] found within. Allowing default.`);
    return; // Not a table line rendered by us, allow default collection
  }
  // --- Found a table line --- 
  log(`${funcName} (${node?.id}): Found rendered table.dataTable. Proceeding with custom collection.`);

  // --- 1. Retrieve Existing Metadata Attribute --- 
  let existingMetadata = null;
  const lineNum = rep?.lines?.indexOfKey(node.id); // Get line number
  log(`${funcName} (${node?.id}): Determined line number: ${lineNum}`);

  if (typeof lineNum !== 'number' || lineNum < 0 || !docAttrManager) {
    console.error(`[ep_tables5] ${funcName} (${node?.id}): Could not get valid line number (${lineNum}) or docAttrManager. Cannot preserve metadata.`);
    log(`${funcName} (${node?.id}): Aborting custom collection due to missing lineNum/docAttrManager.`);
    // Maybe return undefined to prevent default, but state is inconsistent?
    // Let's allow default for now to avoid breaking things further.
    return;
  }

         try { 
    const existingAttrString = docAttrManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
    log(`${funcName} (${node?.id}): Retrieved existing attribute string for line ${lineNum}:`, existingAttrString);
    if (existingAttrString) {
      existingMetadata = JSON.parse(existingAttrString);
      log(`${funcName} (${node?.id}): Parsed existing metadata:`, existingMetadata);
      // Basic validation of retrieved metadata
      if (!existingMetadata || typeof existingMetadata.tblId === 'undefined' || typeof existingMetadata.row === 'undefined' || typeof existingMetadata.cols !== 'number') {
        log(`${funcName} (${node?.id}): Warning - Existing metadata is invalid/incomplete. Proceeding cautiously.`);
        console.warn(`[ep_tables5] ${funcName} (${node?.id}): Invalid metadata retrieved from line ${lineNum}.`, existingMetadata);
        existingMetadata = null; // Discard invalid metadata
      }
    } else {
      log(`${funcName} (${node?.id}): No existing ${ATTR_TABLE_JSON} attribute found on line ${lineNum}.`);
      // This shouldn't happen if acePostWriteDomLineHTML ran correctly, but handle defensively.
      return; // Allow default collection if metadata is missing
    }
  } catch (e) {
    console.error(`[ep_tables5] ${funcName} (${node?.id}): Error parsing existing tbljson attribute for line ${lineNum}.`, e);
    log(`${funcName} (${node?.id}): Error details:`, { message: e.message, stack: e.stack });
    // Allow default collection on error
    return;
  }
  if (!existingMetadata) {
      // Should not happen if we passed checks above, but safety first.
      log(`${funcName} (${node?.id}): Failed to secure valid existing metadata. Allowing default collection.`);
      return;
  }
  // --- End Metadata Retrieval --- 

  // --- 2. Construct Canonical Line Text from Rendered TD innerHTML --- 
    const trNode = tableNode.querySelector('tbody > tr');
    if (!trNode) {
    log(`${funcName} (${node?.id}): ERROR - Could not find <tr> in rendered table. Cannot construct canonical text.`);
    console.error(`[ep_tables5] ${funcName} (${node?.id}): Could not find <tr> in rendered table.`);
    // Allow default collection as we can't determine the correct text
    return;
    }

  // Extract innerHTML from each TD in the rendered row
  const cellHTMLSegments = Array.from(trNode.children).map((td, index) => {
    // Assuming buildTableFromDelimitedHTML placed content directly in TD
    let segmentHTML = td.innerHTML || ''; // Get raw HTML content
    let cleanContent = segmentHTML;

    // For cells after the first, remove the hidden delimiter span wrapper before joining
    if (index > 0) {
       // Regex to match the specific span structure at the beginning
       const hiddenDelimRegex = /^<span class="ep-tables5-delim">\|<\/span>/;
       cleanContent = segmentHTML.replace(hiddenDelimRegex, '');
    }

    log(`${funcName} (${node?.id}): Reading rendered TD #${index} innerHTML:`, segmentHTML);
    log(`${funcName} (${node?.id}): Cleaned segment content for cell ${index}:`, cleanContent);
    return cleanContent;
  });
  log(`${funcName} (${node?.id}): Extracted HTML segments from TDs:`, cellHTMLSegments);

  // Join segments with the delimiter to form the canonical line text
  const canonicalLineText = cellHTMLSegments.join(DELIMITER);
  ctx.state.line = canonicalLineText;
  log(`${funcName} (${node?.id}): Set ctx.state.line to delimited HTML: "${canonicalLineText}"`);
  // --- End Canonical Text Construction --- 

  // --- 3. Preserve Existing Metadata Attribute --- 
  // Push the *retrieved* metadata string back onto the attributes for this line.
  // This ensures it persists without modification by this hook.
  const attributeString = JSON.stringify(existingMetadata);
    ctx.lineAttributes.push([ATTR_TABLE_JSON, attributeString]);
  log(`${funcName} (${node?.id}): Pushed existing metadata attribute back onto lineAttributes.`);
  // --- End Attribute Preservation --- 

    // Prevent default processing for the line div's children (the table)
  // This stops Etherpad from trying to re-collect text from the rendered table.
  log(`${funcName} (${node?.id}): END - Preventing default collection for line children.`);
    return undefined; 
};

// ───────────── attribute → span‑class mapping (linestylefilter hook) ─────────
exports.aceAttribsToClasses = (hook, ctx) => {
  const funcName = 'aceAttribsToClasses';
  log(`>>>> ${funcName}: Called with key: ${ctx.key}`); // Log entry
  if (ctx.key === ATTR_TABLE_JSON) {
    log(`${funcName}: Processing ATTR_TABLE_JSON.`);
    // ctx.value is the raw JSON string from Etherpad's attribute pool
    const rawJsonValue = ctx.value;
    log(`${funcName}: Received raw attribute value (ctx.value):`, rawJsonValue);

    // Attempt to parse for logging purposes
    let parsedMetadataForLog = '[JSON Parse Error]';
    try {
        parsedMetadataForLog = JSON.parse(rawJsonValue);
        log(`${funcName}: Value parsed for logging:`, parsedMetadataForLog);
    } catch(e) {
        log(`${funcName}: Error parsing raw JSON value for logging:`, e);
        // Continue anyway, enc() might still work if it's just a string
    }

    // Generate the class name by base64 encoding the raw JSON string.
    // This ensures acePostWriteDomLineHTML receives the expected encoded format.
    const className = `tbljson-${enc(rawJsonValue)}`;
    log(`${funcName}: Generated class name by encoding raw JSON: ${className}`);
    return [className];
  }
  if (ctx.key === ATTR_CELL) {
    // Keep this in case we want cell-specific styling later
    // log(`${funcName}: Processing ATTR_CELL: ${ctx.value}`); // Optional: Uncomment if needed
    return [`tblCell-${ctx.value}`];
  }
  // log(`${funcName}: Processing other key: ${ctx.key}`); // Optional: Uncomment if needed
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

// NEW Helper function to build table HTML from pre-rendered delimited content
function buildTableFromDelimitedHTML(metadata, innerHTMLSegments) {
  const funcName = 'buildTableFromDelimitedHTML';
  log(`${funcName}: START`, { metadata, innerHTMLSegments });

  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined') {
    console.error(`[ep_tables5] ${funcName}: Invalid or missing metadata. Aborting.`);
    log(`${funcName}: END - Error`);
    return '<table class="dataTable dataTable-error"><tbody><tr><td>Error: Missing table metadata</td></tr></tbody></table>'; // Return error table
  }

  // Basic styling - can be moved to CSS later
  const tdStyle = `padding: 5px 7px; word-wrap:break-word; vertical-align: top; border: 1px solid #000;`; // Changed border style

  // Map the HTML segments directly into TD elements
  // We trust that innerHTMLSegments contains valid, pre-rendered HTML snippets
  const cellsHtml = innerHTMLSegments.map((segment, index) => {
    const hidden = index === 0 ? '' :
      /* keep the char in the DOM but make it visually disappear */
      `<span class="ep-tables5-delim">${HIDDEN_DELIM}</span>`;
    const cellContent = segment || '&nbsp;'; // Use non-breaking space for empty segments
    log(`${funcName}: Processing segment ${index}. Content:`, segment); // Log segment content
    // Wrap the raw segment HTML in a span for consistency? Or directly in TD? Let's try TD directly first.
    // Maybe add back table-cell-content span later if needed for styling/selection.
    const tdContent = `<td style="${tdStyle}">${hidden}${cellContent}</td>`;
    log(`${funcName}: Generated TD HTML for segment ${index}:`, tdContent);
    return tdContent;
  }).join('');
  log(`${funcName}: Joined all cellsHtml:`, cellsHtml);

  // Add 'dataTable-first-row' class if it's the logical first row (row index 0)
  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';
  log(`${funcName}: First row class applied: '${firstRowClass}'`);

  // Construct the final table HTML
  // Rely on CSS for border-collapse, width etc. Add data attributes from metadata.
  const tableHtml = `<table class="dataTable${firstRowClass}" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse: collapse;"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
  log(`${funcName}: Generated final table HTML:`, tableHtml);
  log(`${funcName}: END - Success`);
  return tableHtml;
}

// ───────────────── Populate Table Cells / Render (PostWrite) ──────────────────
exports.acePostWriteDomLineHTML = function (hook_name, args, cb) {
  const funcName = 'acePostWriteDomLineHTML';
  const node = args?.node;
  const nodeId = node?.id;
  const lineNum = args?.lineNumber; // Etherpad >= 1.9 provides lineNumber
  const logPrefix = '[ep_tables5:acePostWriteDomLineHTML]'; // Consistent prefix

  // *** STARTUP LOGGING ***
  log(`${logPrefix} ----- START ----- NodeID: ${nodeId} LineNum: ${lineNum}`);
  if (!node || !nodeId) {
      log(`${logPrefix} ERROR - Received invalid node or node without ID. Aborting.`);
      console.error(`[ep_tables5] ${funcName}: Received invalid node or node without ID.`);
      return cb();
  }
  // log(`${logPrefix} NodeID#${nodeId}: Initial node outerHTML:`, node.outerHTML); // Too verbose usually
  // log(`${logPrefix} NodeID#${nodeId}: Full args object received:`, args); // Too verbose usually
  // ***********************

  let rowMetadata = null;
  let encodedJsonString = null;

  // --- Log node classes BEFORE searching --- 
  // log(`${logPrefix} NodeID#${nodeId}: Checking classes BEFORE search. Node classList:`, node.classList);
  // ... (child class logging removed for brevity) ...
  // --- End logging node classes ---

  // --- 1. Find and Parse Metadata Attribute --- 
  log(`${logPrefix} NodeID#${nodeId}: Searching for tbljson-* class...`);
  // Check the node itself first
  if (node.classList) {
      for (const cls of node.classList) { 
          if (cls.startsWith('tbljson-')) {
              encodedJsonString = cls.substring(8);
              log(`${logPrefix} NodeID#${nodeId}: Found encoded tbljson on node itself: ${encodedJsonString}`);
              break;
      }
  } 
  }
  // Check children if not found on node
  if (!encodedJsonString && node.children) {
       // log(`${logPrefix} NodeID#${nodeId}: Not found on node, checking children...`);
        for (const child of node.children) {
             if (child.classList) {
                for (const cls of child.classList) {
                 if (cls.startsWith('tbljson-')) {
                     encodedJsonString = cls.substring(8);
                     log(`${logPrefix} NodeID#${nodeId}: Found encoded tbljson on child ${child.tagName}: ${encodedJsonString}`);
                     break;
                 }
              }
          } 
          if (encodedJsonString) break;
      }
  } 

  // If no attribute found, it's not a table line managed by us
  if (!encodedJsonString) {
      log(`${logPrefix} NodeID#${nodeId}: No tbljson-* class found. Assuming not a table line. END.`);
      return cb(); 
  }

  // *** NEW CHECK: If table already rendered, skip regeneration ***
  const existingTable = node.querySelector('table.dataTable[data-tblId]');
  if (existingTable) {
      log(`${logPrefix} NodeID#${nodeId}: Table already exists in DOM. Skipping innerHTML replacement.`);
      // Optionally, verify tblId matches metadata? For now, assume it's correct.
      // const existingTblId = existingTable.getAttribute('data-tblId');
      // try {
      //    const decoded = dec(encodedJsonString); 
      //    const currentMetadata = JSON.parse(decoded);
      //    if (existingTblId === currentMetadata?.tblId) { ... } 
      // } catch(e) { /* ignore validation error */ }
      return cb(); // Do nothing further
  }
  // *** END NEW CHECK ***

  log(`${logPrefix} NodeID#${nodeId}: Decoding and parsing metadata...`);
  try { 
      const decoded = dec(encodedJsonString); 
      log(`${logPrefix} NodeID#${nodeId}: Decoded string: ${decoded}`);
      if (!decoded) throw new Error('Decoded string is null or empty.');
      rowMetadata = JSON.parse(decoded);
      log(`${logPrefix} NodeID#${nodeId}: Parsed rowMetadata:`, rowMetadata);

      // Validate essential metadata
      if (!rowMetadata || typeof rowMetadata.tblId === 'undefined' || typeof rowMetadata.row === 'undefined' || typeof rowMetadata.cols !== 'number') {
          throw new Error('Invalid or incomplete metadata (missing tblId, row, or cols).');
      }
      log(`${logPrefix} NodeID#${nodeId}: Metadata validated successfully.`);

  } catch(e) { 
      log(`${logPrefix} NodeID#${nodeId}: FATAL ERROR - Failed to decode/parse/validate tbljson metadata. Rendering cannot proceed.`, e);
      console.error(`[ep_tables5] ${funcName} NodeID#${nodeId}: Failed to decode/parse/validate tbljson.`, encodedJsonString, e);
      // Optionally render an error state in the node?
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_tables5] Error: Invalid table metadata attribute found.</div>';
      log(`${logPrefix} NodeID#${nodeId}: Rendered error message in node. END.`);
      return cb(); 
  }
  // --- End Metadata Parsing ---

  // --- 2. Get and Parse Line Content ---
  const lineInnerHTML = node.innerHTML;
  log(`${logPrefix} NodeID#${nodeId}: Parsing line content via Placeholder Replace/Split...`);
  log(`${logPrefix} NodeID#${nodeId}: Original node.innerHTML:`, lineInnerHTML);

  let htmlSegments = [];
  // Define a unique placeholder unlikely to be in content
  const PLACEHOLDER = '@@EP_TABLES5_DELIM@@'; 

  try {
      // Replace the simple pipe delimiter with the placeholder
      // Need to use a RegExp with the global flag for replaceAll behavior
      // Escape the pipe for RegExp: \|
      const replacedHtml = lineInnerHTML.replace(new RegExp('\\|', 'g'), PLACEHOLDER);
      log(`${logPrefix} NodeID#${nodeId}: innerHTML after replacing delimiter ('|') with placeholder:`, replacedHtml);

      // Split the string using the placeholder
      htmlSegments = replacedHtml.split(PLACEHOLDER);
      log(`${logPrefix} NodeID#${nodeId}: Final parsed HTML segments (${htmlSegments.length}) after splitting by placeholder:`, htmlSegments);

      // --- Validation --- 
      if (htmlSegments.length !== rowMetadata.cols) {
          log(`${logPrefix} NodeID#${nodeId}: WARNING - Parsed segment count (${htmlSegments.length}) does not match metadata cols (${rowMetadata.cols}). Table structure might be incorrect.`);
          console.warn(`[ep_tables5] ${funcName} NodeID#${nodeId}: Parsed segment count (${htmlSegments.length}) mismatch with metadata cols (${rowMetadata.cols}). Segments:`, htmlSegments);
      } else {
          log(`${logPrefix} NodeID#${nodeId}: Parsed segment count matches metadata cols (${rowMetadata.cols}).`);
      }
  } catch (parseError) {
      log(`${logPrefix} NodeID#${nodeId}: ERROR during placeholder replace/split parsing. Cannot build table.`, parseError);
      console.error(`[ep_tables5] ${funcName} NodeID#${nodeId}: Error parsing line content via placeholder replace/split.`, parseError);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_tables5] Error: Could not parse table cell content.</div>';
      log(`${logPrefix} NodeID#${nodeId}: Rendered parse error message in node. END.`);
      return cb();
  }
  // --- End Content Parsing ---

  // --- 3. Build and Render Table ---
  log(`${logPrefix} NodeID#${nodeId}: Calling buildTableFromDelimitedHTML...`);
      try {
      const newTableHTML = buildTableFromDelimitedHTML(rowMetadata, htmlSegments);
      log(`${logPrefix} NodeID#${nodeId}: Received new table HTML from helper. Replacing node.innerHTML.`);
      // Replace the node's content entirely with the generated table
      node.innerHTML = newTableHTML;
      log(`${logPrefix} NodeID#${nodeId}: Successfully replaced node.innerHTML with new table structure.`);
      } catch (renderError) {
      log(`${logPrefix} NodeID#${nodeId}: ERROR during table building or rendering.`, renderError);
      console.error(`[ep_tables5] ${funcName} NodeID#${nodeId}: Error building/rendering table.`, renderError);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_tables5] Error: Failed to render table structure.</div>';
      log(`${logPrefix} NodeID#${nodeId}: Rendered build/render error message in node. END.`);
      return cb();
  }
  // --- End Table Building ---

  // *** REMOVED CACHING LOGIC ***
  // The old logic based on tableRowNodes cache is completely removed.

  log(`${logPrefix}: ----- END ----- NodeID: ${nodeId}`);
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
  const logPrefix = '[ep_tables5:aceKeyEvent]';
  log(`${logPrefix} START Key='${evt?.key}' Code=${evt?.keyCode} Type=${evt?.type} Modifiers={ctrl:${evt?.ctrlKey},alt:${evt?.altKey},meta:${evt?.metaKey},shift:${evt?.shiftKey}}`, { selStart: rep?.selStart, selEnd: rep?.selEnd });

  if (!rep || !rep.selStart || !editorInfo || !evt || !docManager) {
    log(`${logPrefix} Skipping - Missing critical context.`);
    return false;
  }

  // Get caret info from event context - may be stale
  const reportedLineNum = rep.selStart[0];
  const reportedCol = rep.selStart[1]; 
  log(`${logPrefix} Reported caret from rep: Line=${reportedLineNum}, Col=${reportedCol}`);

  // --- Get Table Metadata for the reported line --- 
  let tableMetadata = null;
  let lineAttrString = null; // Store for potential use later
  try {
    lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
    if (lineAttrString) {
        tableMetadata = JSON.parse(lineAttrString);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
             log(`${logPrefix} Line ${reportedLineNum} has attribute, but metadata invalid/missing cols.`);
             tableMetadata = null; // Ensure it's null if invalid
        }
    } else {
        // Not a table line based on reported caret line
    }
  } catch(e) {
    console.error(`${logPrefix} Error checking/parsing line attribute for line ${reportedLineNum}.`, e);
    tableMetadata = null; // Ensure it's null on error
  }

  // Get last known good state
  const editor = editorInfo.editor; // Get editor instance
  const lastClick = editor?.ep_tables5_last_clicked; // Read shared state
  log(`${logPrefix} Reading stored click/caret info:`, lastClick);

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
      log(`${logPrefix} Attempting to validate stored click info for Line=${lastClick.lineNum}...`);
      let storedLineAttrString = null;
      let storedLineMetadata = null;
      try {
          storedLineAttrString = docManager.getAttributeOnLine(lastClick.lineNum, ATTR_TABLE_JSON);
          if (storedLineAttrString) storedLineMetadata = JSON.parse(storedLineAttrString);
          
          // Check if metadata is valid and tblId matches
          if (storedLineMetadata && typeof storedLineMetadata.cols === 'number' && storedLineMetadata.tblId === lastClick.tblId) {
              log(`${logPrefix} Stored click info VALIDATED (Metadata OK and tblId matches). Trusting stored state.`);
              trustedLastClick = true;
              currentLineNum = lastClick.lineNum; 
              targetCellIndex = lastClick.cellIndex;
              metadataForTargetLine = storedLineMetadata; 
              lineAttrString = storedLineAttrString; // Use the validated attr string
              
              lineText = rep.lines.atIndex(currentLineNum)?.text || '';
              cellTexts = lineText.split(DELIMITER);
              log(`${logPrefix} Using Line=${currentLineNum}, CellIndex=${targetCellIndex}. Text: "${lineText}"`);

              if (cellTexts.length !== metadataForTargetLine.cols) {
                  log(`${logPrefix} WARNING: Stored cell count mismatch for trusted line ${currentLineNum}.`);
              }

              cellStartCol = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                  cellStartCol += (cellTexts[i]?.length ?? 0) + DELIMITER.length;
              }
              precedingCellsOffset = cellStartCol;
              log(`${logPrefix} Calculated cellStartCol=${cellStartCol} from trusted cellIndex=${targetCellIndex}.`);

              if (typeof lastClick.relativePos === 'number' && lastClick.relativePos >= 0) {
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(lastClick.relativePos, currentCellTextLength));
                  log(`${logPrefix} Using and validated stored relative position: ${relativeCaretPos}.`);
  } else {
                  relativeCaretPos = reportedCol - cellStartCol; // Use reportedCol for initial calc if relative is missing
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(relativeCaretPos, currentCellTextLength)); 
                  log(`${logPrefix} Stored relativePos missing, calculated from reportedCol (${reportedCol}): ${relativeCaretPos}`);
              }
          } else {
              log(`${logPrefix} Stored click info INVALID (Metadata missing/invalid or tblId mismatch). Clearing stored state.`);
              if (editor) editor.ep_tables5_last_clicked = null;
          }
      } catch (e) {
           console.error(`${logPrefix} Error validating stored click info for line ${lastClick.lineNum}.`, e);
           if (editor) editor.ep_tables5_last_clicked = null; // Clear on error
      }
  }
  
  // ** Scenario 2: Fallback - Use reported line/col ONLY if stored info wasn't trusted **
  if (!trustedLastClick) {
      log(`${logPrefix} Fallback: Using reported caret position Line=${reportedLineNum}, Col=${reportedCol}.`);
      // Fetch metadata for the reported line again, in case it wasn't fetched or was invalid earlier
      try {
          lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
          if (lineAttrString) tableMetadata = JSON.parse(lineAttrString);
          if (!tableMetadata || typeof tableMetadata.cols !== 'number') tableMetadata = null;
      } catch(e) { tableMetadata = null; } // Ignore errors here, handled below

      if (!tableMetadata) {
          log(`${logPrefix} Fallback: Reported line ${reportedLineNum} is not a valid table line. Allowing default.`);
           return false;
      }
      
      currentLineNum = reportedLineNum;
      metadataForTargetLine = tableMetadata;
      log(`${logPrefix} Fallback: Processing based on reported line ${currentLineNum}.`);
      
      lineText = rep.lines.atIndex(currentLineNum)?.text || '';
      cellTexts = lineText.split(DELIMITER);
      log(`${logPrefix} Fallback: Fetched text for reported line ${currentLineNum}: "${lineText}"`);

      if (cellTexts.length !== metadataForTargetLine.cols) {
          log(`${logPrefix} WARNING (Fallback): Cell count mismatch for reported line ${currentLineNum}.`);
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
              log(`${logPrefix} --> (Fallback Calc) Found target cell ${foundIndex}. RelativePos: ${relativeCaretPos}.`);
              break; 
          }
          if (i < cellTexts.length - 1 && reportedCol === cellEndCol + DELIMITER.length) {
              foundIndex = i + 1;
              relativeCaretPos = 0; 
              cellStartCol = currentOffset + cellLength + DELIMITER.length;
              precedingCellsOffset = cellStartCol;
              log(`${logPrefix} --> (Fallback Calc) Caret at delimiter AFTER cell ${i}. Treating as start of cell ${foundIndex}.`);
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
                log(`${logPrefix} --> (Fallback Calc) Caret detected at END of last cell (${foundIndex}).`);
          } else {
            log(`${logPrefix} (Fallback Calc) FAILED to determine target cell for caret col ${reportedCol}. Allowing default handling.`);
            return false; 
          }
      }
      targetCellIndex = foundIndex;
  }

  // --- Final Validation --- 
  if (currentLineNum < 0 || targetCellIndex < 0 || !metadataForTargetLine || targetCellIndex >= metadataForTargetLine.cols) {
       log(`${logPrefix} FAILED final validation: Line=${currentLineNum}, Cell=${targetCellIndex}, Metadata=${!!metadataForTargetLine}. Allowing default.`);
       if (editor) editor.ep_tables5_last_clicked = null; 
       return false; 
      }

  log(`${logPrefix} --> Final Target: Line=${currentLineNum}, CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}`);
  // --- End Cell/Position Determination ---

  // --- Define Key Types ---
  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete' || evt.keyCode === 46;
  const isBackspaceKey = evt.key === 'Backspace' || evt.keyCode === 8;
  const isNavigationKey = [33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode);
  const isTabKey = evt.key === 'Tab';
  const isEnterKey = evt.key === 'Enter';
  log(`${logPrefix} Key classification: Typing=${isTypingKey}, Backspace=${isBackspaceKey}, Delete=${isDeleteKey}, Nav=${isNavigationKey}, Tab=${isTabKey}, Enter=${isEnterKey}`);

  // --- Handle Keys --- 

  // 1. Allow non-Tab navigation keys immediately
  if (isNavigationKey && !isTabKey) {
      log(`${logPrefix} Allowing navigation key: ${evt.key}. Clearing click state.`);
      if (editor) editor.ep_tables5_last_clicked = null; // Clear state on navigation
      return false;
  }

  // 2. Handle Tab - Prevent default (implement navigation later)
  if (isTabKey) { 
     log(`${logPrefix} Tab key pressed. Preventing default.`);
     evt.preventDefault();
     // TODO: Implement cell navigation logic
     return true;
  }

  // 3. Handle Enter - Prevent default (usually splits the line)
  if (isEnterKey) {
      log(`${logPrefix} Enter key pressed. Preventing default line split.`);
      evt.preventDefault();
      // Optional TODO: Implement behavior like moving to the next row/cell?
      return true; 
  }

  // 4. Intercept destructive keys ONLY at cell boundaries to protect delimiters
      const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
  // Backspace at the very beginning of cell > 0
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex > 0) {
      log(`${logPrefix} Intercepted Backspace at start of cell ${targetCellIndex}. Preventing default.`);
          evt.preventDefault();
          return true;
      }
  // Delete at the very end of cell < last cell
  if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex < metadataForTargetLine.cols - 1) {
      log(`${logPrefix} Intercepted Delete at end of cell ${targetCellIndex}. Preventing default.`);
          evt.preventDefault();
          return true;
      }

  // 5. Handle Typing/Backspace/Delete WITHIN a cell via manual modification
  const isInternalBackspace = isBackspaceKey && relativeCaretPos > 0;
  const isInternalDelete = isDeleteKey && relativeCaretPos < currentCellTextLength;

  if (isTypingKey || isInternalBackspace || isInternalDelete) {
    // *** Use the validated currentLineNum and currentCol derived from relativeCaretPos ***
    const currentCol = cellStartCol + relativeCaretPos;
    log(`${logPrefix} Handling INTERNAL key='${evt.key}' Type='${evt.type}' at Line=${currentLineNum}, Col=${currentCol} (CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}).`);

    // Only process keydown events for modifications
    if (evt.type !== 'keydown') {
        log(`${logPrefix} Ignoring non-keydown event type ('${evt.type}') for handled key.`);
        return false; 
    }

    log(`${logPrefix} Preventing default browser action for keydown event.`);
    evt.preventDefault();

    let newAbsoluteCaretCol = -1;
    let repBeforeEdit = null; // Store rep before edits for attribute helper

    try {
        repBeforeEdit = editorInfo.ace_getRep(); // Get rep *before* making changes

    if (isTypingKey) {
            const insertPos = [currentLineNum, currentCol];
            log(`${logPrefix} -> Inserting text '${evt.key}' at [${insertPos}]`);
            editorInfo.ace_performDocumentReplaceRange(insertPos, insertPos, evt.key);
            newAbsoluteCaretCol = currentCol + 1;

        } else if (isInternalBackspace) {
            const delRangeStart = [currentLineNum, currentCol - 1];
            const delRangeEnd = [currentLineNum, currentCol];
            log(`${logPrefix} -> Deleting (Backspace) range [${delRangeStart}]-[${delRangeEnd}]`);
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol - 1;

        } else if (isInternalDelete) {
            const delRangeStart = [currentLineNum, currentCol];
            const delRangeEnd = [currentLineNum, currentCol + 1];
            log(`${logPrefix} -> Deleting (Delete) range [${delRangeStart}]-[${delRangeEnd}]`);
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol; // Caret stays at the same column for delete
        }

        // *** CRITICAL: Re-apply the line attribute after ANY modification ***
        log(`${logPrefix} -> Re-applying tbljson line attribute...`);
        const applyHelper = editorInfo.ep_tables5_applyMeta; 
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) { 
             // Pass the original lineAttrString if available AND if it belongs to the currentLineNum
             const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;
             applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply);
             log(`${logPrefix} -> tbljson line attribute re-applied (using rep before edit).`);
                } else {
             console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
             const currentRepFallback = editorInfo.ace_getRep();
             if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
                 log(`${logPrefix} -> Retrying attribute application with current rep...`);
                 applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null); // Cannot guarantee old attr string is valid here
                 log(`${logPrefix} -> tbljson line attribute re-applied (using current rep fallback).`);
            } else {
                  console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute even with fallback rep.`);
             }
        }
        
        // Set caret position immediately
        if (newAbsoluteCaretCol >= 0) {
             const newCaretPos = [currentLineNum, newAbsoluteCaretCol]; // Use the trusted currentLineNum
             log(`${logPrefix} -> Setting selection immediately to:`, newCaretPos);
             try {
                editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                log(`${logPrefix} -> Selection set immediately.`);

                // Add sync hint AFTER setting selection
                editorInfo.ace_fastIncorp(5); 
                log(`${logPrefix} -> Requested sync hint (fastIncorp 5).`);

                // Store the updated caret info for the next event
                const newRelativePos = newAbsoluteCaretCol - cellStartCol;
                editor.ep_tables5_last_clicked = {
                    lineNum: currentLineNum, 
                    tblId: metadataForTargetLine.tblId,
                    cellIndex: targetCellIndex,
                    relativePos: newRelativePos
                };
                log(`${logPrefix} -> Updated stored click/caret info:`, editor.ep_tables5_last_clicked);

            } catch (selError) {
                 console.error(`${logPrefix} -> ERROR setting selection immediately:`, selError);
             }
        } else {
            log(`${logPrefix} -> Warning: newAbsoluteCaretCol not set, skipping selection update.`);
            }

        } catch (error) {
        log(`${logPrefix} ERROR during manual key handling:`, error);
            console.error('[ep_tables5] Error processing key event update:', error);
        // Maybe return false to allow default as a fallback on error?
        // For now, return true as we prevented default.
        return true;
    }
       
    const endLogTime = Date.now();
    log(`${logPrefix} END (Handled Internal Edit Manually) Key='${evt.key}' Type='${evt.type}' -> Returned true. Duration: ${endLogTime - startLogTime}ms`);
    return true; // We handled the key event

  } // End if(isTypingKey || isInternalBackspace || isInternalDelete)


  // Fallback for any other keys or edge cases not handled above
  const endLogTimeFinal = Date.now();
  log(`${logPrefix} END (Fell Through / Unhandled Case) Key='${evt.key}' Type='${evt.type}'. Allowing default. Duration: ${endLogTimeFinal - startLogTime}ms`);
  // Clear click state if it wasn't handled?
  // if (editor?.ep_tables5_last_clicked) editor.ep_tables5_last_clicked = null;
  return false; // Allow default browser/ACE handling
};

// ───────────────────── ace init + public helpers ─────────────────────
exports.aceInitialized = (h, ctx) => {
  const logPrefix = '[ep_tables5:aceInitialized]';
  log(`${logPrefix} START`, { h, ctx });
  const ed  = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

  // Attach the helper function to editorInfo for later retrieval in aceKeyEvent
  ed.ep_tables5_applyMeta = applyTableLineMetadataAttribute;
  log(`${logPrefix}: Attached applyTableLineMetadataAttribute helper to ed.ep_tables5_applyMeta`);

  // *** REMOVED: Setup mousedown listener via callWithAce (Moved to postAceInit) ***

  // Helper function to apply the metadata attribute to a line
  // Moved to module scope to be accessible by aceKeyEvent
  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo, attributeString = null) {
    const funcName = 'applyTableLineMetadataAttribute';
    log(`${logPrefix}:${funcName}: Applying METADATA attribute to line ${lineNum}`, {tblId, rowIndex, numCols});

    // If attributeString is not provided, construct it. Otherwise, use the provided one.
    if (!attributeString) {
        log(`${logPrefix}:${funcName}: Constructing attribute string as none was provided.`);
        const metadata = {
            tblId: tblId,
            row: rowIndex,
            cols: numCols // Store column count in metadata
        };
        attributeString = JSON.stringify(metadata);
    } else {
         log(`${logPrefix}:${funcName}: Using pre-provided attribute string: ${attributeString}`); // Log the provided string
    }

    log(`${logPrefix}:${funcName}: Metadata Attribute String to apply: ${attributeString}`);
    try {
       // Get a FRESH rep to ensure line length is current after edits
       const liveRep   = editorInfo.ace_getRep();
       const lineEntry = liveRep.lines.atIndex(lineNum);
       if (!lineEntry) {
           log(`${logPrefix}:${funcName}: ERROR - Could not find line entry in provided rep for line number ${lineNum}. Rep lines:`, rep.lines);
           throw new Error(`Could not find line entry for line number ${lineNum}`);
       }
       const lineLength = lineEntry.text.length;
       // Ensure length is at least 1 if line is technically empty but exists
       const effectiveLineLength = Math.max(1, lineLength); 
       log(`${logPrefix}:${funcName}: Line ${lineNum} current text length (from live rep): ${lineLength}. Effective length for attribute: ${effectiveLineLength}. Line text: "${lineEntry.text}"`);

       // Use ace_performDocumentApplyAttributesToRange
       // *** Apply attribute to the FULL line range ***
       const start = [lineNum, 0];
       const end = [lineNum, effectiveLineLength]; // Use potentially corrected length
       log(`${logPrefix}:${funcName}: Applying attribute via ace_performDocumentApplyAttributesToRange to range [${start}]-[${end}]`);
       editorInfo.ace_performDocumentApplyAttributesToRange(start, end, [[ATTR_TABLE_JSON, attributeString]]);
       log(`${logPrefix}:${funcName}: Applied METADATA attribute to line ${lineNum} over range [${start}]-[${end}].`);
    } catch(e) {
        console.error(`[ep_tables5] ${logPrefix}:${funcName}: Error applying metadata attribute on line ${lineNum}:`, e);
        log(`[ep_tables5] ${logPrefix}:${funcName}: Error details:`, { message: e.message, stack: e.stack });
    }
  }

  /** Insert a fresh rows×cols blank table at the caret */
  ed.ace_createTableViaAttributes = (rows = 2, cols = 2) => {
    const funcName = 'ace_createTableViaAttributes';
    log(`${funcName}: START - Refactored Phase 4 (Get Selection Fix)`, { rows, cols });
    rows = Math.max(1, rows); cols = Math.max(1, cols);
    log(`${funcName}: Ensuring minimum 1 row, 1 col.`);

    // --- Phase 1: Prepare Data --- 
    const tblId   = rand();
    log(`${funcName}: Generated table ID: ${tblId}`);
    const initialCellContent = ' '; // Start with a single space per cell
    const lineTxt = Array.from({ length: cols }).fill(initialCellContent).join(DELIMITER);
    log(`${funcName}: Constructed initial line text for ${cols} cols: "${lineTxt}"`);
    const block = Array.from({ length: rows }).fill(lineTxt).join('\n') + '\n';
    log(`${funcName}: Constructed block for ${rows} rows:\n${block}`);

    // Get current selection BEFORE making changes using ace_getRep()
    log(`${funcName}: Getting current representation and selection...`);
    const currentRepInitial = ed.ace_getRep(); 
    if (!currentRepInitial || !currentRepInitial.selStart || !currentRepInitial.selEnd) {
        console.error(`[ep_tables5] ${funcName}: Could not get current representation or selection via ace_getRep(). Aborting.`);
        log(`${funcName}: END - Error getting initial rep/selection`);
        return;
    }
    const start = currentRepInitial.selStart;
    const end = currentRepInitial.selEnd;
    const initialStartLine = start[0]; // Store the starting line number
    log(`${funcName}: Current selection from initial rep:`, { start, end });

    // --- Phase 2: Insert Text Block --- 
    log(`${funcName}: Phase 2 - Inserting text block...`);
    ed.ace_performDocumentReplaceRange(start, end, block);
    log(`${funcName}: Inserted block of delimited text lines.`);
    log(`${funcName}: Requesting text sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20); // Sync text insertion
    log(`${funcName}: Text sync requested.`);

    // --- Phase 3: Apply Metadata Attributes --- 
    log(`${funcName}: Phase 3 - Applying metadata attributes to ${rows} inserted lines...`);
    // Need rep to be updated after text insertion to apply attributes correctly
    const currentRep = ed.ace_getRep(); // Get potentially updated rep
    if (!currentRep || !currentRep.lines) {
        console.error(`[ep_tables5] ${funcName}: Could not get updated rep after text insertion. Cannot apply attributes reliably.`);
        log(`${funcName}: END - Error getting updated rep`);
        // Maybe attempt to continue without rep? Risky.
        return; 
    }
    log(`${funcName}: Fetched updated rep for attribute application.`);

    for (let r = 0; r < rows; r++) {
      const lineNumToApply = initialStartLine + r;
      log(`${funcName}: -> Processing row ${r} on line ${lineNumToApply}`);
      // Call the module-level helper, passing necessary context (currentRep, ed)
      applyTableLineMetadataAttribute(lineNumToApply, tblId, r, cols, currentRep, ed); 
    }
    log(`${funcName}: Finished applying metadata attributes.`);
    log(`${funcName}: Requesting attribute sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20); // Final sync after attributes
    log(`${funcName}: Attribute sync requested.`);

    // --- Phase 4: Set Caret Position --- 
    log(`${funcName}: Phase 4 - Setting final caret position...`);
    const finalCaretLine = initialStartLine + rows; // Line number after the last inserted row
    const finalCaretPos = [finalCaretLine, 0];
    log(`${funcName}: Target caret position:`, finalCaretPos);
    try {
      ed.ace_performSelectionChange(finalCaretPos, finalCaretPos, false);
       log(`${funcName}: Successfully set caret position.`);
    } catch(e) {
       console.error(`[ep_tables5] ${funcName}: Error setting caret position after table creation:`, e);
       log(`[ep_tables5] ${funcName}: Error details:`, { message: e.message, stack: e.stack });
    }

    log(`${funcName}: END - Refactored Phase 4`);
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

// *** ADDED: postAceInit hook for attaching listeners ***
exports.postAceInit = (hookName, ctx) => {
  const func = '[ep_tables5:postAceInit]';
  log(`${func} START`);
  const editorInfo = ctx.ace; // Get editorInfo from context

  if (!editorInfo) {
    console.error(`${func} ERROR: editorInfo (ctx.ace) is not available.`);
    return;
  }

  // Setup mousedown listener via callWithAce
  editorInfo.ace_callWithAce((ace) => {
      const editor = ace.editor;
      const inner = ace.editor.container; // Use the main container

      if (!editor || !inner) {
          console.error(`${func} ERROR: ace.editor or ace.editor.container not found within ace_callWithAce.`);
          return;
      }

      log(`${func} Inside callWithAce for attaching mousedown listeners.`);

      // Initialize shared state on the editor object
      if (!editor.ep_tables5_last_clicked) {
          editor.ep_tables5_last_clicked = null;
          log(`${func} Initialized ace.editor.ep_tables5_last_clicked`);
      }

      log(`${func} Attempting to attach mousedown listener to editor container for cell selection...`);

      inner.addEventListener('mousedown', (evt) => {
          const target = evt.target;
          const mousedownFuncName = '[ep_tables5 mousedown]';
          log(`${mousedownFuncName} RAW MOUSE DOWN detected. Target:`, target);

          // Check if the click is inside a TD of our table
          const clickedTD = target.closest('td');
          const clickedTR = target.closest('tr');
          const clickedTable = target.closest('table.dataTable');

          // Clear previous selection state regardless of where click happened
          if (editor.ep_tables5_last_clicked) {
              log(`${mousedownFuncName} Clearing previous selection info.`);
              // TODO: Add visual class removal if needed
          }
          editor.ep_tables5_last_clicked = null; // Clear state first

          if (clickedTD && clickedTR && clickedTable) {
              log(`${mousedownFuncName} Click detected inside table.dataTable td.`);
              try {
                  const cellIndex = Array.from(clickedTR.children).indexOf(clickedTD);
                  const lineNode = clickedTable.closest('div.ace-line');
                  const tblId = clickedTable.getAttribute('data-tblId');

                  // Ensure ace.rep and ace.rep.lines are available
                  if (!ace.rep || !ace.rep.lines) {
                      console.error(`${mousedownFuncName} ERROR: ace.rep or ace.rep.lines not available inside mousedown listener.`);
                      return;
                  }

                  if (lineNode && lineNode.id && tblId !== null && cellIndex !== -1) {
                      const lineNum = ace.rep.lines.indexOfKey(lineNode.id);
                      if (lineNum !== -1) {
                           // Store the accurately determined cell info
                           // Initialize relative position - might be refined later if needed
                           const clickInfo = { lineNum, tblId, cellIndex, relativePos: 0 }; // Set initial relativePos to 0
                           editor.ep_tables5_last_clicked = clickInfo;
                           log(`${mousedownFuncName} Clicked cell (SUCCESS): Line=${lineNum}, TblId=${tblId}, CellIndex=${cellIndex}. Stored click info:`, clickInfo);

                           // TODO: Add visual class for selection if desired
                           log(`${mousedownFuncName} TEST: Skipped adding/removing selected-table-cell class`);

                      } else {
                          log(`${mousedownFuncName} Clicked cell (ERROR): Could not find line number for node ID: ${lineNode.id}`);
                      }
                  } else {
                       log(`${mousedownFuncName} Clicked cell (ERROR): Missing required info (lineNode, lineNode.id, tblId, or valid cellIndex).`, { lineNode, tblId, cellIndex });
                  }
              } catch (e) {
                  console.error(`${mousedownFuncName} Error processing table cell click:`, e);
                  log(`${mousedownFuncName} Error details:`, { message: e.message, stack: e.stack });
                  editor.ep_tables5_last_clicked = null; // Ensure state is clear on error
              }
          } else {
               log(`${mousedownFuncName} Click was outside a table.dataTable td.`);
          }
      });
      log(`${func} Mousedown listeners for cell selection attached successfully (inside callWithAce).`);

  }, 'tableCellSelectionPostAce', true); // Unique name for callstack

  log(`${func} END`);
};

// END OF FILE