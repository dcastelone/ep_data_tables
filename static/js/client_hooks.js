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
const ZWSP            = '\u200B'; // Zero-Width Space - Keep for reference maybe, but not in delimiter
const DELIMITER       = '|'; // SIMPLIFIED DELIMITER

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
    const segmentHTML = td.innerHTML || ''; // Get raw HTML content
    log(`${funcName} (${node?.id}): Reading rendered TD #${index} innerHTML:`, segmentHTML);
    return segmentHTML;
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
    const cellContent = segment || '&nbsp;'; // Use non-breaking space for empty segments
    log(`${funcName}: Processing segment ${index}. Content:`, segment); // Log segment content
    // Wrap the raw segment HTML in a span for consistency? Or directly in TD? Let's try TD directly first.
    // Maybe add back table-cell-content span later if needed for styling/selection.
    const tdContent = `<td style="${tdStyle}">${cellContent}</td>`;
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

  // *** STARTUP LOGGING ***
  log(`${funcName}: ----- START ----- NodeID: ${nodeId} LineNum: ${lineNum}`);
  if (!node || !nodeId) {
      log(`${funcName}: ERROR - Received invalid node or node without ID. Aborting.`);
      console.error(`[ep_tables5] ${funcName}: Received invalid node or node without ID.`);
      return cb();
  }
  log(`${funcName} NodeID#${nodeId}: Initial node outerHTML:`, node.outerHTML);
  log(`${funcName} NodeID#${nodeId}: Full args object received:`, args);
  // ***********************

  let rowMetadata = null;
  let encodedJsonString = null;

  // --- Log node classes BEFORE searching --- 
  log(`${funcName} NodeID#${nodeId}: Checking classes BEFORE search. Node classList:`, node.classList);
  if (node.children) {
      for (let i = 0; i < node.children.length; i++) {
         log(`${funcName} NodeID#${nodeId}:   Child ${i} (${node.children[i].tagName}) classList:`, node.children[i].classList);
      }
  } else {
       log(`${funcName} NodeID#${nodeId}: Node has no children.`);
  }
  // --- End logging node classes ---

  // --- 1. Find and Parse Metadata Attribute ---
  log(`${funcName} NodeID#${nodeId}: Searching for tbljson-* class on node or children...`);
  // Check the node itself first
  if (node.classList) {
      for (const cls of node.classList) { 
          if (cls.startsWith('tbljson-')) {
              encodedJsonString = cls.substring(8);
              log(`${funcName} NodeID#${nodeId}: Found encoded tbljson on node itself: ${encodedJsonString}`);
              break;
      }
  } 
  }
  // Check children if not found on node
  if (!encodedJsonString && node.children) {
       log(`${funcName} NodeID#${nodeId}: Not found on node, checking children...`);
        for (const child of node.children) {
             if (child.classList) {
                for (const cls of child.classList) {
                 if (cls.startsWith('tbljson-')) {
                     encodedJsonString = cls.substring(8);
                     log(`${funcName} NodeID#${nodeId}: Found encoded tbljson on child ${child.tagName}: ${encodedJsonString}`);
                     break;
                 }
              }
          } 
          if (encodedJsonString) break;
      }
  } 

  // If no attribute found, it's not a table line managed by us (or attribute is missing)
  if (!encodedJsonString) {
      log(`${funcName} NodeID#${nodeId}: No tbljson-* class found. Assuming not a table line for rendering. END.`);
      return cb(); 
  }

  log(`${funcName} NodeID#${nodeId}: Decoding and parsing metadata...`);
  try { 
      const decoded = dec(encodedJsonString); 
      log(`${funcName} NodeID#${nodeId}: Decoded string: ${decoded}`);
      if (!decoded) throw new Error('Decoded string is null or empty.');
      rowMetadata = JSON.parse(decoded);
      log(`${funcName} NodeID#${nodeId}: Parsed rowMetadata:`, rowMetadata);

      // Validate essential metadata
      if (!rowMetadata || typeof rowMetadata.tblId === 'undefined' || typeof rowMetadata.row === 'undefined' || typeof rowMetadata.cols !== 'number') {
          throw new Error('Invalid or incomplete metadata (missing tblId, row, or cols).');
      }
      log(`${funcName} NodeID#${nodeId}: Metadata validated successfully.`);

  } catch(e) { 
      log(`${funcName} NodeID#${nodeId}: FATAL ERROR - Failed to decode/parse/validate tbljson metadata. Rendering cannot proceed.`, e);
      console.error(`[ep_tables5] ${funcName} NodeID#${nodeId}: Failed to decode/parse/validate tbljson.`, encodedJsonString, e);
      // Optionally render an error state in the node?
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_tables5] Error: Invalid table metadata attribute found.</div>';
      log(`${funcName} NodeID#${nodeId}: Rendered error message in node. END.`);
      return cb(); 
  }
  // --- End Metadata Parsing ---

  // --- 2. Get and Parse Line Content ---
  const lineInnerHTML = node.innerHTML;
  log(`${funcName} NodeID#${nodeId}: Parsing line content via Placeholder Replace/Split...`);
  log(`${funcName} NodeID#${nodeId}: Original node.innerHTML:`, lineInnerHTML);

  let htmlSegments = [];
  // Define a unique placeholder unlikely to be in content
  const PLACEHOLDER = '@@EP_TABLES5_DELIM@@'; 

  try {
      // Replace the simple pipe delimiter with the placeholder
      // Need to use a RegExp with the global flag for replaceAll behavior
      // Escape the pipe for RegExp: \|
      const replacedHtml = lineInnerHTML.replace(new RegExp('\\|', 'g'), PLACEHOLDER);
      log(`${funcName} NodeID#${nodeId}: innerHTML after replacing delimiter ('|') with placeholder:`, replacedHtml);

      // Split the string using the placeholder
      htmlSegments = replacedHtml.split(PLACEHOLDER);
      log(`${funcName} NodeID#${nodeId}: Final parsed HTML segments (${htmlSegments.length}) after splitting by placeholder:`, htmlSegments);

      // --- Validation --- 
      if (htmlSegments.length !== rowMetadata.cols) {
          log(`${funcName} NodeID#${nodeId}: WARNING - Parsed segment count (${htmlSegments.length}) does not match metadata cols (${rowMetadata.cols}). Table structure might be incorrect.`);
          console.warn(`[ep_tables5] ${funcName} NodeID#${nodeId}: Parsed segment count (${htmlSegments.length}) mismatch with metadata cols (${rowMetadata.cols}). Segments:`, htmlSegments);
      } else {
          log(`${funcName} NodeID#${nodeId}: Parsed segment count matches metadata cols (${rowMetadata.cols}).`);
      }
  } catch (parseError) {
      log(`${funcName} NodeID#${nodeId}: ERROR during placeholder replace/split parsing. Cannot build table.`, parseError);
      console.error(`[ep_tables5] ${funcName} NodeID#${nodeId}: Error parsing line content via placeholder replace/split.`, parseError);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_tables5] Error: Could not parse table cell content.</div>';
      log(`${funcName} NodeID#${nodeId}: Rendered parse error message in node. END.`);
      return cb();
  }
  // --- End Content Parsing ---

  // --- 3. Build and Render Table ---
  log(`${funcName} NodeID#${nodeId}: Calling buildTableFromDelimitedHTML...`);
      try {
      const newTableHTML = buildTableFromDelimitedHTML(rowMetadata, htmlSegments);
      log(`${funcName} NodeID#${nodeId}: Received new table HTML from helper. Replacing node.innerHTML.`);
      // Replace the node's content entirely with the generated table
      node.innerHTML = newTableHTML;
      log(`${funcName} NodeID#${nodeId}: Successfully replaced node.innerHTML with new table structure.`);
      } catch (renderError) {
      log(`${funcName} NodeID#${nodeId}: ERROR during table building or rendering.`, renderError);
      console.error(`[ep_tables5] ${funcName} NodeID#${nodeId}: Error building/rendering table.`, renderError);
      node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_tables5] Error: Failed to render table structure.</div>';
      log(`${funcName} NodeID#${nodeId}: Rendered build/render error message in node. END.`);
      return cb();
  }
  // --- End Table Building ---

  // *** REMOVED CACHING LOGIC ***
  // The old logic based on tableRowNodes cache is completely removed.

  log(`${funcName}: ----- END ----- NodeID: ${nodeId}`);
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
  log(`[CaretTrace] ${funcName}: START Key='${evt?.key}' Code=${evt?.keyCode} Type=${evt?.type} Modifiers={ctrl:${evt?.ctrlKey},alt:${evt?.altKey},meta:${evt?.metaKey},shift:${evt?.shiftKey}}`, { selStart: rep?.selStart, selEnd: rep?.selEnd });

  if (!rep || !rep.selStart || !editorInfo || !evt || !docManager) {
    log(`[CaretTrace] ${funcName}: Skipping - Missing critical context.`);
    return false;
  }

  const currentLineNum = rep.selStart[0];
  const currentCol = rep.selStart[1]; // Absolute column in the line model
  log(`[CaretTrace] ${funcName}: Caret Line=${currentLineNum}, Col=${currentCol}`);

  // --- Check if the current line is a table line ---
  let isTableLine = false;
  let tableMetadata = null;
  try {
    const lineAttrString = docManager.getAttributeOnLine(currentLineNum, ATTR_TABLE_JSON);
    if (lineAttrString) {
        tableMetadata = JSON.parse(lineAttrString);
        if (tableMetadata && typeof tableMetadata.cols === 'number') {
             isTableLine = true;
             log(`[CaretTrace] ${funcName}: Current line IS a table line. Metadata:`, tableMetadata);
        } else {
             log(`[CaretTrace] ${funcName}: Line has attribute, but metadata invalid/missing cols.`);
        }
    } else {
        // log(`[CaretTrace] ${funcName}: Current line has no ${ATTR_TABLE_JSON} attribute.`);
    }
  } catch(e) {
    console.error(`[ep_tables5] ${funcName}: Error checking/parsing line attribute.`, e);
  }

  if (!isTableLine) {
    log(`[CaretTrace] ${funcName}: Not a table line. Allowing default. END (Default)`);
    return false; // Not our line, let Etherpad handle it
  }
  // --- End Table Line Check ---

  // --- Determine Target Cell and Relative Caret Position ---
  let targetCellIndex = -1;
  let relativeCaretPos = -1;
  let precedingCellsOffset = 0; // Sum of lengths + delimiters before target cell
  let cellStartCol = 0; // Absolute column where the target cell starts

  const lineText = rep.lines.atIndex(currentLineNum)?.text || '';
  log(`[CaretTrace] ${funcName}: Line text: "${lineText}"`);
  const cellTexts = lineText.split(DELIMITER);
  log(`[CaretTrace] ${funcName}: Split line into segments:`, cellTexts);

  // Basic validation - ensure we have segments if it's a table line
  if (cellTexts.length === 0 && lineText.length > 0) {
       log(`[CaretTrace] ${funcName}: ERROR - Line identified as table but splitting text resulted in zero segments. Text: "${lineText}"`);
       return false; // Avoid errors, let default (might break things)
  }
  // Ensure segment count matches metadata if possible
  if (cellTexts.length !== tableMetadata.cols) {
      log(`[CaretTrace] ${funcName}: WARNING - Cell text segment count (${cellTexts.length}) mismatch with metadata cols (${tableMetadata.cols}). Proceeding cautiously.`);
      // Allow proceeding, but cell index calculation might be off
  }

  let currentOffset = 0;
  for (let i = 0; i < cellTexts.length; i++) {
      const cellLength = cellTexts[i].length;
      const cellEndCol = currentOffset + cellLength;
       log(`[CaretTrace] ${funcName}: Checking cell ${i}: Text="${cellTexts[i]}", Length=${cellLength}, StartsAt=${currentOffset}, EndsAt=${cellEndCol}`);

      // Check if caret is within this cell OR exactly at the start of the next cell (treat as end of current)
      if (currentCol >= currentOffset && currentCol <= cellEndCol) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          relativeCaretPos = currentCol - currentOffset;
          precedingCellsOffset = currentOffset;
          log(`[CaretTrace] ${funcName}: --> Caret is in Cell ${targetCellIndex} (StartsAt ${cellStartCol}). Relative Pos: ${relativeCaretPos}`);
          break; // Found the cell
      }
      // Move offset past the current cell and its delimiter
      currentOffset += cellLength + DELIMITER.length;
  }

  if (targetCellIndex === -1) {
      if (currentCol === lineText.length && cellTexts.length > 0) {
            targetCellIndex = cellTexts.length - 1;
            cellStartCol = currentOffset - (cellTexts[targetCellIndex].length + DELIMITER.length); // Calculate start of last cell
            relativeCaretPos = cellTexts[targetCellIndex].length; // Caret is at the end
            precedingCellsOffset = cellStartCol;
            log(`[CaretTrace] ${funcName}: --> Caret detected at END of last cell (${targetCellIndex}). Relative Pos: ${relativeCaretPos}`);
      } else {
        log(`[CaretTrace] ${funcName}: FAILED to determine target cell for caret col ${currentCol}. Aborting event handling.`);
        return false; // Let default handle
      }
  }
  // --- End Cell/Position Calculation ---

  // --- Define Key Types --- 
  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete' || evt.keyCode === 46;
  const isBackspaceKey = evt.key === 'Backspace' || evt.keyCode === 8;
  const isNavigationKey = [33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode);
  const isTabKey = evt.key === 'Tab';
  // Add other keys to ignore explicitly if needed
  // const isEnterKey = evt.key === 'Enter';

  log(`[CaretTrace] ${funcName}: Key classification: Typing=${isTypingKey}, Backspace=${isBackspaceKey}, Delete=${isDeleteKey}, Nav=${isNavigationKey}, Tab=${isTabKey}`);

  // --- Handle Keys --- 

  // 1. Allow most navigation keys by default (except maybe Tab)
  if (isNavigationKey && !isTabKey) {
      log(`[CaretTrace] ${funcName}: Allowing navigation key: ${evt.key}`);
      return false;
  }

  // 2. Handle Tab - Prevent default, implement navigation later
  if (isTabKey) {
     log(`[CaretTrace] ${funcName}: Tab key pressed in cell - Preventing default. Needs cell navigation impl.`);
     evt.preventDefault();
     return true;
  }

  // 3. Intercept destructive keys at boundaries
  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0; // Handle potential undefined cell
  if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex > 0) {
      log(`[CaretTrace] ${funcName}: Intercepted Backspace at start of cell ${targetCellIndex}. Preventing default.`);
      evt.preventDefault();
      return true;
  }
  if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex < tableMetadata.cols - 1) {
      log(`[CaretTrace] ${funcName}: Intercepted Delete at end of cell ${targetCellIndex}. Preventing default.`);
      evt.preventDefault();
      return true;
  }

  // 4. Handle Typing/Backspace/Delete WITHIN a cell using minimal replace range
  const shouldModifyCellText = isTypingKey || isBackspaceKey || isDeleteKey; // Redefine based on keys we handle here

  if (shouldModifyCellText && targetCellIndex >= 0) { // Ensure we have a valid target cell
      log(`[CaretTrace] ${funcName}: HANDLED KEY (Cell Modify) - Key='${evt.key}' Type='${evt.type}' CellIndex=${targetCellIndex}`);

      if (evt.type !== 'keydown') {
          log(`[CaretTrace] ${funcName}: Ignoring non-keydown event type ('${evt.type}') for cell modification.`);
          return false;
      }

      evt.preventDefault();
      log(`[CaretTrace] ${funcName}: Prevented default browser action.`);

      let textModified = false;
      let newRelativeCaretPos = relativeCaretPos;
      const currentCellText = cellTexts[targetCellIndex];
      let newCellText = currentCellText;

      log(`[CaretTrace] ${funcName}: -> Current cell text: "${currentCellText}" Relative caret: ${relativeCaretPos}`);

      // Apply modification
      if (isTypingKey) {
          newCellText = currentCellText.slice(0, relativeCaretPos) + evt.key + currentCellText.slice(relativeCaretPos);
          newRelativeCaretPos++;
          textModified = true;
          log(`[CaretTrace] ${funcName}: -> Applied Typing: New cell text: "${newCellText}" New relative caret: ${newRelativeCaretPos}`);
      } else if (isBackspaceKey) {
          // We already handled the boundary case (relativeCaretPos === 0)
          if (relativeCaretPos > 0) {
              newCellText = currentCellText.slice(0, relativeCaretPos - 1) + currentCellText.slice(relativeCaretPos);
              newRelativeCaretPos--;
              textModified = true;
              log(`[CaretTrace] ${funcName}: -> Applied Backspace: New cell text: "${newCellText}" New relative caret: ${newRelativeCaretPos}`);
          } else {
              log(`[CaretTrace] ${funcName}: -> Backspace at start of cell (but cell 0 or boundary handled). No change here.`);
          }
      } else if (isDeleteKey) {
          // We already handled the boundary case (relativeCaretPos === currentCellTextLength)
          if (relativeCaretPos < currentCellText.length) {
              newCellText = currentCellText.slice(0, relativeCaretPos) + currentCellText.slice(relativeCaretPos + 1);
              textModified = true;
              log(`[CaretTrace] ${funcName}: -> Applied Delete: New cell text: "${newCellText}" New relative caret: ${newRelativeCaretPos}`);
          } else {
              log(`[CaretTrace] ${funcName}: -> Delete at end of cell (but last cell or boundary handled). No change here.`);
          }
      }

      // If text was potentially modified, perform minimal replace
      if (textModified) {
          log(`[CaretTrace] ${funcName}: Text modified. Updating document via MINIMAL range replace...`);

          const cellAbsStart = precedingCellsOffset;
          const cellAbsEnd = precedingCellsOffset + currentCellText.length;
          const replaceRangeStart = [currentLineNum, cellAbsStart];
          const replaceRangeEnd = [currentLineNum, cellAbsEnd];

          log(`[CaretTrace] ${funcName}: -> Replacing range [${replaceRangeStart}]-[${replaceRangeEnd}] with new cell text: "${newCellText}"`);

          try {
              // Perform the minimal replacement
              editorInfo.ace_performDocumentReplaceRange(replaceRangeStart, replaceRangeEnd, newCellText);
              log(`[CaretTrace] ${funcName}: -> ace_performDocumentReplaceRange (minimal) called.`);

              // <<< RE-ADD ATTRIBUTE RE-APPLICATION >>>
              log(`[CaretTrace] ${funcName}: -> Attempting to re-apply metadata attribute via editorInfo context...`);
              const applyHelper = editorInfo.ep_tables5_applyMeta; // Retrieve helper attached in aceInitialized
              if (applyHelper && typeof applyHelper === 'function') {
                  // Get the UPDATED representation after text change
                  const updatedRep = editorInfo.ace_getRep();
                  if (!updatedRep) {
                      console.error(`[ep_tables5] ${funcName}: Failed to get updated rep after text replace. Cannot re-apply attribute reliably.`);
                  } else {
                       log(`[CaretTrace] ${funcName}: -> Got updated rep. Calling helper via editorInfo.`);
                       // Call the retrieved helper, passing the necessary context (updatedRep, editorInfo)
                       applyHelper(currentLineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, updatedRep, editorInfo);
                       log(`[CaretTrace] ${funcName}: -> Metadata attribute re-applied via helper.`);
                  }
              } else {
                   console.error(`[ep_tables5] ${funcName}: Could not find applyTableLineMetadataAttribute helper on editorInfo.ep_tables5_applyMeta.`);
                   log(`[CaretTrace] ${funcName}: -> FAILED to re-apply metadata attribute.`);
              }
              // <<< END ATTRIBUTE RE-APPLICATION >>>

              // Calculate and set the new caret position immediately
              const newAbsoluteCaretCol = precedingCellsOffset + newRelativeCaretPos;
              const newCaretPos = [currentLineNum, newAbsoluteCaretCol];
              log(`[CaretTrace] ${funcName}: -> Calculated new absolute caret: Col=${newAbsoluteCaretCol}`);
              try {
                  log(`[CaretTrace] ${funcName}: -> Setting selection immediately to:`, newCaretPos);
                  editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                  log(`[CaretTrace] ${funcName}: -> Selection set successfully.`);

                  // <<< ADDED: Verify attribute presence immediately after setting >>>
                  try {
                     const verifyAttr = docManager.getAttributeOnLine(currentLineNum, ATTR_TABLE_JSON);
                     log(`[CaretTrace] ${funcName}: -> VERIFY attribute on line ${currentLineNum} immediately after setting: ${verifyAttr ? 'FOUND' : 'NOT FOUND'}`, verifyAttr || '(null)');
                  } catch (verifyError) {
                     log(`[CaretTrace] ${funcName}: -> ERROR verifying attribute immediately after setting:`, verifyError);
                  }
                  // <<< END VERIFY >>>

              } catch (selError) {
                   console.error(`[CaretTrace] ${funcName}: ERROR during immediate selection setting:`, selError);
                   log(`[CaretTrace] ${funcName}: Error details:`, { message: selError.message, stack: selError.stack });
              }

          } catch (replaceError) {
              log(`[CaretTrace] ${funcName}: ERROR during minimal ace_performDocumentReplaceRange:`, replaceError);
              console.error('[ep_tables5] Error replacing cell text range:', replaceError);
              // Might need to return false here to allow default handling as fallback?
          }
      } else {
          log(`[CaretTrace] ${funcName}: No text modification needed for this key.`);
      }

      const endLogTime = Date.now();
      log(`[CaretTrace] ${funcName}: END (Handled Key) -> Returned true. Duration: ${endLogTime - startLogTime}ms`);
      return true; // Indicate we handled the key event
  }

  // --- Fallback for unhandled keys --- 
  const endLogTimeUnhandled = Date.now();
  log(`[CaretTrace] ${funcName}: Key not explicitly handled (or not a modification key in a valid cell). Allowing default. END (Default Allowed). Duration: ${endLogTimeUnhandled - startLogTime}ms`);
  return false;
};

// ───────────────────── ace init + public helpers ─────────────────────
exports.aceInitialized = (h, ctx) => {
  log('aceInitialized: START', { h, ctx });
  const ed  = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

  // Attach the helper function to editorInfo for later retrieval in aceKeyEvent
  ed.ep_tables5_applyMeta = applyTableLineMetadataAttribute;
  log('aceInitialized: Attached applyTableLineMetadataAttribute helper to ed.ep_tables5_applyMeta');

  // Helper function to apply the metadata attribute to a line
  // Moved to module scope to be accessible by aceKeyEvent
  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo) {
    const funcName = 'applyTableLineMetadataAttribute';
    log(`${funcName}: Applying METADATA attribute to line ${lineNum}`, {tblId, rowIndex, numCols});
    const metadata = {
        tblId: tblId,
        row: rowIndex,
        cols: numCols // Store column count in metadata
    };
    const attributeString = JSON.stringify(metadata);
    log(`${funcName}: Metadata Attribute String: ${attributeString}`);
    try {
       // Get the current text length of the line
       const lineEntry = rep.lines.atIndex(lineNum);
       if (!lineEntry) {
           throw new Error(`Could not find line entry for line number ${lineNum}`);
       }
       const lineLength = lineEntry.text.length;
       log(`${funcName}: Line ${lineNum} current text length: ${lineLength}. Line text: "${lineEntry.text}"`);

       // Use ace_performDocumentApplyAttributesToRange for potentially better hook triggering
       // Apply attribute to the ENTIRE line text range
       const start = [lineNum, 0];
       const end = [lineNum, Math.max(1, lineLength)]; // Ensure range is at least 1 char wide
       log(`${funcName}: Applying attribute via ace_performDocumentApplyAttributesToRange to FULL range [${start}]-[${end}]`);
       editorInfo.ace_performDocumentApplyAttributesToRange(start, end, [[ATTR_TABLE_JSON, attributeString]]);
       log(`${funcName}: Applied METADATA attribute to line ${lineNum} over full range.`);
    } catch(e) {
        console.error(`[ep_tables5] ${funcName}: Error applying metadata attribute on line ${lineNum}:`, e);
        log(`[ep_tables5] ${funcName}: Error details:`, { message: e.message, stack: e.stack });
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

// END OF FILE
