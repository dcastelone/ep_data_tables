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
        if (typeof atob === 'function') {
            return atob(str); // Browser environment
        } else if (typeof Buffer === 'function') {
            // Node.js environment
            return Buffer.from(str, 'base64').toString('utf8');
        } else {
            console.error('[ep_tables5] Base64 decoding function (atob or Buffer) not found.');
            return null;
        }
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
  // !!!!! VERY IMPORTANT DIAGNOSTIC LOG !!!!!
  console.log(`[ep_tables5:client_hooks] ${funcName}: ENTERING HOOK. Node Tag: ${node?.tagName}, Node Class: ${node?.className}, Node Type: ${node?.nodeType}`);
  if (node && typeof node.getAttribute === 'function') {
    console.log(`[ep_tables5:client_hooks] ${funcName}: Node outerHTML (first 250 chars): ${(node.outerHTML || '').substring(0,250)}`);
  }

  // Log entry point
  log(`${funcName}: START - Processing node ID: ${node?.id}, Class: ${node?.className}, Tag: ${node?.tagName}`);
  if (node?.tagName === 'DIV') {
    log(`${funcName} (${node?.id}): Server-side DIV check. OuterHTML: ${node.outerHTML ? node.outerHTML.substring(0, 200) + (node.outerHTML.length > 200 ? '...':'') : '[No outerHTML]'}`);
    // Check if this DIV contains our target SPAN (Now P)
    const childP = node.querySelector('p[class*="tbljson-"]');
    if (childP) {
        log(`${funcName} (${node?.id}): DIV contains a tbljson- P. Will process P directly if it comes up.`);
    }
  } else if (node?.tagName === 'P') {
     log(`${funcName} (${node?.id}): Server-side P check. Class: ${node?.className}. OuterHTML: ${node.outerHTML ? node.outerHTML.substring(0, 200) + (node.outerHTML.length > 200 ? '...':'') : '[No outerHTML]'}`);
  }

  // !!!!! DIAGNOSTIC LOG BEFORE P CHECK !!!!!
  // Removed the P check diagnostic log as the P-specific block is being removed.

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
  const rep = ctx?.rep; // Ensure rep is defined in this scope if not already
  const docAttrManager = ctx?.documentAttributeManager; // Ensure docAttrManager is defined

  const lineNum = rep?.lines?.indexOfKey(node.id); // Get line number
  log(`${funcName} (${node?.id}): Determined line number: ${lineNum}`);

  if (typeof lineNum !== 'number' || lineNum < 0 || !docAttrManager) {
    console.error(`[ep_tables5] ${funcName} (${node?.id}): Could not get valid line number (${lineNum}) or docAttrManager. Aborting table collection for this line.`);
    log(`${funcName} (${node?.id}): Aborting custom collection due to missing lineNum/docAttrManager. Current node outerHTML:`, node.outerHTML?.substring(0, 500));
    return; // Allow default handlers if we can't get line info or manager
  }

         try { 
    const existingAttrString = docAttrManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
    log(`${funcName} (${node?.id}): For line ${lineNum}, retrieved existing ATTR_TABLE_JSON string:`, existingAttrString);
    if (existingAttrString) {
      existingMetadata = JSON.parse(existingAttrString);
      log(`${funcName} (${node?.id}): Parsed existing metadata for line ${lineNum}:`, existingMetadata);
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
  
  // Helper function to recursively search for tbljson class in all descendants
  function findTbljsonClass(element) {
    // Check the element itself
    if (element.classList) {
      for (const cls of element.classList) {
          if (cls.startsWith('tbljson-')) {
          return cls.substring(8);
        }
      }
    }
    // Recursively check all descendants
    if (element.children) {
      for (const child of element.children) {
        const found = findTbljsonClass(child);
        if (found) return found;
      }
    }
    return null;
  }

  // Search for tbljson class starting from the node
  encodedJsonString = findTbljsonClass(node);
  
  if (encodedJsonString) {
    log(`${logPrefix} NodeID#${nodeId}: Found encoded tbljson class: ${encodedJsonString}`);
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
  // ALWAYS get the innerHTML of the line div itself to preserve all styling spans and attributes.
  // This innerHTML is set by Etherpad based on the line's current text in atext and includes
  // all the span elements with author colors, bold, italic, and other styling.
  // For an imported line's first render, atext is "Cell1|Cell2", so node.innerHTML will be "Cell1|Cell2".
  // For a natively created line, node.innerHTML is also "Cell1|Cell2".
  // After an edit, aceKeyEvent updates atext, and node.innerHTML reflects that new "EditedCell1|Cell2" string.
  // When styling is applied, it will include spans like: <span class="author-xxx bold">Cell1</span>|<span class="author-yyy italic">Cell2</span>
  const delimitedTextFromLine = node.innerHTML;
  log(`${logPrefix} NodeID#${nodeId}: Using node.innerHTML for delimited text to preserve styling. Value: "${(delimitedTextFromLine || '').substring(0,100)}..."`);
  
  // The DELIMITER const is defined at the top of this file.
  const htmlSegments = (delimitedTextFromLine || '').split(DELIMITER); 

  log(`${logPrefix} NodeID#${nodeId}: Parsed HTML segments (${htmlSegments.length}):`, htmlSegments.map(s => (s || '').substring(0,50) + (s && s.length > 50 ? '...' : '')));

  // --- Validation --- 
  if (htmlSegments.length !== rowMetadata.cols) {
      log(`${logPrefix} NodeID#${nodeId}: WARNING - Parsed segment count (${htmlSegments.length}) does not match metadata cols (${rowMetadata.cols}). Table structure might be incorrect.`);
      console.warn(`[ep_tables5] ${funcName} NodeID#${nodeId}: Parsed segment count (${htmlSegments.length}) mismatch with metadata cols (${rowMetadata.cols}). Segments:`, htmlSegments);
  } else {
      log(`${logPrefix} NodeID#${nodeId}: Parsed segment count matches metadata cols (${rowMetadata.cols}).`);
  }

  // --- 3. Build and Render Table ---
  log(`${logPrefix} NodeID#${nodeId}: Calling buildTableFromDelimitedHTML...`);
      try {
      const newTableHTML = buildTableFromDelimitedHTML(rowMetadata, htmlSegments);
      log(`${logPrefix} NodeID#${nodeId}: Received new table HTML from helper. Replacing content.`);
      
      // Find the element that contains the tbljson class to determine if we need to preserve block wrappers
      function findTbljsonElement(element) {
        // Check if this element has the tbljson class
        if (element.classList) {
          for (const cls of element.classList) {
            if (cls.startsWith('tbljson-')) {
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
      
      const tbljsonElement = findTbljsonElement(node);
      
      // If we found a tbljson element and it's nested in a block element, 
      // we need to preserve the block wrapper while replacing the content
      if (tbljsonElement && tbljsonElement.parentElement && tbljsonElement.parentElement !== node) {
        // Check if the parent is a block-level element that should be preserved
        const parentTag = tbljsonElement.parentElement.tagName.toLowerCase();
        const blockElements = ['center', 'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'right', 'left', 'ul', 'ol', 'li', 'code'];
        
        if (blockElements.includes(parentTag)) {
          log(`${logPrefix} NodeID#${nodeId}: Preserving block element ${parentTag} and replacing its content with table.`);
          tbljsonElement.parentElement.innerHTML = newTableHTML;
        } else {
          log(`${logPrefix} NodeID#${nodeId}: Parent element ${parentTag} is not a block element, replacing entire node content.`);
          node.innerHTML = newTableHTML;
        }
      } else {
      // Replace the node's content entirely with the generated table
        log(`${logPrefix} NodeID#${nodeId}: No nested block element found, replacing entire node content.`);
      node.innerHTML = newTableHTML;
      }
      
      log(`${logPrefix} NodeID#${nodeId}: Successfully replaced content with new table structure.`);
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
    // Add debugging to see what's happening with attribute retrieval
    log(`${logPrefix} DEBUG: Attempting to get ${ATTR_TABLE_JSON} attribute from line ${reportedLineNum}`);
    lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
    log(`${logPrefix} DEBUG: getAttributeOnLine returned: ${lineAttrString ? `"${lineAttrString}"` : 'null/undefined'}`);
    
    // Also check if there are any attributes on this line at all
    if (typeof docManager.getAttributesOnLine === 'function') {
      try {
        const allAttribs = docManager.getAttributesOnLine(reportedLineNum);
        log(`${logPrefix} DEBUG: All attributes on line ${reportedLineNum}:`, allAttribs);
      } catch(e) {
        log(`${logPrefix} DEBUG: Error getting all attributes:`, e);
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
            log(`${logPrefix} DEBUG: Found table in DOM without attribute! TblId=${domTblId}, Row=${domRow}`);
            // Try to reconstruct the metadata from DOM
            const domCells = tableInDOM.querySelectorAll('td');
            if (domTblId && domRow !== null && domCells.length > 0) {
              log(`${logPrefix} DEBUG: Attempting to reconstruct metadata from DOM...`);
              const reconstructedMetadata = {
                tblId: domTblId,
                row: parseInt(domRow, 10),
                cols: domCells.length
              };
              lineAttrString = JSON.stringify(reconstructedMetadata);
              log(`${logPrefix} DEBUG: Reconstructed metadata: ${lineAttrString}`);
            }
          }
        }
      } catch(e) {
        log(`${logPrefix} DEBUG: Error checking DOM for table:`, e);
      }
    }
    
    if (lineAttrString) {
        tableMetadata = JSON.parse(lineAttrString);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
             log(`${logPrefix} Line ${reportedLineNum} has attribute, but metadata invalid/missing cols.`);
             tableMetadata = null; // Ensure it's null if invalid
        }
    } else {
        log(`${logPrefix} DEBUG: No ${ATTR_TABLE_JSON} attribute found on line ${reportedLineNum}`);
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
          log(`${logPrefix} DEBUG: Getting ${ATTR_TABLE_JSON} attribute from stored line ${lastClick.lineNum}`);
          storedLineAttrString = docManager.getAttributeOnLine(lastClick.lineNum, ATTR_TABLE_JSON);
          log(`${logPrefix} DEBUG: Stored line attribute result: ${storedLineAttrString ? `"${storedLineAttrString}"` : 'null/undefined'}`);
          
          if (storedLineAttrString) {
            storedLineMetadata = JSON.parse(storedLineAttrString);
            log(`${logPrefix} DEBUG: Parsed stored metadata:`, storedLineMetadata);
          }
          
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

  // --- START NEW: Handle Highlight Deletion/Replacement ---
  const selStartActual = rep.selStart;
  const selEndActual = rep.selEnd;
  const hasSelection = selStartActual[0] !== selEndActual[0] || selStartActual[1] !== selEndActual[1];

  if (hasSelection) {
    log(`${logPrefix} [selection] Active selection detected. Start:[${selStartActual[0]},${selStartActual[1]}], End:[${selEndActual[0]},${selEndActual[1]}]`);
    log(`${logPrefix} [caretTrace] [selection] Initial rep.selStart: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

    if (selStartActual[0] !== currentLineNum || selEndActual[0] !== currentLineNum) {
      log(`${logPrefix} [selection] Selection spans multiple lines (${selStartActual[0]}-${selEndActual[0]}) or is not on the current focused table line (${currentLineNum}). Preventing default action.`);
      evt.preventDefault();
      return true; 
    }

    const selectionStartColInLine = selStartActual[1];
    const selectionEndColInLine = selEndActual[1];

    const currentCellFullText = cellTexts[targetCellIndex] || '';
    // cellStartCol is already defined and calculated based on trustedLastClick or fallback
    const cellContentStartColInLine = cellStartCol;
    const cellContentEndColInLine = cellStartCol + currentCellFullText.length;

    log(`${logPrefix} [selection] Cell context for selection: targetCellIndex=${targetCellIndex}, cellStartColInLine=${cellContentStartColInLine}, cellEndColInLine=${cellContentEndColInLine}, currentCellFullText='${currentCellFullText}'`);

    const isSelectionEntirelyWithinCell =
      selectionStartColInLine >= cellContentStartColInLine &&
      selectionEndColInLine <= cellContentEndColInLine;

    log(`${logPrefix} [selection] Checking if selection [${selectionStartColInLine}-${selectionEndColInLine}] is entirely within cell [${cellContentStartColInLine}-${cellContentEndColInLine}]. Result: ${isSelectionEntirelyWithinCell}`);

    if (!isSelectionEntirelyWithinCell) {
      log(`${logPrefix} [selection] Selection is NOT entirely within cell ${targetCellIndex} or spans delimiters. Preventing default action to protect table structure.`);
      evt.preventDefault();
      return true;
    }

    const isCurrentKeyDelete = evt.key === 'Delete' || evt.keyCode === 46;
    const isCurrentKeyBackspace = evt.key === 'Backspace' || evt.keyCode === 8;
    // Check if it's a printable character, not a modifier
    const isCurrentKeyTyping = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;


    if (isSelectionEntirelyWithinCell && (isCurrentKeyDelete || isCurrentKeyBackspace || isCurrentKeyTyping)) {
      log(`${logPrefix} [selection] Handling key='${evt.key}' (Type: ${evt.type}) for valid intra-cell selection.`);
      
      if (evt.type !== 'keydown') {
        log(`${logPrefix} [selection] Ignoring non-keydown event type ('${evt.type}') for selection handling. Allowing default.`);
        return false; 
      }
      evt.preventDefault();

      const rangeStart = [currentLineNum, selectionStartColInLine];
      const rangeEnd = [currentLineNum, selectionEndColInLine];
      let replacementText = '';
      let newAbsoluteCaretCol = selectionStartColInLine;
      const repBeforeEdit = editorInfo.ace_getRep(); // Get rep before edit for attribute helper
      log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

      if (isCurrentKeyTyping) {
        replacementText = evt.key;
        newAbsoluteCaretCol = selectionStartColInLine + replacementText.length;
        log(`${logPrefix} [selection] -> Replacing selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]] with text '${replacementText}'`);
      } else { // Delete or Backspace
        log(`${logPrefix} [selection] -> Deleting selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]]`);
      }

      try {
        // const repBeforeEdit = editorInfo.ace_getRep(); // Get rep before edit for attribute helper - MOVED UP
        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, replacementText);
        const repAfterReplace = editorInfo.ace_getRep();
        log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


        log(`${logPrefix} [selection] -> Re-applying tbljson line attribute...`);
        const applyHelper = editorInfo.ep_tables5_applyMeta;
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) {
          const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;
          applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
          log(`${logPrefix} [selection] -> tbljson line attribute re-applied (using rep before edit).`);
        } else {
          console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
          const currentRepFallback = editorInfo.ace_getRep();
          if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
            log(`${logPrefix} [selection] -> Retrying attribute application with current rep...`);
            applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
            log(`${logPrefix} [selection] -> tbljson line attribute re-applied (using current rep fallback).`);
          } else {
            console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute even with fallback rep.`);
          }
        }

        log(`${logPrefix} [selection] -> Setting selection/caret to: [${currentLineNum}, ${newAbsoluteCaretCol}]`);
        log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterSelectionChange = editorInfo.ace_getRep();
        log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);
        
        // Add sync hint AFTER setting selection
        editorInfo.ace_fastIncorp(1);
        const repAfterFastIncorp = editorInfo.ace_getRep();
        log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
        log(`${logPrefix} [selection] -> Requested sync hint (fastIncorp 1).`);

        // --- Re-assert selection --- 
        log(`${logPrefix} [caretTrace] [selection] Attempting to re-assert selection post-fastIncorp to [${currentLineNum}, ${newAbsoluteCaretCol}]`);
        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterReassert = editorInfo.ace_getRep();
        log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

        const newRelativePos = newAbsoluteCaretCol - cellStartCol;
        if (editor) {
            editor.ep_tables5_last_clicked = {
                lineNum: currentLineNum,
                tblId: metadataForTargetLine.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos
            };
            log(`${logPrefix} [selection] -> Updated stored click/caret info:`, editor.ep_tables5_last_clicked);
        } else {
            log(`${logPrefix} [selection] -> Editor instance not found, cannot update ep_tables5_last_clicked.`);
        }
        
        log(`${logPrefix} END [selection] (Handled highlight modification) Key='${evt.key}' Type='${evt.type}'. Duration: ${Date.now() - startLogTime}ms`);
        return true;
      } catch (error) {
        log(`${logPrefix} [selection] ERROR during highlight modification:`, error);
        console.error('[ep_tables5] Error processing highlight modification:', error);
        return true; // Still return true as we prevented default.
      }
    }
  }
  // --- END NEW: Handle Highlight Deletion/Replacement ---

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
    log(`${logPrefix} [caretTrace] Initial rep.selStart for internal edit: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

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
        log(`${logPrefix} [caretTrace] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

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
        const repAfterReplace = editorInfo.ace_getRep();
        log(`${logPrefix} [caretTrace] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


        // *** CRITICAL: Re-apply the line attribute after ANY modification ***
        log(`${logPrefix} -> Re-applying tbljson line attribute...`);
        const applyHelper = editorInfo.ep_tables5_applyMeta; 
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) { 
             // Pass the original lineAttrString if available AND if it belongs to the currentLineNum
             const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;
             applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
             log(`${logPrefix} -> tbljson line attribute re-applied (using rep before edit).`);
                } else {
             console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
             const currentRepFallback = editorInfo.ace_getRep();
             if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
                 log(`${logPrefix} -> Retrying attribute application with current rep...`);
                 applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager); // Cannot guarantee old attr string is valid here
                 log(`${logPrefix} -> tbljson line attribute re-applied (using current rep fallback).`);
            } else {
                  console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute even with fallback rep.`);
             }
        }
        
        // Set caret position immediately
        if (newAbsoluteCaretCol >= 0) {
             const newCaretPos = [currentLineNum, newAbsoluteCaretCol]; // Use the trusted currentLineNum
             log(`${logPrefix} -> Setting selection immediately to:`, newCaretPos);
             log(`${logPrefix} [caretTrace] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
             try {
                editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                const repAfterSelectionChange = editorInfo.ace_getRep();
                log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);
                log(`${logPrefix} -> Selection set immediately.`);

                // Add sync hint AFTER setting selection
                editorInfo.ace_fastIncorp(1); 
                const repAfterFastIncorp = editorInfo.ace_getRep();
                log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
                log(`${logPrefix} -> Requested sync hint (fastIncorp 1).`);

                // --- Re-assert selection --- 
                const targetCaretPosForReassert = [currentLineNum, newAbsoluteCaretCol];
                log(`${logPrefix} [caretTrace] Attempting to re-assert selection post-fastIncorp to [${targetCaretPosForReassert[0]}, ${targetCaretPosForReassert[1]}]`);
                editorInfo.ace_performSelectionChange(targetCaretPosForReassert, targetCaretPosForReassert, false);
                const repAfterReassert = editorInfo.ace_getRep();
                log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

                // Store the updated caret info for the next event
                const newRelativePos = newAbsoluteCaretCol - cellStartCol;
                editor.ep_tables5_last_clicked = {
                    lineNum: currentLineNum, 
                    tblId: metadataForTargetLine.tblId,
                    cellIndex: targetCellIndex,
                    relativePos: newRelativePos
                };
                log(`${logPrefix} -> Updated stored click/caret info:`, editor.ep_tables5_last_clicked);
                log(`${logPrefix} [caretTrace] Updated ep_tables5_last_clicked. Line=${editor.ep_tables5_last_clicked.lineNum}, Cell=${editor.ep_tables5_last_clicked.cellIndex}, RelPos=${editor.ep_tables5_last_clicked.relativePos}`);


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
  log(`${logPrefix} [caretTrace] Final rep.selStart at end of aceKeyEvent (if unhandled): Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);
  return false; // Allow default browser/ACE handling
};

// ───────────────────── ace init + public helpers ─────────────────────
exports.aceInitialized = (h, ctx) => {
  const logPrefix = '[ep_tables5:aceInitialized]';
  log(`${logPrefix} START`, { hook_name: h, context: ctx });
  const ed = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

  log(`${logPrefix} Attaching ep_tables5_applyMeta helper to editorInfo.`);
  ed.ep_tables5_applyMeta = applyTableLineMetadataAttribute;
  log(`${logPrefix}: Attached applyTableLineMetadataAttribute helper to ed.ep_tables5_applyMeta successfully.`);

  // Store the documentAttributeManager reference for later use
  log(`${logPrefix} Storing documentAttributeManager reference on editorInfo.`);
  ed.ep_tables5_docManager = docManager;
  log(`${logPrefix}: Stored documentAttributeManager reference as ed.ep_tables5_docManager.`);

  // *** ADDED: Paste event listener ***
  log(`${logPrefix} Preparing to attach paste listener via ace_callWithAce.`);
  ed.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = '[ep_tables5:aceInitialized:callWithAceForPaste]';
    log(`${callWithAceLogPrefix} Entered ace_callWithAce callback for paste listener.`);

    if (!ace || !ace.editor) {
      console.error(`${callWithAceLogPrefix} ERROR: ace or ace.editor is not available. Cannot attach paste listener.`);
      log(`${callWithAceLogPrefix} Aborting paste listener attachment due to missing ace.editor.`);
      return;
    }
    const editor = ace.editor;
    log(`${callWithAceLogPrefix} ace.editor obtained successfully.`);

    // Store editor reference for later use in table operations
    log(`${logPrefix} Storing editor reference on editorInfo.`);
    ed.ep_tables5_editor = editor;
    log(`${logPrefix}: Stored editor reference as ed.ep_tables5_editor.`);

    // Attempt to find the inner iframe body, similar to ep_image_insert
    let $inner;
    try {
      log(`${callWithAceLogPrefix} Attempting to find inner iframe body for paste listener attachment.`);
      const $iframeOuter = $('iframe[name="ace_outer"]');
      if ($iframeOuter.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find outer iframe (ace_outer).`);
        log(`${callWithAceLogPrefix} Failed to find ace_outer.`);
        return;
      }
      log(`${callWithAceLogPrefix} Found ace_outer:`, $iframeOuter);

      const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
      if ($iframeInner.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find inner iframe (ace_inner).`);
        log(`${callWithAceLogPrefix} Failed to find ace_inner within ace_outer.`);
        return;
      }
      log(`${callWithAceLogPrefix} Found ace_inner:`, $iframeInner);

      const innerDocBody = $iframeInner.contents().find('body');
      if (innerDocBody.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find body element in inner iframe.`);
        log(`${callWithAceLogPrefix} Failed to find body in ace_inner.`);
        return;
      }
      $inner = $(innerDocBody[0]); // Ensure it's a jQuery object of the body itself
      log(`${callWithAceLogPrefix} Successfully found inner iframe body:`, $inner);
    } catch (e) {
      console.error(`${callWithAceLogPrefix} ERROR: Exception while trying to find inner iframe body:`, e);
      log(`${callWithAceLogPrefix} Exception details:`, { message: e.message, stack: e.stack });
      return;
    }

    if (!$inner || $inner.length === 0) {
      console.error(`${callWithAceLogPrefix} ERROR: $inner is not valid after attempting to find iframe body. Cannot attach paste listener.`);
      log(`${callWithAceLogPrefix} $inner is invalid. Aborting.`);
      return;
    }

    log(`${callWithAceLogPrefix} Attaching paste event listener to $inner (inner iframe body).`);
    $inner.on('paste', (evt) => {
      const pasteLogPrefix = '[ep_tables5:pasteHandler]';
      log(`${pasteLogPrefix} PASTE EVENT TRIGGERED. Event object:`, evt);

      log(`${pasteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        log(`${pasteLogPrefix} WARNING: Could not get representation or selection. Allowing default paste.`);
        console.warn(`${pasteLogPrefix} Could not get rep or selStart.`);
        return; // Allow default
      }
      log(`${pasteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      log(`${pasteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      // NEW: Check if selection spans multiple lines
      if (selStart[0] !== selEnd[0]) {
        log(`${pasteLogPrefix} WARNING: Selection spans multiple lines. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

      log(`${pasteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      const lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      if (!lineAttrString) {
        log(`${pasteLogPrefix} Line ${lineNum} is NOT a table line (no '${ATTR_TABLE_JSON}' attribute found). Allowing default paste.`);
        return; // Not a table line
      }
      log(`${pasteLogPrefix} Line ${lineNum} IS a table line. Attribute string: "${lineAttrString}".`);

      let tableMetadata;
      try {
        log(`${pasteLogPrefix} Parsing table metadata from attribute string.`);
        tableMetadata = JSON.parse(lineAttrString);
        log(`${pasteLogPrefix} Parsed table metadata:`, tableMetadata);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
          log(`${pasteLogPrefix} WARNING: Invalid or incomplete table metadata on line ${lineNum}. Allowing default paste. Metadata:`, tableMetadata);
          console.warn(`${pasteLogPrefix} Invalid table metadata for line ${lineNum}.`);
          return; // Allow default
        }
        log(`${pasteLogPrefix} Table metadata validated successfully: tblId=${tableMetadata.tblId}, row=${tableMetadata.row}, cols=${tableMetadata.cols}.`);
      } catch(e) {
        console.error(`${pasteLogPrefix} ERROR parsing table metadata for line ${lineNum}:`, e);
        log(`${pasteLogPrefix} Metadata parse error. Allowing default paste. Error details:`, { message: e.message, stack: e.stack });
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

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        log(`${pasteLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

      log(`${pasteLogPrefix} Accessing clipboard data.`);
      const clipboardData = evt.originalEvent.clipboardData || window.clipboardData;
      if (!clipboardData) {
        log(`${pasteLogPrefix} WARNING: No clipboard data found. Allowing default paste.`);
        return; // Allow default
      }
      log(`${pasteLogPrefix} Clipboard data object obtained:`, clipboardData);

      log(`${pasteLogPrefix} Getting 'text/plain' from clipboard.`);
      const pastedTextRaw = clipboardData.getData('text/plain');
      log(`${pasteLogPrefix} Pasted text raw: "${pastedTextRaw}" (Type: ${typeof pastedTextRaw})`);

      // ENHANCED: More thorough sanitization of pasted content
      let pastedText = pastedTextRaw
        .replace(/(\r\n|\n|\r)/gm, " ") // Replace newlines with space
        .replace(/\|/g, " ") // Replace pipe characters with space to prevent delimiter injection
        .replace(/\t/g, " ") // Replace tabs with space
        .replace(/\s+/g, " ") // Normalize whitespace
        .trim(); // Trim leading/trailing whitespace

      log(`${pasteLogPrefix} Pasted text after sanitization: "${pastedText}"`);

      if (typeof pastedText !== 'string' || pastedText.length === 0) {
        log(`${pasteLogPrefix} No plain text in clipboard or text is empty (after sanitization). Allowing default paste.`);
        const types = clipboardData.types;
        log(`${pasteLogPrefix} Clipboard types available:`, types);
        if (types && types.includes('text/html')) {
            log(`${pasteLogPrefix} Clipboard also contains HTML:`, clipboardData.getData('text/html'));
        }
        return; // Allow default if no plain text
      }
      log(`${pasteLogPrefix} Plain text obtained from clipboard: "${pastedText}". Length: ${pastedText.length}.`);

      // NEW: Check if paste would exceed cell boundaries
      const currentCellText = cells[targetCellIndex] || '';
      const selectionLength = selEnd[1] - selStart[1];
      const newCellLength = currentCellText.length - selectionLength + pastedText.length;
      
      // Optional: Add a reasonable maximum cell length if desired
      const MAX_CELL_LENGTH = 1000; // Example maximum
      if (newCellLength > MAX_CELL_LENGTH) {
        log(`${pasteLogPrefix} WARNING: Paste would exceed maximum cell length (${newCellLength} > ${MAX_CELL_LENGTH}). Truncating paste.`);
        const truncatedPaste = pastedText.substring(0, MAX_CELL_LENGTH - (currentCellText.length - selectionLength));
        if (truncatedPaste.length === 0) {
          log(`${pasteLogPrefix} Paste would be completely truncated. Preventing paste.`);
          evt.preventDefault();
          return;
        }
        log(`${pasteLogPrefix} Using truncated paste: "${truncatedPaste}"`);
        pastedText = truncatedPaste;
      }

      log(`${pasteLogPrefix} INTERCEPTING paste of plain text into table line ${lineNum}. PREVENTING DEFAULT browser action.`);
      evt.preventDefault();

      try {
        log(`${pasteLogPrefix} Preparing to perform paste operations via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
            const callAceLogPrefix = `${pasteLogPrefix}[ace_callWithAceOps]`;
            log(`${callAceLogPrefix} Entered ace_callWithAce for paste operations. selStart:`, selStart, `selEnd:`, selEnd);
            
            log(`${callAceLogPrefix} Original line text from initial rep: "${rep.lines.atIndex(lineNum).text}". SelStartCol: ${selStart[1]}, SelEndCol: ${selEnd[1]}.`);
            
            log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to insert text: "${pastedText}".`);
            aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, pastedText);
            log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

            log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
            const repAfterReplace = aceInstance.ace_getRep();
            log(`${callAceLogPrefix} Fetched rep after replace for applyMeta. Line ${lineNum} text now: "${repAfterReplace.lines.atIndex(lineNum).text}"`);
            
            ed.ep_tables5_applyMeta(
              lineNum,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );
            log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_tables5_applyMeta.`);

            const newCaretCol = selStart[1] + pastedText.length;
            const newCaretPos = [lineNum, newCaretCol];
            log(`${callAceLogPrefix} New calculated caret position: [${newCaretPos}]. Setting selection.`);
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            log(`${callAceLogPrefix} Selection change successful.`);
            
            log(`${callAceLogPrefix} Requesting fastIncorp(10) for sync.`);
            aceInstance.ace_fastIncorp(10);
            log(`${callAceLogPrefix} fastIncorp requested.`);

            // Update stored click/caret info
            if (editor && editor.ep_tables5_last_clicked && editor.ep_tables5_last_clicked.tblId === tableMetadata.tblId) {
               const newRelativePos = newCaretCol - cellStartCol;
               editor.ep_tables5_last_clicked = {
                  lineNum: lineNum,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
               };
               log(`${callAceLogPrefix} Updated stored click/caret info:`, editor.ep_tables5_last_clicked);
            }

            log(`${callAceLogPrefix} Paste operations within ace_callWithAce completed successfully.`);
        }, 'tablePasteTextOperations', true);
        log(`${pasteLogPrefix} ed.ace_callWithAce for paste operations was called.`);

      } catch (error) {
        console.error(`${pasteLogPrefix} CRITICAL ERROR during paste handling operation:`, error);
        log(`${pasteLogPrefix} Error details:`, { message: error.message, stack: error.stack });
        log(`${pasteLogPrefix} Paste handling FAILED. END OF HANDLER.`);
      }
    });
    log(`${callWithAceLogPrefix} Paste event listener attached.`);
  }, 'tables5_paste_listener', true);
  log(`${logPrefix} ace_callWithAce for paste listener setup completed.`);
  // *** END ADDED: Paste event listener ***

  // Helper function to apply the metadata attribute to a line
  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo, attributeString = null, documentAttributeManager = null) {
    const funcName = 'applyTableLineMetadataAttribute';
    log(`${logPrefix}:${funcName}: START - Applying METADATA attribute to line ${lineNum}`, {tblId, rowIndex, numCols});

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

       // Get existing attributes on the line to preserve them (especially block attributes)
       const docManager = documentAttributeManager || editorInfo?.documentAttributeManager;
       let existingAttributes = [];
       
       log(`${logPrefix}:${funcName}: DEBUG - docManager available: ${!!docManager}`);
       log(`${logPrefix}:${funcName}: DEBUG - getAttributesOnLine method: ${typeof docManager?.getAttributesOnLine}`);
       log(`${logPrefix}:${funcName}: DEBUG - getAttributeOnLine method: ${typeof docManager?.getAttributeOnLine}`);
       
       if (docManager && typeof docManager.getAttributesOnLine === 'function') {
         try {
           // Get all attributes currently on this line
           const lineAttribs = docManager.getAttributesOnLine(lineNum);
           log(`${logPrefix}:${funcName}: DEBUG - getAttributesOnLine result:`, lineAttribs);
           if (lineAttribs && typeof lineAttribs === 'object') {
             // Convert to the format expected by ace_performDocumentApplyAttributesToRange
             for (const [key, value] of Object.entries(lineAttribs)) {
               if (key !== ATTR_TABLE_JSON) { // Don't duplicate the tbljson attribute
                 existingAttributes.push([key, value]);
                 log(`${logPrefix}:${funcName}: Preserving existing attribute: ${key}=${value}`);
               }
             }
           }
         } catch (e) {
           log(`${logPrefix}:${funcName}: Warning - Could not get existing attributes via getAttributesOnLine, proceeding with fallback:`, e);
         }
       } 
       
       if (existingAttributes.length === 0 && docManager && typeof docManager.getAttributeOnLine === 'function') {
         // Fallback: try to get common block attributes individually
         log(`${logPrefix}:${funcName}: DEBUG - Using fallback individual attribute retrieval`);
         const commonBlockAttribs = ['align', 'center', 'heading', 'list', 'indent'];
         for (const attrKey of commonBlockAttribs) {
           try {
             const attrValue = docManager.getAttributeOnLine(lineNum, attrKey);
             if (attrValue) {
               existingAttributes.push([attrKey, attrValue]);
               log(`${logPrefix}:${funcName}: Preserving existing block attribute: ${attrKey}=${attrValue}`);
             }
           } catch (e) {
             // Ignore errors for individual attributes
             log(`${logPrefix}:${funcName}: DEBUG - Error getting individual attribute ${attrKey}:`, e);
           }
         }
       } else if (!docManager) {
         log(`${logPrefix}:${funcName}: Warning - No documentAttributeManager available, cannot preserve existing attributes.`);
       }

       // Combine existing attributes with the new tbljson attribute
       const allAttributes = [...existingAttributes, [ATTR_TABLE_JSON, attributeString]];
       log(`${logPrefix}:${funcName}: Applying all attributes:`, allAttributes);

       // Use ace_performDocumentApplyAttributesToRange
       // *** Apply all attributes to the FULL line range ***
       const start = [lineNum, 0];
       const end = [lineNum, effectiveLineLength]; // Use potentially corrected length
       log(`${logPrefix}:${funcName}: Applying attributes via ace_performDocumentApplyAttributesToRange to range [${start}]-[${end}]`);
       editorInfo.ace_performDocumentApplyAttributesToRange(start, end, allAttributes);
       log(`${logPrefix}:${funcName}: Applied all attributes to line ${lineNum} over range [${start}]-[${end}].`);
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
      // Note: documentAttributeManager not available in this context for new table creation
      applyTableLineMetadataAttribute(lineNumToApply, tblId, r, cols, currentRep, ed, null, null); 
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

  ed.ace_doDatatableOptions = (action) => {
    const funcName = 'ace_doDatatableOptions';
    log(`${funcName}: START - Processing action: ${action}`);
    
    // Get the last clicked cell info to determine which table to operate on
    const editor = ed.ep_tables5_editor;
    if (!editor) {
      console.error(`[ep_tables5] ${funcName}: Could not get editor reference.`);
      return;
    }
    
    const lastClick = editor.ep_tables5_last_clicked;
    if (!lastClick || !lastClick.tblId) {
      log(`${funcName}: No table selected. Please click on a table cell first.`);
      console.warn('[ep_tables5] No table selected. Please click on a table cell first.');
      return;
    }
    
    log(`${funcName}: Operating on table ${lastClick.tblId}, clicked line ${lastClick.lineNum}, cell ${lastClick.cellIndex}`);
    
    try {
      // Get current representation and document manager
      const currentRep = ed.ace_getRep();
      if (!currentRep || !currentRep.lines) {
        console.error(`[ep_tables5] ${funcName}: Could not get current representation.`);
        return;
      }
      
      // Use the stored documentAttributeManager reference
      const docManager = ed.ep_tables5_docManager;
      if (!docManager) {
        console.error(`[ep_tables5] ${funcName}: Could not get document attribute manager from stored reference.`);
        return;
      }
      
      log(`${funcName}: Successfully obtained documentAttributeManager from stored reference.`);
      
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
                    log(`${funcName}: Reconstructed metadata from DOM for line ${lineIndex}: ${lineAttrString}`);
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
        log(`${funcName}: No table lines found for table ${lastClick.tblId}`);
        return;
      }
      
      // Sort by row number to ensure correct order
      tableLines.sort((a, b) => a.row - b.row);
      log(`${funcName}: Found ${tableLines.length} table lines`);
      
      // Determine table dimensions and target indices with robust matching
      const numRows = tableLines.length;
      const numCols = tableLines[0].cols;
      
      // More robust way to find the target row - match by both line number AND row metadata
      let targetRowIndex = -1;
      
      // First try to match by line number
      targetRowIndex = tableLines.findIndex(line => line.lineIndex === lastClick.lineNum);
      
      // If that fails, try to match by finding the row that contains the clicked table
      if (targetRowIndex === -1) {
        log(`${funcName}: Direct line number match failed, searching by DOM structure...`);
        const clickedLineEntry = currentRep.lines.atIndex(lastClick.lineNum);
        if (clickedLineEntry && clickedLineEntry.lineNode) {
          const clickedTable = clickedLineEntry.lineNode.querySelector('table.dataTable[data-tblId="' + lastClick.tblId + '"]');
          if (clickedTable) {
            const clickedRowAttr = clickedTable.getAttribute('data-row');
            if (clickedRowAttr !== null) {
              const clickedRowNum = parseInt(clickedRowAttr, 10);
              targetRowIndex = tableLines.findIndex(line => line.row === clickedRowNum);
              log(`${funcName}: Found target row by DOM attribute matching: row ${clickedRowNum}, index ${targetRowIndex}`);
            }
          }
        }
      }
      
      // If still not found, default to first row but log the issue
      if (targetRowIndex === -1) {
        log(`${funcName}: Warning: Could not find target row, defaulting to row 0`);
        targetRowIndex = 0;
      }
      
      const targetColIndex = lastClick.cellIndex || 0;
      
      log(`${funcName}: Table dimensions: ${numRows} rows x ${numCols} cols. Target: row ${targetRowIndex}, col ${targetColIndex}`);
      
      // Perform table operations with both text and metadata updates
      let newNumCols = numCols;
      let success = false;
      
      switch (action) {
        case 'addTblRowA': // Insert row above
          log(`${funcName}: Inserting row above row ${targetRowIndex}`);
          success = addTableRowAboveWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;
          
        case 'addTblRowB': // Insert row below
          log(`${funcName}: Inserting row below row ${targetRowIndex}`);
          success = addTableRowBelowWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;
          
        case 'addTblColL': // Insert column left
          log(`${funcName}: Inserting column left of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnLeftWithText(tableLines, targetColIndex, ed, docManager);
          break;
          
        case 'addTblColR': // Insert column right
          log(`${funcName}: Inserting column right of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnRightWithText(tableLines, targetColIndex, ed, docManager);
          break;
          
        case 'delTblRow': // Delete row
          log(`${funcName}: Deleting row ${targetRowIndex}`);
          success = deleteTableRowWithText(tableLines, targetRowIndex, ed, docManager);
          break;
          
        case 'delTblCol': // Delete column
          log(`${funcName}: Deleting column ${targetColIndex}`);
          newNumCols = numCols - 1;
          success = deleteTableColumnWithText(tableLines, targetColIndex, ed, docManager);
          break;
          
        default:
          log(`${funcName}: Unknown action: ${action}`);
          return;
      }
      
      if (!success) {
        console.error(`[ep_tables5] ${funcName}: Table operation failed for action: ${action}`);
        return;
      }
      
      log(`${funcName}: Table operation completed successfully with text and metadata synchronization`);
      
    } catch (error) {
      console.error(`[ep_tables5] ${funcName}: Error during table operation:`, error);
      log(`${funcName}: Error details:`, { message: error.message, stack: error.stack });
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
      
      // Update metadata for all subsequent rows
      for (let i = targetRowIndex; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1; // +1 because we inserted a line
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex };
        
        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Apply metadata to the new row
      const newMetadata = { tblId, row: targetLine.metadata.row, cols: numCols };
      applyTableLineMetadataAttribute(insertLineIndex, tblId, targetLine.metadata.row, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_tables5] Error adding row above with text:', e);
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
      
      // Update metadata for all subsequent rows
      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1; // +1 because we inserted a line
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex };
        
        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      // Apply metadata to the new row
      const newMetadata = { tblId, row: targetLine.metadata.row + 1, cols: numCols };
      applyTableLineMetadataAttribute(insertLineIndex, tblId, targetLine.metadata.row + 1, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_tables5] Error adding row below with text:', e);
      return false;
    }
  }
  
  function addTableColumnLeftWithText(tableLines, targetColIndex, editorInfo, docManager) {
    try {
      // Update text content for all table lines using precise character insertion
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);
        
        // Calculate the exact insertion position
        let insertPos = 0;
        for (let i = 0; i < targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0) + DELIMITER.length;
        }
        
        // Insert empty cell with delimiter at the calculated position
        const textToInsert = (targetColIndex === 0) ? ' ' + DELIMITER : DELIMITER + ' ';
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];
        
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);
        
        // Update metadata
        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols + 1 };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols + 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_tables5] Error adding column left with text:', e);
      return false;
    }
  }
  
  function addTableColumnRightWithText(tableLines, targetColIndex, editorInfo, docManager) {
    try {
      // Update text content for all table lines using precise character insertion
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);
        
        // Calculate the exact insertion position (after the target column)
        let insertPos = 0;
        for (let i = 0; i <= targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0);
          if (i < cells.length - 1) insertPos += DELIMITER.length;
        }
        
        // Insert delimiter and empty cell at the calculated position
        const textToInsert = DELIMITER + ' ';
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];
        
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);
        
        // Update metadata
        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols + 1 };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols + 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_tables5] Error adding column right with text:', e);
      return false;
    }
  }
  
  function deleteTableRowWithText(tableLines, targetRowIndex, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      
      // Delete the entire line
      const deleteStart = [targetLine.lineIndex, 0];
      const deleteEnd = [targetLine.lineIndex + 1, 0];
      editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      
      // Update metadata for all subsequent rows
      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex - 1; // -1 because we deleted a line
        const newRowIndex = tableLines[i].metadata.row - 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex };
        
        applyTableLineMetadataAttribute(lineToUpdate, tableLines[i].metadata.tblId, newRowIndex, tableLines[i].cols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_tables5] Error deleting row with text:', e);
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
          log(`[ep_tables5] Warning: Target column ${targetColIndex} doesn't exist in line with ${cells.length} columns`);
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
        
        log(`[ep_tables5] Deleting column ${targetColIndex} from line ${tableLine.lineIndex}: chars ${deleteStart}-${deleteEnd} from "${lineText}"`);
        
        // Perform the precise deletion
        const rangeStart = [tableLine.lineIndex, deleteStart];
        const rangeEnd = [tableLine.lineIndex, deleteEnd];
        
        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, '');
        
        // Update metadata
        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols - 1 };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols - 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }
      
      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_tables5] Error deleting column with text:', e);
      return false;
    }
  }
  
  // ... existing code ...

  log('aceInitialized: END - helpers defined.');
};

// ───────────────────── required no‑op stubs ─────────────────────
exports.aceStartLineAndCharForPoint = () => { return undefined; };
exports.aceEndLineAndCharForPoint   = () => { return undefined; };

// NEW: Style protection for table cells
exports.aceSetAuthorStyle = (hook, ctx) => {
  const logPrefix = '[ep_tables5:aceSetAuthorStyle]';
  log(`${logPrefix} START`, { hook, ctx });

  // If no selection or no style to apply, allow default
  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd || !ctx.key) {
    log(`${logPrefix} No selection or style key. Allowing default.`);
    return;
  }

  // Check if selection is within a table
  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];
  
  // If selection spans multiple lines, prevent style application
  if (startLine !== endLine) {
    log(`${logPrefix} Selection spans multiple lines. Preventing style application to protect table structure.`);
    return false;
  }

  // Check if the line is a table line
  const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(startLine, ATTR_TABLE_JSON);
  if (!lineAttrString) {
    log(`${logPrefix} Line ${startLine} is not a table line. Allowing default style application.`);
    return;
  }

  // List of styles that could break table structure
  const BLOCKED_STYLES = [
    'list', 'listType', 'indent', 'align', 'heading', 'code', 'quote',
    'horizontalrule', 'pagebreak', 'linebreak', 'clear'
  ];

  if (BLOCKED_STYLES.includes(ctx.key)) {
    log(`${logPrefix} Blocked potentially harmful style '${ctx.key}' from being applied to table cell.`);
    return false;
  }

  // For allowed styles, ensure they only apply within cell boundaries
  try {
    const tableMetadata = JSON.parse(lineAttrString);
    if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
      log(`${logPrefix} Invalid table metadata. Preventing style application.`);
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
      log(`${logPrefix} Selection spans multiple cells. Preventing style application to protect table structure.`);
      return false;
    }

    // If selection includes cell delimiters, prevent style application
    const cellStartCol = cells.slice(0, selectionStartCell).reduce((acc, cell) => acc + cell.length + DELIMITER.length, 0);
    const cellEndCol = cellStartCol + cells[selectionStartCell].length;
    
    if (ctx.rep.selStart[1] <= cellStartCol || ctx.rep.selEnd[1] >= cellEndCol) {
      log(`${logPrefix} Selection includes cell delimiters. Preventing style application to protect table structure.`);
      return false;
    }

    log(`${logPrefix} Style '${ctx.key}' allowed within cell boundaries.`);
    return; // Allow the style to be applied
  } catch (e) {
    console.error(`${logPrefix} Error processing style application:`, e);
    log(`${logPrefix} Error details:`, { message: e.message, stack: e.stack });
    return false; // Prevent style application on error
  }
};

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

// NEW: Undo/Redo protection
exports.aceUndoRedo = (hook, ctx) => {
  const logPrefix = '[ep_tables5:aceUndoRedo]';
  log(`${logPrefix} START`, { hook, ctx });

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd) {
    log(`${logPrefix} No selection or context. Allowing default.`);
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
    log(`${logPrefix} No table lines affected. Allowing default undo/redo.`);
    return;
  }

  log(`${logPrefix} Table lines affected:`, { tableLines });

  // Validate table structure after undo/redo
  try {
    for (const line of tableLines) {
      const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
      if (!lineAttrString) continue;

      const tableMetadata = JSON.parse(lineAttrString);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
        log(`${logPrefix} Invalid table metadata after undo/redo. Attempting recovery.`);
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
          log(`${logPrefix} Recovered table structure for line ${line}`);
        } else {
          // If we can't recover, remove the table attribute
          ctx.documentAttributeManager.removeAttributeOnLine(line, ATTR_TABLE_JSON);
          log(`${logPrefix} Removed invalid table attribute from line ${line}`);
        }
      }
    }
  } catch (e) {
    console.error(`${logPrefix} Error during undo/redo validation:`, e);
    log(`${logPrefix} Error details:`, { message: e.message, stack: e.stack });
  }
};

// END OF FILE