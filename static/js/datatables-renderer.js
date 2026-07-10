/* ep_data_tables – datatables-renderer.js
 *
 * Only used by:
 *   • export pipeline           (`context === "export"`)
 *   • timeslider frame          (`context === "timeslider"`)
 * Regular pad rendering is done by client_hooks.js.
 */

const log = (...m) => console.debug('[ep_data_tables:datatables-renderer]', ...m);
const DELIMITER = '\u241F';
const HIDDEN_DELIM = DELIMITER;

const enhanceTableHtml = (html, metadata = {}) => {
  if (typeof document === 'undefined') return html;
  const template = document.createElement('template');
  template.innerHTML = html;
  const table = template.content.querySelector('table.dataTable');
  if (!table || typeof table.querySelectorAll !== 'function') return html;
  table.setAttribute('data-ep-data-tables-accessible', 'true');
  table.setAttribute('aria-label', `Table row ${(Number(metadata.row) || 0) + 1}`);
  Array.from(table.querySelectorAll('td, th')).forEach((cell, index) => {
    cell.setAttribute('aria-colindex', String(index + 1));
  });
  for (const delimiter of table.querySelectorAll('.ep-data_tables-delim, .ep-data_tables-caret-anchor')) {
    delimiter.setAttribute('aria-hidden', 'true');
  }
  for (const handle of table.querySelectorAll('.ep-data_tables-resize-handle')) {
    handle.setAttribute('aria-hidden', 'true');
    handle.setAttribute('tabindex', '-1');
  }
  return template.innerHTML;
};

const enc = (s) => btoa(s).replace(/\+/g, '-').replace(/\//g, '_');
const dec = (s) => {
  const str = s.replace(/-/g, '+').replace(/_/g, '/');
  try {
    if (typeof atob === 'function') return atob(str);
    if (typeof Buffer === 'function') return Buffer.from(str, 'base64').toString('utf8');
  } catch (e) {
    console.error('[ep_data_tables:datatables-renderer] Error decoding base64 string:', s, e);
  }
  return null;
};

// buildExportHtml now expects an array of pre-rendered HTML strings for each cell.
const buildExportHtml = (metadata, richHtmlCellArray) => {
  const funcName = 'buildExportHtml';
  log(`${funcName}: START`, { metadata, numberOfCells: richHtmlCellArray.length });

  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined' || typeof metadata.cols !== 'number') {
    log(`${funcName}: ERROR - Invalid or missing metadata. Metadata:`, metadata);
    return `<div>Error: Missing table metadata for export.</div>`;
  }

  if (metadata.cols !== richHtmlCellArray.length) {
    log(`${funcName}: WARN - Column count in metadata (${metadata.cols}) does not match provided HTML cell array length (${richHtmlCellArray.length}).`);
    // Adjusting to render only the minimum of available cells or metadata.cols to prevent errors.
    // Or, one might choose to pad with empty cells if segments are fewer.
  }

  // Get column widths from metadata, or use equal distribution if not set
  const numCols = richHtmlCellArray.length;
  const columnWidths = metadata.columnWidths || Array(numCols).fill(100 / numCols);
  
  // Ensure we have the right number of column widths
  while (columnWidths.length < numCols) {
    columnWidths.push(100 / numCols);
  }
  if (columnWidths.length > numCols) {
    columnWidths.splice(numCols);
  }

  const tdStyle = `padding: 5px 7px; word-wrap:break-word; vertical-align: top; border: 1px solid #000;`;
  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';

  // Each item in richHtmlCellArray is now the complete inner HTML for a cell.
  // Apply column widths to each cell
  const cellsHtml = richHtmlCellArray.map((cellHtml, index) => {
    const widthPercent = columnWidths[index] || (100 / numCols);
    const cellStyle = `${tdStyle} width: ${widthPercent}%;`;
    return `<td style="${cellStyle}">${cellHtml || '&nbsp;'}</td>`;
  }).join('');

  const tableHtml = 
    `<table class="dataTable${firstRowClass}" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse:collapse; table-layout: fixed;">` +
    `<tbody><tr>${cellsHtml}</tr></tbody>` +
    `</table>`;
  
  log(`${funcName}: END - Generated HTML for export: ${tableHtml.substring(0,150)}...`);
  return tableHtml;
};

const findTbljsonClass = (element) => {
  if (!element || !element.classList) return null;
  for (const cls of element.classList) {
    if (cls.startsWith('tbljson-')) return cls.substring(8);
  }
  return null;
};

const findTbljsonEncodedMetadata = (element) => {
  if (!element || typeof element.querySelectorAll !== 'function') return null;
  const own = findTbljsonClass(element);
  if (own) return own;
  for (const candidate of element.querySelectorAll('[class*="tbljson-"]')) {
    const encoded = findTbljsonClass(candidate);
    if (encoded) return encoded;
  }
  return null;
};

const buildTimesliderHtml = (metadata, innerHTMLSegments) => {
  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined') {
    return '<table class="dataTable dataTable-error"><tbody><tr><td>Error: Missing table metadata</td></tr></tbody></table>';
  }

  const numCols = innerHTMLSegments.length;
  const columnWidths = metadata.columnWidths || Array(numCols).fill(100 / numCols);
  while (columnWidths.length < numCols) columnWidths.push(100 / numCols);
  if (columnWidths.length > numCols) columnWidths.splice(numCols);

  let encodedTbljsonClass = '';
  try {
    encodedTbljsonClass = `tbljson-${enc(JSON.stringify(metadata))}`;
  } catch (_) {}

  const tdBaseStyle = 'padding: 5px 7px; word-wrap:break-word; vertical-align: top; border: 1px solid #000; position: relative;';
  const cellsHtml = innerHTMLSegments.map((segment, index) => {
    const textOnly = (segment || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
    const containsImage = /\bimage-placeholder\b/.test(segment || '');
    const isEmpty = (!segment || textOnly === '') && !containsImage;
    let modifiedSegment = segment || '';

    if (isEmpty) {
      const cellClass = encodedTbljsonClass ? `${encodedTbljsonClass} tblCell-${index}` : `tblCell-${index}`;
      modifiedSegment = `<span class="${cellClass}">&nbsp;</span>`;
    }

    if (index > 0) {
      const delimSpan = `<span class="ep-data_tables-delim" contenteditable="false">${HIDDEN_DELIM}</span>`;
      modifiedSegment = modifiedSegment.replace(/^(<span[^>]*>)/i, `$1${delimSpan}`);
      if (!/^<span[^>]*>/i.test(modifiedSegment)) modifiedSegment = `${delimSpan}${modifiedSegment}`;
    }

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
        classList = classList.filter((c) => !/^tblCell-\d+$/.test(c));
        classList.push(requiredCellClass);
        const newClassAttr = ` class="${Array.from(new Set(classList)).join(' ')}"`;
        const attrsWithoutClass = classMatch ? attrs.replace(/\s*class\s*=\s*"[^"]*"/i, '') : attrs;
        const cleanedTail = tail.slice(fullOpen.length).replace(/(<span[^>]*class=")([^"]*)(")/ig, (m, p1, classes, p3) => {
          const filtered = classes.split(/\s+/).filter((c) => c && !/^tblCell-\d+$/.test(c)).join(' ');
          return p1 + filtered + p3;
        });
        modifiedSegment = `${head}<span${newClassAttr}${attrsWithoutClass}>${cleanedTail}`;
      }
    } catch (_) {}

    const widthPercent = columnWidths[index] || (100 / numCols);
    return `<td style="${tdBaseStyle} width: ${widthPercent}%;" data-column="${index}" draggable="false" autocorrect="off" autocapitalize="off" spellcheck="false">${modifiedSegment}</td>`;
  }).join('');

  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';
  const tableHtml = `<table class="dataTable${firstRowClass}" writingsuggestions="false" autocorrect="off" autocapitalize="off" spellcheck="false" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse: collapse; table-layout: fixed;" draggable="false"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
  return enhanceTableHtml(tableHtml, metadata);
};

const renderTimesliderLine = (line) => {
  if (!line || typeof line.querySelectorAll !== 'function') return false;
  if (line.querySelector('table.dataTable[data-tblId], table.dataTable[data-tblid]')) return false;

  const encodedJsonString = findTbljsonEncodedMetadata(line);
  if (!encodedJsonString) return false;

  let metadata;
  try {
    const decoded = dec(encodedJsonString);
    if (!decoded) return false;
    metadata = JSON.parse(decoded);
  } catch (e) {
    console.error('[ep_data_tables:datatables-renderer] Failed to parse tbljson metadata:', e);
    return false;
  }

  const sanitizedHTMLForSplit = (line.innerHTML || '')
    .replace(/<span class="ep-data_tables-delim"[^>]*>[\s\S]*?<\/span>/ig, DELIMITER)
    .replace(/<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig, '')
    .replace(/\r?\n/g, ' ')
    .replace(/<br\s*\/?>/gi, ' ');
  const htmlSegments = sanitizedHTMLForSplit.split(DELIMITER);

  line.innerHTML = buildTimesliderHtml(metadata, htmlSegments);
  line.classList.add('ep-data_tables-timeslider-rendered');
  return true;
};

const renderTimesliderTables = (root = document) => {
  const body = root.querySelector ? (root.querySelector('#innerdocbody') || root) : root;
  if (!body || typeof body.querySelectorAll !== 'function') return 0;
  let rendered = 0;
  for (const line of body.querySelectorAll('.ace-line')) {
    if (renderTimesliderLine(line)) rendered++;
  }
  if (rendered) log(`Rendered ${rendered} timeslider table row(s).`);
  return rendered;
};

const getVisibleTimesliderRevision = () => {
  const label = document.querySelector('#revision_label')?.textContent || '';
  const match = label.match(/\bVersion\s+(\d+)\b/i);
  return match ? Number(match[1]) : null;
};

const getRequestedTimesliderRevision = () => {
  const match = String(window.location.hash || '').match(/^#(\d+)$/);
  return match ? Number(match[1]) : null;
};

const isTimesliderReadyToRender = () => {
  const body = document.querySelector('#innerdocbody') || document.body;
  if (!body || !body.querySelector('.ace-line')) return false;

  const requestedRevision = getRequestedTimesliderRevision();
  if (requestedRevision == null) return true;

  const visibleRevision = getVisibleTimesliderRevision();
  return visibleRevision === requestedRevision;
};

const startTimesliderRenderer = () => {
  if (typeof document === 'undefined' || !document.body || !document.body.classList?.contains('timeslider')) return;
  const target = document.querySelector('#innerdocbody') || document.body;
  if (!target || target.__epDataTablesTimesliderObserver) return;

  let renderTimer = null;
  const schedule = () => {
    clearTimeout(renderTimer);
    renderTimer = setTimeout(() => {
      renderTimer = null;
      if (!isTimesliderReadyToRender()) {
        schedule();
        return;
      }
      renderTimesliderTables(document);
    }, 250);
  };

  const scheduleAfterSliderUpdate = () => {
    setTimeout(() => {
      schedule();
    }, 0);
  };

  if (window.BroadcastSlider && typeof window.BroadcastSlider.onSlider === 'function') {
    window.BroadcastSlider.onSlider(scheduleAfterSliderUpdate);
  }
  window.addEventListener('hashchange', scheduleAfterSliderUpdate);

  const observer = new MutationObserver(schedule);
  observer.observe(target, {childList: true, subtree: true});
  target.__epDataTablesTimesliderObserver = observer;

  schedule();
};

if (typeof window !== 'undefined') {
  window.epDataTablesRenderTimeslider = renderTimesliderTables;
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startTimesliderRenderer, {once: true});
  } else {
    startTimesliderRenderer();
  }
}

if (typeof DatatablesRenderer === 'undefined') {
    log('Defining DatatablesRenderer');
    var DatatablesRenderer = (() => {
      const buildHtmlLegacy = (rowJSON, props = {}) => {
        log('buildHtmlLegacy: START', { rowJSON, props });
        const p = { borderWidth: 1, width: 100, ...rowJSON.tblProperties, ...props };
        const border  = `border:${p.borderWidth}px solid ${p.borderColor || '#ccc'}`;
        const tdStyle = `${border};padding:4px;word-wrap:break-word`;
        const tr = (rowJSON.payload && rowJSON.payload[0] && Array.isArray(rowJSON.payload[0])) 
          ? rowJSON.payload[0].map(txt => `<td style="${tdStyle}">${txt}</td>`).join('')
          : '<td>Error: Invalid payload structure for buildHtmlLegacy</td>';
        const html = `<table class="dataTable" style="border-collapse:collapse;width:${p.width}%"><tbody><tr>${tr}</tr></tbody></table>`;
        log('buildHtmlLegacy: Generated HTML:', html);
        return html;
      };
  
      return {
        render(ctx, el, metadataJsonString) { 
          log('render: START', { ctx, elContext: el, metadataJson: metadataJsonString });

          if (ctx === 'export') {
            log('render (export context): Processing for export.');
            // el is now expected to be an object like { cellsRichHtml: ["<td>Cell1HTML</td>", ...] }
            if (!el || !el.cellsRichHtml || !Array.isArray(el.cellsRichHtml)) {
              log('render (export context): ERROR - el.cellsRichHtml is missing or not an array.', el);
              return '<div>Error: Missing pre-rendered cell HTML for table export.</div>';
            }
            if (!metadataJsonString) {
              log('render (export context): ERROR - metadataJsonString is undefined.');
              return `<div>Error: Missing table metadata for export.</div>`;
            }

            let rowMetadata;
            try {
              rowMetadata = JSON.parse(metadataJsonString);
              log('render (export context): Parsed rowMetadata:', rowMetadata);
            } catch (e) {
              log('render (export context): ERROR - Failed to parse metadata JSON:', metadataJsonString, e);
              return `<div>Error: Invalid table metadata.</div>`;
            }
            
            const exportHtml = buildExportHtml(rowMetadata, el.cellsRichHtml);
            log('render (export context): Returning HTML string for export.');
            return exportHtml;
          }
          
          // Fallback to existing logic for other contexts (e.g., timeslider)
          log(`render (context: ${ctx}): Using legacy/timeslider rendering path.`);
          if (!ctx || (ctx !== 'timeslider')) { 
            log('render: Skipping - Invalid context for legacy path:', ctx);
            return;
          }
          log('render: Context is valid for legacy path:', ctx);
  
          let row;
          if (ctx.parsedJson) {
            log('render: Using provided ctx.parsedJson for legacy path');
            row = ctx.parsedJson;
          } else {
            log('render: Parsing JSON from element text/html for legacy path');
            const txt = ctx === 'timeslider'
              ? (typeof $ !== 'undefined' ? $('<div>').html(el.innerHTML).text() : el.innerHTML) 
              : el.innerHTML; 
            log('render: Text to parse for legacy path:', txt);
  
            try {
              row = JSON.parse(txt);
              log('render: Parsed JSON successfully for legacy path:', row);
            } catch (e) {
              log('render: ERROR - Failed to parse JSON from text for legacy path:', e);
              console.error('[ep_data_tables:datatables-renderer] Failed to parse JSON (legacy path):', txt, e);
              return; 
            }
          }
          
          const extraPropsLegacy = metadataJsonString && typeof metadataJsonString === 'string' ? JSON.parse(metadataJsonString) : (metadataJsonString || {});
          log('render: Parsed extraProps for legacy path:', extraPropsLegacy);
          const html = buildHtmlLegacy(row, extraPropsLegacy);
  
          if (el.innerHTML !== html) {
            log('render: Updating element innerHTML for legacy path (timeslider?)');
            el.innerHTML = html;
          } else {
            log('render: Skipping innerHTML update, content matches for legacy path (timeslider?)');
          }
          log('render: END (legacy path)');
        },
      };
    })();
} else {
  log('DatatablesRenderer already defined.');
}
  
if (typeof exports !== 'undefined') {
  log('Exporting DatatablesRenderer for Node.js');
  exports.DatatablesRenderer = DatatablesRenderer;
}
  
