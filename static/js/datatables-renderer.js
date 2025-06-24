/* ep_data_tables – datatables-renderer.js
 *
 * Only used by:
 *   • export pipeline           (`context === "export"`)
 *   • timeslider frame          (`context === "timeslider"`)
 * Regular pad rendering is done by client_hooks.js.
 */

const log = (...m) => console.debug('[ep_data_tables:datatables-renderer]', ...m);
const DELIMITER = '|'; // Used for splitting cell content if ever needed by legacy, export now receives pre-split, pre-rendered cells

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
  