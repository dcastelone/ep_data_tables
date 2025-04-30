/* ep_tables5 – datatables-renderer.js
 *
 * Only used by:
 *   • export pipeline           (`context === "export"`)
 *   • timeslider frame          (`context === "timeslider"`)
 * Regular pad rendering is done by client_hooks.js.
 */

const log = (...m) => console.debug('[ep_tables5:datatables-renderer]', ...m);

if (typeof DatatablesRenderer === 'undefined') {
    log('Defining DatatablesRenderer');
    var DatatablesRenderer = (() => {
      const buildHtml = (rowJSON, props = {}) => {
        log('buildHtml: START', { rowJSON, props });
        const p       = { borderWidth: 1, width: 100, ...rowJSON.tblProperties, ...props };
        log('buildHtml: Combined properties:', p);
        const border  = `border:${p.borderWidth}px solid ${p.borderColor || '#ccc'}`;
        const tdStyle = `${border};padding:4px;word-wrap:break-word`;
        log('buildHtml: TD Style:', tdStyle);
        const tr      = rowJSON.payload[0]
          .map(txt => `<td style="${tdStyle}">${txt}</td>`)
          .join('');
        const html = `<table class="dataTable" style="border-collapse:collapse;width:${p.width}%"><tbody><tr>${tr}</tr></tbody></table>`;
        log('buildHtml: Generated HTML:', html);
        return html;
      };
  
      return {
        /* Main entry -------------------------------------------------------- */
        render(ctx, el, attribs) {
          log('render: START', { ctx, el, attribs });
          // Editor context now handled elsewhere – bail early
          if (!ctx || (ctx !== 'timeslider' && ctx !== 'export')) {
            log('render: Skipping - Invalid context:', ctx);
            return;
          }
          log('render: Context is valid:', ctx);
  
          let row;
          if (ctx.parsedJson) {
            log('render: Using provided ctx.parsedJson');
            row = ctx.parsedJson;
          } else {
            log('render: Parsing JSON from element text/html');
            const txt = ctx === 'timeslider'
              ? $('<div>').html(el.innerHTML).text()      // unescape & strip <span>
              : (ctx === 'export' ? el.text : el.innerHTML);
            log('render: Text to parse:', txt);
  
            try {
              row = JSON.parse(txt);
              log('render: Parsed JSON successfully:', row);
            } catch (e) {
              log('render: ERROR - Failed to parse JSON from text:', e);
              console.error('[ep_tables5:datatables-renderer] Failed to parse JSON:', txt, e);
              return; // Bail if JSON parsing fails
            }
          }
  
          const extraProps = attribs ? JSON.parse(attribs) : {};
          log('render: Parsed extraProps:', extraProps);
          const html = buildHtml(row, extraProps);
  
          if (ctx === 'export') {
            log('render: Returning HTML string for export');
            return html; // exporter expects string
          }
          if (el.innerHTML !== html) {
            log('render: Updating element innerHTML for timeslider');
            el.innerHTML = html; // timeslider
          } else {
            log('render: Skipping innerHTML update, content matches (timeslider)');
          }
          log('render: END');
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
  