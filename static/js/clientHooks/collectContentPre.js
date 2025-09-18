'use strict';

const {
  ATTR_TABLE_JSON,
  DELIMITER,
  decodeTbljsonClass,
  extractEncodedTbljson,
} = require('./shared');

module.exports = (hook, ctx) => {
  const funcName = 'collectContentPre';
  const node = ctx.domNode;
  const state = ctx.state;
  const cc = ctx.cc;

  if (node?.classList?.contains('ace-line')) {
    const tableNode = node.querySelector('table.dataTable[data-tblId]');
    if (tableNode) {
      const docManager = cc.documentAttributeManager;
      const rep = cc.rep;
      const lineNum = rep?.lines?.indexOfKey(node.id);

      if (typeof lineNum === 'number' && lineNum >= 0 && docManager) {
        try {
          const existingAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);

          if (existingAttrString) {
            const existingMetadata = JSON.parse(existingAttrString);
            if (existingMetadata && typeof existingMetadata.tblId !== 'undefined' &&
                typeof existingMetadata.row !== 'undefined' && typeof existingMetadata.cols === 'number') {
              const trNode = tableNode.querySelector('tbody > tr');
              if (trNode) {
                let cellHTMLSegments = Array.from(trNode.children).map((td, index) => {
                  let segmentHTML = td.innerHTML || '';

                  const resizeHandleRegex = /<div class="ep-data_tables-resize-handle"[^>]*><\/div>/ig;
                  segmentHTML = segmentHTML.replace(resizeHandleRegex, '');
                  const hiddenDelimRegexPrimary = /<span class="ep-data_tables-delim"[^>]*>.*?<\/span>/ig;
                  segmentHTML = segmentHTML.replace(hiddenDelimRegexPrimary, '');
                  const caretAnchorRegex = /<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig;
                  segmentHTML = segmentHTML.replace(caretAnchorRegex, '');
                  const textCheck = segmentHTML.replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
                  if (textCheck === '') segmentHTML = '';

                  return segmentHTML;
                });

                if (cellHTMLSegments.length !== existingMetadata.cols) {
                  while (cellHTMLSegments.length < existingMetadata.cols) cellHTMLSegments.push('');
                  if (cellHTMLSegments.length > existingMetadata.cols) cellHTMLSegments.length = existingMetadata.cols;
                }

                const canonicalLineText = cellHTMLSegments.join(DELIMITER);
                state.line = canonicalLineText;

                state.lineAttributes = state.lineAttributes || [];
                state.lineAttributes = state.lineAttributes.filter((attr) => attr[0] !== ATTR_TABLE_JSON);
                state.lineAttributes.push([ATTR_TABLE_JSON, existingAttrString]);

                return undefined;
              }
            }
          } else {
            const domTblId = tableNode.getAttribute('data-tblId');
            const domRow = tableNode.getAttribute('data-row');
            const trNode = tableNode.querySelector('tbody > tr');
            if (domTblId && domRow !== null && trNode && trNode.children.length > 0) {
              const domCols = trNode.children.length;
              const tempMetadata = {tblId: domTblId, row: parseInt(domRow, 10), cols: domCols};
              const tempAttrString = JSON.stringify(tempMetadata);

              let cellHTMLSegments = Array.from(trNode.children).map((td, index) => {
                let segmentHTML = td.innerHTML || '';
                const resizeHandleRegex = /<div class="ep-data_tables-resize-handle"[^>]*><\/div>/ig;
                segmentHTML = segmentHTML.replace(resizeHandleRegex, '');
                if (index > 0) {
                  const hiddenDelimRegex = new RegExp('^<span class="ep-data_tables-delim" contenteditable="false">' + DELIMITER + '(<\\/span>)?<\\/span>', 'i');
                  segmentHTML = segmentHTML.replace(hiddenDelimRegex, '');
                }
                const caretAnchorRegex = /<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig;
                segmentHTML = segmentHTML.replace(caretAnchorRegex, '');
                const textCheck = segmentHTML.replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
                if (textCheck === '') segmentHTML = '';
                return segmentHTML;
              });

              if (cellHTMLSegments.length !== domCols) {
                while (cellHTMLSegments.length < domCols) cellHTMLSegments.push('');
                if (cellHTMLSegments.length > domCols) cellHTMLSegments.length = domCols;
              }

              const canonicalLineText = cellHTMLSegments.join(DELIMITER);
              state.line = canonicalLineText;
              state.lineAttributes = state.lineAttributes || [];
              state.lineAttributes = state.lineAttributes.filter((attr) => attr[0] !== ATTR_TABLE_JSON);
              state.lineAttributes.push([ATTR_TABLE_JSON, tempAttrString]);
              return undefined;
            }
          }
        } catch (e) {
          console.error(`[ep_data_tables] ${funcName}: Line ${lineNum} error during DOM reconstruction:`, e);
        }
      }
    }
  }

  const encodedFromCls = extractEncodedTbljson(ctx.cls || '');
  if (encodedFromCls) {
    const info = decodeTbljsonClass(ctx.cls || '');
    if (info) {
      if (!info.isWellFormed) {
        if (info.error) {
          console.warn(`[ep_data_tables] ${funcName}: Decoded tbljson class could not be parsed as JSON.`, info.error);
        } else {
          console.warn(`[ep_data_tables] ${funcName}: Decoded tbljson metadata is missing required fields.`, info.metadata);
        }
      }
      cc.doAttrib(state, `${ATTR_TABLE_JSON}::${info.json}`);
    } else {
      console.warn(`[ep_data_tables] ${funcName}: Failed to decode tbljson metadata from classes on ${node?.tagName}.`);
    }
  }
};
