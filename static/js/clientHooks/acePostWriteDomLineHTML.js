'use strict';

const {
  ATTR_TABLE_JSON,
  DELIMITER,
  buildTableFromDelimitedHTML,
  decodeTbljsonClass,
  findTbljsonElement,
} = require('./shared');

module.exports = function acePostWriteDomLineHTML(_hookName, args, cb) {
  const funcName = 'acePostWriteDomLineHTML';
  const node = args?.node;
  const nodeId = node?.id;
  const lineNum = args?.lineNumber;

  if (!node || !nodeId) {
    console.error(`[ep_data_tables] ${funcName}: Received invalid node or node without ID.`);
    return cb();
  }

  if (node.children) {
    for (let i = 0; i < Math.min(node.children.length, 10); i++) {
      const child = node.children[i];
      if (child.className && child.className.includes('tbljson-')) {
        // diagnostic log placeholder
      }
    }
  }

  let rowMetadata = null;
  let metadataInfo = decodeTbljsonClass(node.classList || node.className || '');
  if (!metadataInfo) {
    const tbljsonElement = findTbljsonElement(node);
    if (tbljsonElement && tbljsonElement !== node) {
      metadataInfo = decodeTbljsonClass(tbljsonElement.classList || tbljsonElement.className || '');
    }
  }

  if (!metadataInfo) {
    const existingTable = node.querySelector('table.dataTable[data-tblId]');
    if (existingTable) {
      const existingTblId = existingTable.getAttribute('data-tblId');
      const existingRow = existingTable.getAttribute('data-row');

      if (existingTblId && existingRow !== null) {
        const _tableCells = existingTable.querySelectorAll('td');
        if (lineNum !== undefined && args?.documentAttributeManager) {
          try {
            args.documentAttributeManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
          } catch (e) {
            console.warn('[ep_data_tables] Error getting line attribute for orphaned table:', e);
          }
        }
      }
    }

    return cb();
  }

  const existingTable = node.querySelector('table.dataTable[data-tblId]');
  if (existingTable) {
    return cb();
  }

  if (!metadataInfo.isWellFormed || !metadataInfo.metadata) {
    const errorDetails = metadataInfo.error || metadataInfo.metadata;
    console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Failed to decode/parse/validate tbljson.`, errorDetails);
    node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Invalid table metadata attribute found.</div>';
    return cb();
  }

  rowMetadata = metadataInfo.metadata;

  const delimitedTextFromLine = node.innerHTML;

  const _delimiterCount = (delimitedTextFromLine || '').split(DELIMITER).length - 1;
  let pos = -1;
  const _delimiterPositions = [];
  while ((pos = delimitedTextFromLine.indexOf(DELIMITER, pos + 1)) !== -1) {
    _delimiterPositions.push(pos);
  }

  const spanDelimRegex = /<span class="ep-data_tables-delim"[^>]*>[\s\S]*?<\/span>/ig;
  const sanitizedHTMLForSplit = (delimitedTextFromLine || '')
    .replace(spanDelimRegex, DELIMITER)
    .replace(/<span class="ep-data_tables-caret-anchor"[^>]*><\/span>/ig, '')
    .replace(/\r?\n/g, ' ')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/\u00A0/gu, ' ')
    .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ');
  const htmlSegments = sanitizedHTMLForSplit.split(DELIMITER);

  for (let i = 0; i < htmlSegments.length; i++) {
    const segment = htmlSegments[i] || '';
    try {
      const tblCellMatches = segment.match(/\btblCell-(\d+)\b/g) || [];
      const uniqueCells = Array.from(new Set(tblCellMatches));
      if (uniqueCells.length > 1) {
        console.warn('[ep_data_tables][diag] segment contains multiple tblCell-* markers', { segIndex: i, uniqueCells });
      }
    } catch (_) {}
  }

  let finalHtmlSegments = htmlSegments;

  if (htmlSegments.length !== rowMetadata.cols) {
    console.warn('[ep_data_tables][diag] Segment/column mismatch', { nodeId, lineNum, segs: htmlSegments.length, cols: rowMetadata.cols, tblId: rowMetadata.tblId, row: rowMetadata.row });

    const _hasImageSelected = delimitedTextFromLine.includes('currently-selected');
    const _hasImageContent = delimitedTextFromLine.includes('image:');
    let usedClassReconstruction = false;

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

    if (finalHtmlSegments.length !== rowMetadata.cols) {
      console.warn(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Could not reconstruct to expected ${rowMetadata.cols} segments. Got ${finalHtmlSegments.length}.`);
    }
  }

  try {
    const newTableHTML = buildTableFromDelimitedHTML(rowMetadata, finalHtmlSegments);

    const tbljsonElement = findTbljsonElement(node);

    if (tbljsonElement && tbljsonElement.parentElement && tbljsonElement.parentElement !== node) {
      const parentTag = tbljsonElement.parentElement.tagName.toLowerCase();
      const blockElements = ['center', 'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'right', 'left', 'ul', 'ol', 'li', 'code'];

      if (blockElements.includes(parentTag)) {
        tbljsonElement.parentElement.innerHTML = newTableHTML;
      } else {
        node.innerHTML = newTableHTML;
      }
    } else {
      node.innerHTML = newTableHTML;
    }
  } catch (renderError) {
    console.error(`[ep_data_tables] ${funcName} NodeID#${nodeId}: Error building/rendering table.`, renderError);
    node.innerHTML = '<div style="color:red; border: 1px solid red; padding: 5px;">[ep_data_tables] Error: Failed to render table structure.</div>';
    return cb();
  }

  return cb();
};
