'use strict';

const {
  ATTR_TABLE_JSON,
  DELIMITER,
} = require('./shared');

module.exports = (_hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceSetAuthorStyle]';

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd || !ctx.key) {
    return;
  }

  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];

  if (startLine !== endLine) {
    return false;
  }

  const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(startLine, ATTR_TABLE_JSON);
  if (!lineAttrString) {
    return;
  }

  const BLOCKED_STYLES = [
    'list', 'listType', 'indent', 'align', 'heading', 'code', 'quote',
    'horizontalrule', 'pagebreak', 'linebreak', 'clear'
  ];

  if (BLOCKED_STYLES.includes(ctx.key)) {
    return false;
  }

  try {
    const tableMetadata = JSON.parse(lineAttrString);
    if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
      return false;
    }

    const lineText = ctx.rep.lines.atIndex(startLine)?.text || '';
    const cells = lineText.split(DELIMITER);
    let currentOffset = 0;
    let selectionStartCell = -1;
    let selectionEndCell = -1;

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

    if (selectionStartCell !== selectionEndCell) {
      return false;
    }

    const cellStartCol = cells
      .slice(0, selectionStartCell)
      .reduce((acc, cell) => acc + cell.length + DELIMITER.length, 0);
    const cellEndCol = cellStartCol + (cells[selectionStartCell]?.length ?? 0);

    if (ctx.rep.selStart[1] <= cellStartCol || ctx.rep.selEnd[1] >= cellEndCol) {
      return false;
    }

    return;
  } catch (e) {
    console.error(`${logPrefix} Error processing style application:`, e);
    return false;
  }
};
