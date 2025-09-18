'use strict';

const {
  ATTR_TABLE_JSON,
  DELIMITER,
} = require('./shared');

module.exports = (_hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceUndoRedo]';

  if (!ctx || !ctx.rep || !ctx.rep.selStart || !ctx.rep.selEnd) {
    return;
  }

  const startLine = ctx.rep.selStart[0];
  const endLine = ctx.rep.selEnd[0];

  let hasTableLines = false;
  const tableLines = [];

  for (let line = startLine; line <= endLine; line++) {
    const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
    if (lineAttrString) {
      hasTableLines = true;
      tableLines.push(line);
    }
  }

  if (!hasTableLines) {
    return;
  }

  try {
    for (const line of tableLines) {
      const lineAttrString = ctx.documentAttributeManager?.getAttributeOnLine(line, ATTR_TABLE_JSON);
      if (!lineAttrString) continue;

      const tableMetadata = JSON.parse(lineAttrString);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
        const lineText = ctx.rep.lines.atIndex(line)?.text || '';
        const cells = lineText.split(DELIMITER);

        if (cells.length > 1) {
          const newMetadata = {
            cols: cells.length,
            rows: 1,
            cells: cells.map((_, i) => ({col: i, row: 0})),
          };

          ctx.documentAttributeManager.setAttributeOnLine(
            line,
            ATTR_TABLE_JSON,
            JSON.stringify(newMetadata)
          );
        } else {
          ctx.documentAttributeManager.removeAttributeOnLine(line, ATTR_TABLE_JSON);
        }
      }
    }
  } catch (e) {
    console.error(`${logPrefix} Error during undo/redo validation:`, e);
  }
};
