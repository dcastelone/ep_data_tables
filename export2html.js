var DatatablesRendererExport = require('ep_tables5/static/js/datatables-renderer.js');
const Changeset = require('ep_etherpad-lite/static/js/Changeset'); // Required for subattribution
const AttributeMap = require('ep_etherpad-lite/static/js/AttributeMap');

const log = (...m) => console.debug('[ep_tables5:export2html]', ...m);
const DELIMITER = '|'; // Delimiter used in the plaintext representation of the line

// Helper function to escape HTML, simplified (moved here from renderer)
function escapeHtml(text = '') {
  const strText = String(text);
  var map = {
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  };
  return strText.replace(/[&<>"'']/g, function(m) { return map[m]; });
}

// Helper to convert a single cell's plain text and its specific attribute string to HTML
const renderCellRichTextServer = (textSegment, attribsForSegment, apool) => {
  if (textSegment === undefined || textSegment === null) textSegment = ''; // Ensure textSegment is a string
  let html = '';
  let currentPos = 0;
  for (const op of Changeset.deserializeOps(attribsForSegment)) {
    const opChars = op.chars < 0 ? 0 : op.chars; // Ensure op.chars is not negative
    const opText = textSegment.substring(currentPos, currentPos + opChars);
    currentPos += opChars;
    if (op.attribs) {
      const A = AttributeMap.fromString(op.attribs, apool);
      let tags = [];
      if (A.get('bold')) tags.push(['<b>', '</b>']);
      if (A.get('italic')) tags.push(['<i>', '</i>']);
      if (A.get('underline')) tags.push(['<u>', '</u>']);
      if (A.get('strikethrough')) tags.push(['<s>', '</s>']);
      // Add other attribute to HTML tag conversions here if needed
      html += tags.reduce((acc, curr) => curr[0] + acc + curr[1], escapeHtml(opText));
    } else {
      html += escapeHtml(opText);
    }
  }
  return html;
};

exports.getLineHTMLForExport = function (hook, context) {
  log('getLineHTMLForExport: START', { 
    hookName: hook, 
    lineText: context?.text?.substring(0, 50) + (context?.text?.length > 50 ? '...' : ''), 
    attribLine: context?.attribLine, 
    lineNumber: context?.lineNumber 
  });

  var attribIndex = retrieveIndex(context.attribLine, context.apool);
  // log('getLineHTMLForExport: Retrieved attribIndex:', attribIndex);

  if (attribIndex !== null) { // Ensure attribIndex is not null before proceeding
    let tblJsonMetadataString = null;
    const attribData = context.apool.numToAttrib[attribIndex];
    if (attribData && (attribData[0] === 'tblProp' || attribData[0] === 'tbljson')) {
      tblJsonMetadataString = attribData[1];
      // log('getLineHTMLForExport: Found table attribute value (tblProp/tbljson):', tblJsonMetadataString);
    } else {
      // log('getLineHTMLForExport: Attribute at index is not tblProp or tbljson or attribData missing.');
    }

    if (tblJsonMetadataString) {
      let rowMetadata;
      try {
        rowMetadata = JSON.parse(tblJsonMetadataString);
      } catch (e) {
        log('getLineHTMLForExport: ERROR - Failed to parse tblJsonMetadataString:', tblJsonMetadataString, e);
        return true; // Default handling for corrupted metadata
      }

      log('getLineHTMLForExport: Parsed rowMetadata:', rowMetadata);

      const plainTextCells = (context.text || '').split(DELIMITER);
      const richHtmlCells = [];
      let currentTextOffset = 0;

      if (plainTextCells.length !== rowMetadata.cols) {
        log(`getLineHTMLForExport: WARN - Plain text cell count (${plainTextCells.length}) ` +
            `and metadata cols (${rowMetadata.cols}) mismatch. Line text: "${context.text}"`);
      }

      // Iterate based on the number of columns defined in metadata for robustness
      for (let i = 0; i < rowMetadata.cols; i++) {
        const cellText = plainTextCells[i] || ''; // Use empty string if segment is missing
        // Calculate sub-attribute string for this cell
        const cellAttribLine = Changeset.subattribution(context.attribLine, currentTextOffset, cellText.length);
        const cellRichHtml = renderCellRichTextServer(cellText, cellAttribLine, context.apool);
        richHtmlCells.push(cellRichHtml);
        currentTextOffset += cellText.length;
        if (i < rowMetadata.cols - 1) { // Only add delimiter length if not the last cell based on metadata
          currentTextOffset++; // Account for the delimiter character that was in context.text
        }
      }
      
      log('getLineHTMLForExport: Generated rich HTML for cells:', richHtmlCells.map(h => (h||'').substring(0,30)+'...'));

      const renderInput = { cellsRichHtml: richHtmlCells }; // Pass the array of rich HTML cell strings
      let tableHtml = DatatablesRendererExport.DatatablesRenderer.render("export", renderInput, tblJsonMetadataString);
      
      if (tableHtml && typeof tableHtml === 'string') {
        context.lineContent = tableHtml.trim();
        log('getLineHTMLForExport: Renderer returned HTML (trimmed): ', context.lineContent.substring(0,100) + '...');
        return false; 
      } else {
        log('getLineHTMLForExport: Renderer did not return a valid HTML string. Metadata:', tblJsonMetadataString);
      }
    } else {
      log('getLineHTMLForExport: Skipping render, no valid table attributes string found for index.', attribIndex);
    }
  } else {
    // log('getLineHTMLForExport: Not a table line (no relevant attribute index found).');
  }
  // log('getLineHTMLForExport: END');
  return true; // Default handling if not processed as a table line
};

// retrieveIndex function remains the same as the last correct version
retrieveIndex = function (attribLine, apool) {
  log('retrieveIndex: START', { attribLine });
  if (!attribLine || typeof attribLine !== 'string') {
    log('retrieveIndex: Invalid attribLine input.');
    return null;
  }

  const matches = Array.from(attribLine.matchAll(/\*([0-9a-z]+)/g));
  if (!matches || matches.length === 0) {
      log('retrieveIndex: No attribute markers found in attribLine.');
      return null;
  }

  log('retrieveIndex: Found potential attribute markers:', matches.map(m => m[1]));

  for (const match of matches) {
    const base36Index = match[1];
    try {
        const numIndex = parseInt(base36Index, 36);
        log(`retrieveIndex: Checking base36 index: ${base36Index} (numeric: ${numIndex})`);

        if (apool && apool.numToAttrib && apool.numToAttrib[numIndex]) {
            const attribData = apool.numToAttrib[numIndex];
            log(`retrieveIndex: Attribute data in pool for index ${numIndex}:`, attribData);
            if (attribData[0] === 'tblProp' || attribData[0] === 'tbljson') {
                log('retrieveIndex: Found valid table attribute (tblProp or tbljson) at index. Returning numeric index:', numIndex);
                return numIndex; 
            }
        } else {
            log(`retrieveIndex: Index ${numIndex} (base36: ${base36Index}) not found in apool or apool invalid.`);
        }
    } catch (error) {
        log(`retrieveIndex: ERROR converting base36 index: ${base36Index}`, error);
    }
  }
  log('retrieveIndex: No tblProp or tbljson attribute found after checking all markers.');
  return null; 
}
