var DatatablesRendererExport = require('ep_tables5/static/js/datatables-renderer.js');

const log = (...m) => console.debug('[ep_tables5:export2html]', ...m);

exports.getLineHTMLForExport = function (hook, context) {
  log('getLineHTMLForExport: START', { hook, contextText: context?.text, lineNum: context?.lineNumber });
  // The old code checked for "data-tables", but the new system uses attributes.
  // We need a reliable way to check if this line is a table line based on attributes.
  // Checking for the attribute key in the attribLine might be fragile.
  // A better approach might be to check context.lineAttribs if available, or rely on a marker class if one is consistently applied during export attribute processing.
  // For now, let's assume the presence of 'tblProp' or 'tbljson' in the attribute pool might indicate a table line.

  // Attempt to find a table attribute index
  var attribIndex = retrieveIndex(context.attribLine, context.apool);
  log('getLineHTMLForExport: Retrieved attribIndex:', attribIndex);

  if (attribIndex) {
    try {
      // Attempt to get the attribute value, checking for both old ('tblProp') and new ('tbljson') keys
      let dtAttrs = null;
      const attribData = context.apool.numToAttrib[attribIndex];
      if (attribData) {
         log('getLineHTMLForExport: Found attribute data for index:', attribIndex, 'Data:', attribData);
         if (attribData[0] === 'tblProp' || attribData[0] === 'tbljson') {
            dtAttrs = attribData[1];
            log('getLineHTMLForExport: Found table attribute value (tblProp/tbljson):', dtAttrs);
         } else {
             log('getLineHTMLForExport: Attribute at index is not tblProp or tbljson:', attribData[0]);
         }
      } else {
          log('getLineHTMLForExport: No attribute data found in apool for index:', attribIndex);
      }

      // Only render if we found valid attributes
      if (dtAttrs) {
        log('getLineHTMLForExport: Calling DatatablesRendererExport.render with context "export".');
        // The renderer expects the *raw* line text containing the JSON payload for export context.
        // We pass the `el` equivalent which is context.text here.
        const renderInputEl = { text: context.text }; // Simulate the element structure expected by the renderer in export context
        context.lineContent = DatatablesRendererExport.DatatablesRenderer.render("export", renderInputEl, dtAttrs);
        log('getLineHTMLForExport: Renderer returned HTML:', context.lineContent);
      } else {
          log('getLineHTMLForExport: Skipping render, no valid table attributes found for index.');
      }
    } catch (error) {
        log('getLineHTMLForExport: ERROR during rendering:', error);
        console.error("[ep_tables5:export2html] Error processing line for export:", error, "Context:", context);
        // Return true to allow default processing if rendering fails
        return true;
    }
  } else {
    log('getLineHTMLForExport: Not a table line (no relevant attribute index found).');
  }
  log('getLineHTMLForExport: END');
  // Return true allows other plugins or default handling if we didn't modify lineContent
  return true;
};

// Updated retrieveIndex to be more robust and log potential issues
retrieveIndex = function (attribLine, apool) {
  log('retrieveIndex: START', { attribLine });
  if (!attribLine || typeof attribLine !== 'string') {
    log('retrieveIndex: Invalid attribLine input.');
    return null;
  }

  // Regex to find attribute application like *N or *N+L where N is base36 index
  const match = attribLine.match(/\*([0-9a-z]+)/);
  if (!match || !match[1]) {
      log('retrieveIndex: No attribute marker found in attribLine.');
      return null;
  }

  const base36Index = match[1];
  log('retrieveIndex: Found base36 index:', base36Index);

  try {
      // Convert base36 index to number
      const numIndex = parseInt(base36Index, 36);
      log('retrieveIndex: Converted to numeric index:', numIndex);

      // Validate if the index exists in the attribute pool and corresponds to a table attribute
      if (apool && apool.numToAttrib && apool.numToAttrib[numIndex]) {
          const attribData = apool.numToAttrib[numIndex];
          log('retrieveIndex: Attribute data in pool:', attribData);
          if (attribData[0] === 'tblProp' || attribData[0] === 'tbljson') {
              log('retrieveIndex: Found valid table attribute at index. Returning numeric index.');
              return numIndex; // Return the numeric index if it points to a known table attribute
          } else {
              log('retrieveIndex: Attribute at index is not tblProp or tbljson:', attribData[0]);
              return null; // Attribute exists but isn't the one we want
          }
      } else {
          log('retrieveIndex: Index not found in apool or apool invalid.');
          // Optionally, still return the index if apool check is not desired/reliable
          // return numIndex;
          return null;
      }
  } catch (error) {
      log('retrieveIndex: ERROR converting base36 index:', error);
      console.error("[ep_tables5:export2html] Error converting base36 index:", base36Index, error);
      return null;
  }
}
