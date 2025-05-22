const fsp = require('fs').promises;
const path = require('path');
const { JSDOM } = require('jsdom'); // JSDOM for server-side DOM manipulation

// Using console.log for server-side logging, similar to other Etherpad server-side files.
const log = (...m) => console.log('[ep_tables5:tableImport]', ...m);

// Utilities copied/adapted from client_hooks.js for server-side use
// const DELIMITER = '|'; // No longer used for joining cell content in this approach
const rand = () => Math.random().toString(36).slice(2, 8);

// Server-side base64 encoding, URL-safe variant.
// Etherpad's client-side dec() can handle base64 with or without padding,
// and replaces '-' and '_' back to '+' and '/' before decoding.
const enc = (s) => Buffer.from(s, 'utf8').toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_');
    // Padding '=' is kept, as standard atob on client-side handles it.

// The delimiter used to join cell HTML content within the single span
const HTML_CELL_DELIMITER = '|';

exports.import = async (hookName, context) => {
  // ImportError class is passed by Etherpad core
  const { srcFile, fileEnding, destFile, padId, ImportError } = context;
  const logPrefix = '[ep_tables5:importHookRun]'; // Differentiate from module-level log
  log(`${logPrefix} START - fileEnding: ${fileEnding}, srcFile: ${srcFile}, destFile: ${destFile}`);

  if (fileEnding !== '.html' && fileEnding !== '.htm') {
    log(`${logPrefix} Not an HTML file ('${fileEnding}'). This will be handled by ep_docx_html_customizer.`);
    return false; // Let Etherpad core or other plugins handle it
  }

  try {
    let htmlContent = await fsp.readFile(srcFile, 'utf8');
    const dom = new JSDOM(htmlContent);
    const document = dom.window.document;
    const tables = Array.from(document.querySelectorAll('table')); // Ensure it's an array
    let modified = false;

    log(`${logPrefix} Found ${tables.length} table(s) in HTML file: ${srcFile}`);

    if (tables.length === 0) {
        log(`${logPrefix} No tables found. Ensuring destFile exists if different from srcFile.`);
        if (srcFile !== destFile) {
            try {
                await fsp.copyFile(srcFile, destFile);
                log(`${logPrefix} Copied original ${srcFile} to ${destFile}.`);
            } catch (copyErr) {
                console.error(`${logPrefix} Error copying ${srcFile} to ${destFile} (no tables found):`, copyErr);
                // This is a problem for the import process, throw an error.
                throw new ImportError('htmlProcessingFailed', `Failed to copy file: ${copyErr.message}`);
            }
        }
        return false; // No tables to process for this plugin
    }

    for (const tableNode of tables) {
      const tblId = rand(); // Generate a unique ID for this table
      const sourceRows = Array.from(tableNode.querySelectorAll('tr'));
      const elementsToInsertInsteadOfTable = []; // Store new DIVs/P to replace the TABLE

      if (sourceRows.length === 0) {
        log(`${logPrefix} Table (tblId: ${tblId}) has no rows. Creating a placeholder paragraph.`);
        const placeholder = document.createElement('p');
        // User-friendly message that will appear in the pad if an empty table is imported
        placeholder.textContent = `[Empty table (ID: ${tblId}) was removed during import as it had no rows]`;
        elementsToInsertInsteadOfTable.push(placeholder);
      } else {
         log(`${logPrefix} Processing table (tblId: ${tblId}), ${sourceRows.length} rows.`);
      }

      for (let r = 0; r < sourceRows.length; r++) {
        const sourceRowNode = sourceRows[r];
        const sourceCells = Array.from(sourceRowNode.querySelectorAll('td, th'));
        const numCols = sourceCells.length;

        // Skip rows that are completely empty and have no cells,
        // especially if they are the only row in a table (often formatting artifacts).
        if (numCols === 0 && (sourceRowNode.textContent || '').trim() === '') {
            log(`${logPrefix} Skipping empty row (index ${r}) with no cells in table (tblId: ${tblId}).`);
            // If this was the only row and it's skipped, ensure the table placeholder reflects this
            if (sourceRows.length === 1 && elementsToInsertInsteadOfTable.length === 0) {
                const placeholder = document.createElement('p');
                placeholder.textContent = `[Table (ID: ${tblId}) was removed during import as its only row was empty and had no cells]`;
                elementsToInsertInsteadOfTable.push(placeholder);
            }
            continue; // Skip this empty row
        }

        // Create a plain DIV element to act as the line container for this row
        const lineWrapperDiv = document.createElement('div');

        const rowMetadata = { tblId, row: r, cols: numCols };
        const stringifiedMetadata = JSON.stringify(rowMetadata);
        const encodedJsonMetadata = enc(stringifiedMetadata);

        const cellHTMLContents = sourceCells.map(cellNode => {
          return (cellNode.innerHTML || '&nbsp;').replace(/\|/g, '&vert;');
        });
        const delimitedRowHTML = cellHTMLContents.join(HTML_CELL_DELIMITER);

        const lineContentSpan = document.createElement('span');
        lineContentSpan.className = `tbljson-${encodedJsonMetadata}`;
        lineContentSpan.innerHTML = delimitedRowHTML;
        
        lineWrapperDiv.appendChild(lineContentSpan);
        elementsToInsertInsteadOfTable.push(lineWrapperDiv);
        log(`${logPrefix} Created div > span.tbljson-${encodedJsonMetadata.substring(0,10)}... for tblId ${tblId}, row ${r}. Delimited HTML: ${(delimitedRowHTML || '').substring(0,70)}...`);
      }

      // Replace the original <table> with the series of new <p> elements
      if (tableNode.parentNode) {
        if (elementsToInsertInsteadOfTable.length > 0) {
            const fragment = document.createDocumentFragment();
            elementsToInsertInsteadOfTable.forEach(el => fragment.appendChild(el));
            tableNode.parentNode.replaceChild(fragment, tableNode);
            modified = true;
            log(`${logPrefix} Replaced source table with ${elementsToInsertInsteadOfTable.length} div > span element(s).`);
        } else if (sourceRows.length > 0) {
            // This case: rows existed, but all were skipped (e.g., all were empty with no cells)
            const placeholder = document.createElement('p');
            placeholder.textContent = `[Table (ID: ${tblId}) was removed during import as all its rows were empty or had no content]`;
            tableNode.parentNode.replaceChild(placeholder, tableNode);
            modified = true;
            log(`${logPrefix} Replaced table (was to be ${tblId}) with a placeholder as all its rows were skipped.`);
        }
        // If rows.length was 0, elementsToInsertInsteadOfTable should already contain a placeholder.
      } else {
          log(`${logPrefix} Table node (was to be ${tblId}) has no parentNode. Cannot replace. This may occur if the table is the root of the document.`);
          // If the table is the entire body content, or similar, replacement is complex.
          // For now, this means such a table might not be processed if it's the sole root element.
          // `modified` would remain false for this specific table if it can't be replaced.
      }
    }

    if (modified) {
      const newHtmlContent = dom.serialize();
      await fsp.writeFile(srcFile, newHtmlContent, 'utf8');
      log(`${logPrefix} HTML file ${srcFile} was modified with table transformations and saved.`);
    } else {
      log(`${logPrefix} No valid tables were transformed or no modifications were made to HTML file ${srcFile} by table import logic.`);
    }

    // Etherpad's core import logic expects the final HTML in `destFile`.
    // Copy srcFile to destFile to ensure the (potentially modified) content is processed.
    if (srcFile !== destFile) {
      await fsp.copyFile(srcFile, destFile);
      log(`${logPrefix} Copied ${modified ? 'modified ' : 'original '}${srcFile} to ${destFile}.`);
    } else {
        log(`${logPrefix} srcFile and destFile are the same ('${srcFile}'). No separate copy needed beyond in-place modifications (if any).`);
    }
    
    // Return true if we made any modifications, indicating we handled the file.
    // If not modified, return false so Etherpad core might try other handlers or default import.
    return modified;

  } catch (error) {
    console.error(`${logPrefix} CRITICAL ERROR processing HTML file ${srcFile}:`, error);
    log(`${logPrefix} Error Name: ${error.name}, Error Message: ${error.message}, Error Stack: ${error.stack ? error.stack : 'N/A'}`);
    
    // Attempt to ensure destFile is a copy of the original srcFile for Etherpad's default import to try
    if (srcFile !== destFile) {
        try {
            // context.srcFile refers to the original, unmodified srcFile path string from the hook arguments
            await fsp.copyFile(context.srcFile, destFile); 
            log(`${logPrefix} Copied original ${context.srcFile} to ${destFile} during error recovery.`);
        } catch (copyErr) {
            console.error(`${logPrefix} Error copying original ${context.srcFile} to ${destFile} during error recovery:`, copyErr);
        }
    }
    
    if (error instanceof ImportError) { // If it's already an ImportError, rethrow it
        throw error;
    }
    // Otherwise, wrap it in an ImportError so Etherpad can handle it gracefully
    throw new ImportError('tableImportFailed', `Table plugin error: ${error.message || 'Unknown error during table import'}`);
  }
}; 