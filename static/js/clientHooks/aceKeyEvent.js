'use strict';

const {
  ATTR_TABLE_JSON,
  DELIMITER,
  navigateToCellBelow,
  navigateToNextCell,
} = require('./shared');

module.exports = (_hook, ctx) => {
  const funcName = 'aceKeyEvent';
  const evt = ctx.evt;
  const rep = ctx.rep;
  const editorInfo = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

  const startLogTime = Date.now();
  const logPrefix = '[ep_data_tables:aceKeyEvent]';
 // log(`${logPrefix} START Key='${evt?.key}' Code=${evt?.keyCode} Type=${evt?.type} Modifiers={ctrl:${evt?.ctrlKey},alt:${evt?.altKey},meta:${evt?.metaKey},shift:${evt?.shiftKey}}`, { selStart: rep?.selStart, selEnd: rep?.selEnd });

  if (!rep || !rep.selStart || !editorInfo || !evt || !docManager) {
   // log(`${logPrefix} Skipping - Missing critical context.`);
    return false;
  }

  const reportedLineNum = rep.selStart[0];
  const reportedCol = rep.selStart[1]; 
 // log(`${logPrefix} Reported caret from rep: Line=${reportedLineNum}, Col=${reportedCol}`);

  let tableMetadata = null;
  let lineAttrString = null;
  try {
   // log(`${logPrefix} DEBUG: Attempting to get ${ATTR_TABLE_JSON} attribute from line ${reportedLineNum}`);
    lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
   // log(`${logPrefix} DEBUG: getAttributeOnLine returned: ${lineAttrString ? `"${lineAttrString}"` : 'null/undefined'}`);

    if (typeof docManager.getAttributesOnLine === 'function') {
      try {
        const allAttribs = docManager.getAttributesOnLine(reportedLineNum);
       // log(`${logPrefix} DEBUG: All attributes on line ${reportedLineNum}:`, allAttribs);
      } catch(e) {
       // log(`${logPrefix} DEBUG: Error getting all attributes:`, e);
      }
    }

    if (!lineAttrString) {
      try {
        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(reportedLineNum);
        if (lineEntry && lineEntry.lineNode) {
          const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
          if (tableInDOM) {
            const domTblId = tableInDOM.getAttribute('data-tblId');
            const domRow = tableInDOM.getAttribute('data-row');
           // log(`${logPrefix} DEBUG: Found table in DOM without attribute! TblId=${domTblId}, Row=${domRow}`);
            const domCells = tableInDOM.querySelectorAll('td');
            if (domTblId && domRow !== null && domCells.length > 0) {
             // log(`${logPrefix} DEBUG: Attempting to reconstruct metadata from DOM...`);
              const reconstructedMetadata = {
                tblId: domTblId,
                row: parseInt(domRow, 10),
                cols: domCells.length
              };
              lineAttrString = JSON.stringify(reconstructedMetadata);
             // log(`${logPrefix} DEBUG: Reconstructed metadata: ${lineAttrString}`);
            }
          }
        }
      } catch(e) {
       // log(`${logPrefix} DEBUG: Error checking DOM for table:`, e);
      }
    }

    if (lineAttrString) {
        tableMetadata = JSON.parse(lineAttrString);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
            // log(`${logPrefix} Line ${reportedLineNum} has attribute, but metadata invalid/missing cols.`);
             tableMetadata = null;
        }
    } else {
       // log(`${logPrefix} DEBUG: No ${ATTR_TABLE_JSON} attribute found on line ${reportedLineNum}`);
    }
  } catch(e) {
    console.error(`${logPrefix} Error checking/parsing line attribute for line ${reportedLineNum}.`, e);
    tableMetadata = null;
  }

  const editor = editorInfo.editor;
  const lastClick = editor?.ep_data_tables_last_clicked;
 // log(`${logPrefix} Reading stored click/caret info:`, lastClick);

  let currentLineNum = -1;
  let targetCellIndex = -1;
  let relativeCaretPos = -1;
  let precedingCellsOffset = 0; 
  let cellStartCol = 0; 
  let lineText = '';
  let cellTexts = [];
  let metadataForTargetLine = null;
  let trustedLastClick = false;

  if (lastClick) {
     // log(`${logPrefix} Attempting to validate stored click info for Line=${lastClick.lineNum}...`);
      let storedLineAttrString = null;
      let storedLineMetadata = null;
      try {
         // log(`${logPrefix} DEBUG: Getting ${ATTR_TABLE_JSON} attribute from stored line ${lastClick.lineNum}`);
          storedLineAttrString = docManager.getAttributeOnLine(lastClick.lineNum, ATTR_TABLE_JSON);
         // log(`${logPrefix} DEBUG: Stored line attribute result: ${storedLineAttrString ? `"${storedLineAttrString}"` : 'null/undefined'}`);

          if (storedLineAttrString) {
            storedLineMetadata = JSON.parse(storedLineAttrString);
           // log(`${logPrefix} DEBUG: Parsed stored metadata:`, storedLineMetadata);
          }

          if (storedLineMetadata && typeof storedLineMetadata.cols === 'number' && storedLineMetadata.tblId === lastClick.tblId) {
             // log(`${logPrefix} Stored click info VALIDATED (Metadata OK and tblId matches). Trusting stored state.`);
              trustedLastClick = true;
              currentLineNum = lastClick.lineNum; 
              targetCellIndex = lastClick.cellIndex;
              metadataForTargetLine = storedLineMetadata; 
              lineAttrString = storedLineAttrString;

              lineText = rep.lines.atIndex(currentLineNum)?.text || '';
              cellTexts = lineText.split(DELIMITER);
             // log(`${logPrefix} Using Line=${currentLineNum}, CellIndex=${targetCellIndex}. Text: "${lineText}"`);

              if (cellTexts.length !== metadataForTargetLine.cols) {
                 // log(`${logPrefix} WARNING: Stored cell count mismatch for trusted line ${currentLineNum}.`);
              }

              cellStartCol = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                  cellStartCol += (cellTexts[i]?.length ?? 0) + DELIMITER.length;
              }
              precedingCellsOffset = cellStartCol;
             // log(`${logPrefix} Calculated cellStartCol=${cellStartCol} from trusted cellIndex=${targetCellIndex}.`);

              if (typeof lastClick.relativePos === 'number' && lastClick.relativePos >= 0) {
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(lastClick.relativePos, currentCellTextLength));
                 // log(`${logPrefix} Using and validated stored relative position: ${relativeCaretPos}.`);
  } else {
                  relativeCaretPos = reportedCol - cellStartCol;
                  const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
                  relativeCaretPos = Math.max(0, Math.min(relativeCaretPos, currentCellTextLength)); 
                 // log(`${logPrefix} Stored relativePos missing, calculated from reportedCol (${reportedCol}): ${relativeCaretPos}`);
              }
          } else {
             // log(`${logPrefix} Stored click info INVALID (Metadata missing/invalid or tblId mismatch). Clearing stored state.`);
              if (editor) editor.ep_data_tables_last_clicked = null;
          }
      } catch (e) {
           console.error(`${logPrefix} Error validating stored click info for line ${lastClick.lineNum}.`, e);
           if (editor) editor.ep_data_tables_last_clicked = null;
      }
  }

  if (!trustedLastClick) {
     // log(`${logPrefix} Fallback: Using reported caret position Line=${reportedLineNum}, Col=${reportedCol}.`);
      try {
          lineAttrString = docManager.getAttributeOnLine(reportedLineNum, ATTR_TABLE_JSON);
          if (lineAttrString) tableMetadata = JSON.parse(lineAttrString);
          if (!tableMetadata || typeof tableMetadata.cols !== 'number') tableMetadata = null;

          if (!lineAttrString) {
            try {
              const rep = editorInfo.ace_getRep();
              const lineEntry = rep.lines.atIndex(reportedLineNum);
              if (lineEntry && lineEntry.lineNode) {
                const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
                if (tableInDOM) {
                  const domTblId = tableInDOM.getAttribute('data-tblId');
                  const domRow = tableInDOM.getAttribute('data-row');
                 // log(`${logPrefix} Fallback: Found table in DOM without attribute! TblId=${domTblId}, Row=${domRow}`);
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domTblId && domRow !== null && domCells.length > 0) {
                   // log(`${logPrefix} Fallback: Attempting to reconstruct metadata from DOM...`);
                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length
                    };
                    lineAttrString = JSON.stringify(reconstructedMetadata);
                    tableMetadata = reconstructedMetadata;
                   // log(`${logPrefix} Fallback: Reconstructed metadata: ${lineAttrString}`);
                  }
                }
              }
            } catch(e) {
             // log(`${logPrefix} Fallback: Error checking DOM for table:`, e);
            }
          }
      } catch(e) { tableMetadata = null; }

      if (!tableMetadata) {
         // log(`${logPrefix} Fallback: Reported line ${reportedLineNum} is not a valid table line. Allowing default.`);
           return false;
      }

      currentLineNum = reportedLineNum;
      metadataForTargetLine = tableMetadata;
     // log(`${logPrefix} Fallback: Processing based on reported line ${currentLineNum}.`);

      lineText = rep.lines.atIndex(currentLineNum)?.text || '';
      cellTexts = lineText.split(DELIMITER);
     // log(`${logPrefix} Fallback: Fetched text for reported line ${currentLineNum}: "${lineText}"`);

      if (cellTexts.length !== metadataForTargetLine.cols) {
         // log(`${logPrefix} WARNING (Fallback): Cell count mismatch for reported line ${currentLineNum}.`);
      }

      let currentOffset = 0;
      let foundIndex = -1;
      for (let i = 0; i < cellTexts.length; i++) {
          const cellLength = cellTexts[i]?.length ?? 0;
          const cellEndCol = currentOffset + cellLength;
          if (reportedCol >= currentOffset && reportedCol <= cellEndCol) {
              foundIndex = i;
              relativeCaretPos = reportedCol - currentOffset;
              cellStartCol = currentOffset;
              precedingCellsOffset = cellStartCol;
             // log(`${logPrefix} --> (Fallback Calc) Found target cell ${foundIndex}. RelativePos: ${relativeCaretPos}.`);
              break; 
          }
          if (i < cellTexts.length - 1 && reportedCol === cellEndCol + DELIMITER.length) {
              foundIndex = i + 1;
              relativeCaretPos = 0; 
              cellStartCol = currentOffset + cellLength + DELIMITER.length;
              precedingCellsOffset = cellStartCol;
             // log(`${logPrefix} --> (Fallback Calc) Caret at delimiter AFTER cell ${i}. Treating as start of cell ${foundIndex}.`);
              break;
          }
          currentOffset += cellLength + DELIMITER.length;
      }

      if (foundIndex === -1) {
          if (reportedCol === lineText.length && cellTexts.length > 0) {
                foundIndex = cellTexts.length - 1;
                cellStartCol = 0; 
                for (let i = 0; i < foundIndex; i++) { cellStartCol += (cellTexts[i]?.length ?? 0) + DELIMITER.length; }
                precedingCellsOffset = cellStartCol;
                relativeCaretPos = cellTexts[foundIndex]?.length ?? 0; 
               // log(`${logPrefix} --> (Fallback Calc) Caret detected at END of last cell (${foundIndex}).`);
          } else {
           // log(`${logPrefix} (Fallback Calc) FAILED to determine target cell for caret col ${reportedCol}. Allowing default handling.`);
            return false; 
          }
      }
      targetCellIndex = foundIndex;
  }

  if (currentLineNum < 0 || targetCellIndex < 0 || !metadataForTargetLine || targetCellIndex >= metadataForTargetLine.cols) {
      // log(`${logPrefix} FAILED final validation: Line=${currentLineNum}, Cell=${targetCellIndex}, Metadata=${!!metadataForTargetLine}. Allowing default.`);
    if (editor) editor.ep_data_tables_last_clicked = null;
    return false;
  }

 // log(`${logPrefix} --> Final Target: Line=${currentLineNum}, CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}`);

  const selStartActual = rep.selStart;
  const selEndActual = rep.selEnd;
  const hasSelection = selStartActual[0] !== selEndActual[0] || selStartActual[1] !== selEndActual[1];

  if (hasSelection) {
   // log(`${logPrefix} [selection] Active selection detected. Start:[${selStartActual[0]},${selStartActual[1]}], End:[${selEndActual[0]},${selEndActual[1]}]`);
   // log(`${logPrefix} [caretTrace] [selection] Initial rep.selStart: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

    if (selStartActual[0] !== currentLineNum || selEndActual[0] !== currentLineNum) {
     // log(`${logPrefix} [selection] Selection spans multiple lines (${selStartActual[0]}-${selEndActual[0]}) or is not on the current focused table line (${currentLineNum}). Preventing default action.`);
      evt.preventDefault();
      return true; 
    }

    let selectionStartColInLine = selStartActual[1];
    let selectionEndColInLine = selEndActual[1];

    const currentCellFullText = cellTexts[targetCellIndex] || '';
    const cellContentStartColInLine = cellStartCol;
    const cellContentEndColInLine = cellStartCol + currentCellFullText.length;

    /* If the user selected the whole cell plus delimiter characters,
     * clamp the selection to just the cell content.                        */
    const hasTrailingDelim =
      targetCellIndex < metadataForTargetLine.cols - 1 &&
      selectionEndColInLine === cellContentEndColInLine + DELIMITER.length;

    const hasLeadingDelim =
      targetCellIndex > 0 &&
      selectionStartColInLine === cellContentStartColInLine - DELIMITER.length;

    console.log(`[ep_data_tables:highlight-deletion] Selection analysis:`, {
      targetCellIndex,
      totalCols: metadataForTargetLine.cols,
      selectionStartCol: selectionStartColInLine,
      selectionEndCol: selectionEndColInLine,
      cellContentStartCol: cellContentStartColInLine,
      cellContentEndCol: cellContentEndColInLine,
      delimiterLength: DELIMITER.length,
      expectedTrailingDelimiterPos: cellContentEndColInLine + DELIMITER.length,
      expectedLeadingDelimiterPos: cellContentStartColInLine - DELIMITER.length,
      hasTrailingDelim,
      hasLeadingDelim,
      cellText: currentCellFullText
    });

    if (hasLeadingDelim) {
      console.log(`[ep_data_tables:highlight-deletion] CLAMPING selection start from ${selectionStartColInLine} to ${cellContentStartColInLine}`);
      selectionStartColInLine = cellContentStartColInLine;
    }

    if (hasTrailingDelim) {
      console.log(`[ep_data_tables:highlight-deletion] CLAMPING selection end from ${selectionEndColInLine} to ${cellContentEndColInLine}`);
      selectionEndColInLine = cellContentEndColInLine;
    }

   // log(`${logPrefix} [selection] Cell context for selection: targetCellIndex=${targetCellIndex}, cellStartColInLine=${cellContentStartColInLine}, cellEndColInLine=${cellContentEndColInLine}, currentCellFullText='${currentCellFullText}'`);

    const isSelectionEntirelyWithinCell =
      selectionStartColInLine >= cellContentStartColInLine &&
      selectionEndColInLine <= cellContentEndColInLine;


    if (isSelectionEntirelyWithinCell) {

      if (evt.type !== 'keydown') return false;

      if (evt.ctrlKey || evt.metaKey || evt.altKey) return false;

    }

    const isCurrentKeyDelete = evt.key === 'Delete' || evt.keyCode === 46;
    const isCurrentKeyBackspace = evt.key === 'Backspace' || evt.keyCode === 8;
    const isCurrentKeyTyping = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;


    if (isSelectionEntirelyWithinCell && (isCurrentKeyDelete || isCurrentKeyBackspace || isCurrentKeyTyping)) {
     // log(`${logPrefix} [selection] Handling key='${evt.key}' (Type: ${evt.type}) for valid intra-cell selection.`);

      if (evt.type !== 'keydown') {
       // log(`${logPrefix} [selection] Ignoring non-keydown event type ('${evt.type}') for selection handling. Allowing default.`);
        return false; 
      }
      evt.preventDefault();

      const rangeStart = [currentLineNum, selectionStartColInLine];
      const rangeEnd = [currentLineNum, selectionEndColInLine];
      let replacementText = '';
      let newAbsoluteCaretCol = selectionStartColInLine;
      const repBeforeEdit = editorInfo.ace_getRep();
     // log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

      if (isCurrentKeyTyping) {
        replacementText = evt.key;
        newAbsoluteCaretCol = selectionStartColInLine + replacementText.length;
       // log(`${logPrefix} [selection] -> Replacing selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]] with text '${replacementText}'`);
      } else {
       // log(`${logPrefix} [selection] -> Deleting selected range [[${rangeStart[0]},${rangeStart[1]}],[${rangeEnd[0]},${rangeEnd[1]}]]`);
        const isWholeCell = selectionStartColInLine <= cellContentStartColInLine && selectionEndColInLine >= cellContentEndColInLine;
        if (isWholeCell) {
          replacementText = ' ';
          newAbsoluteCaretCol = selectionStartColInLine + 1;
         // log(`${logPrefix} [selection] Whole cell cleared – inserting single space to preserve caret/author span.`);
        }
      }

      try {
        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, replacementText);

        if (replacementText.length > 0) {
          const attrStart = [currentLineNum, selectionStartColInLine];
          const attrEnd   = [currentLineNum, selectionStartColInLine + replacementText.length];
          console.log(`[ep_data_tables:highlight-deletion] Applying cell attribute to replacement text "${replacementText}" at range [${attrStart[0]},${attrStart[1]}] to [${attrEnd[0]},${attrEnd[1]}]`);
          editorInfo.ace_performDocumentApplyAttributesToRange(
            attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
          );
        }
        const repAfterReplace = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


       // log(`${logPrefix} [selection] -> Re-applying tbljson line attribute...`);
        const applyHelper = editorInfo.ep_data_tables_applyMeta;
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) {
          const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;
          applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
         // log(`${logPrefix} [selection] -> tbljson line attribute re-applied (using rep before edit).`);
        } else {
          console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
          const currentRepFallback = editorInfo.ace_getRep();
          if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
           // log(`${logPrefix} [selection] -> Retrying attribute application with current rep...`);
            applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
           // log(`${logPrefix} [selection] -> tbljson line attribute re-applied (using current rep fallback).`);
          } else {
            console.error(`${logPrefix} [selection] -> FAILED to re-apply tbljson attribute even with fallback rep.`);
          }
        }

       // log(`${logPrefix} [selection] -> Setting selection/caret to: [${currentLineNum}, ${newAbsoluteCaretCol}]`);
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterSelectionChange = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);

    editorInfo.ace_fastIncorp(1);
        const repAfterFastIncorp = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
       // log(`${logPrefix} [selection] -> Requested sync hint (fastIncorp 1).`);

       // log(`${logPrefix} [caretTrace] [selection] Attempting to re-assert selection post-fastIncorp to [${currentLineNum}, ${newAbsoluteCaretCol}]`);
        editorInfo.ace_performSelectionChange([currentLineNum, newAbsoluteCaretCol], [currentLineNum, newAbsoluteCaretCol], false);
        const repAfterReassert = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

        const newRelativePos = newAbsoluteCaretCol - cellStartCol;
        if (editor) {
            editor.ep_data_tables_last_clicked = {
                lineNum: currentLineNum,
                tblId: metadataForTargetLine.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos
            };
           // log(`${logPrefix} [selection] -> Updated stored click/caret info:`, editor.ep_data_tables_last_clicked);
        } else {
           // log(`${logPrefix} [selection] -> Editor instance not found, cannot update ep_data_tables_last_clicked.`);
        }

       // log(`${logPrefix} END [selection] (Handled highlight modification) Key='${evt.key}' Type='${evt.type}'. Duration: ${Date.now() - startLogTime}ms`);
      return true;
      } catch (error) {
       // log(`${logPrefix} [selection] ERROR during highlight modification:`, error);
        console.error('[ep_data_tables] Error processing highlight modification:', error);
        return true;
      }
    }
  }

  const isCutKey = (evt.ctrlKey || evt.metaKey) && (evt.key === 'x' || evt.key === 'X' || evt.keyCode === 88);
  if (isCutKey && hasSelection) {
   // log(`${logPrefix} Ctrl+X (Cut) detected with selection. Letting cut event handler manage this.`);
    return false;
  } else if (isCutKey && !hasSelection) {
   // log(`${logPrefix} Ctrl+X (Cut) detected but no selection. Allowing default.`);
    return false;
  }

  const isTypingKey = evt.key && evt.key.length === 1 && !evt.ctrlKey && !evt.metaKey && !evt.altKey;
  const isDeleteKey = evt.key === 'Delete' || evt.keyCode === 46;
  const isBackspaceKey = evt.key === 'Backspace' || evt.keyCode === 8;
  const isNavigationKey = [33, 34, 35, 36, 37, 38, 39, 40].includes(evt.keyCode);
  const isTabKey = evt.key === 'Tab';
  const isEnterKey = evt.key === 'Enter';
 // log(`${logPrefix} Key classification: Typing=${isTypingKey}, Backspace=${isBackspaceKey}, Delete=${isDeleteKey}, Nav=${isNavigationKey}, Tab=${isTabKey}, Enter=${isEnterKey}, Cut=${isCutKey}`);

  const currentCellTextLengthEarly = cellTexts[targetCellIndex]?.length ?? 0;

  if (evt.type === 'keydown' && !evt.ctrlKey && !evt.metaKey && !evt.altKey) {
    if (evt.keyCode === 39 && relativeCaretPos >= currentCellTextLengthEarly && targetCellIndex < metadataForTargetLine.cols - 1) {
     // log(`${logPrefix} ArrowRight at cell boundary – navigating to next cell to avoid anchor zone.`);
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, false, editorInfo, docManager);
      return true;
    }

    if (evt.keyCode === 37 && relativeCaretPos === 0 && targetCellIndex > 0) {
     // log(`${logPrefix} ArrowLeft at cell boundary – navigating to previous cell to avoid anchor zone.`);
      evt.preventDefault();
      navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, true, editorInfo, docManager);
      return true;
    }
  }


  if (isNavigationKey && !isTabKey) {
     // log(`${logPrefix} Allowing navigation key: ${evt.key}. Clearing click state.`);
      if (editor) editor.ep_data_tables_last_clicked = null;
      return false;
  }

  if (isTabKey) { 
    // log(`${logPrefix} Tab key pressed. Event type: ${evt.type}`);
    evt.preventDefault();

     if (evt.type !== 'keydown') {
      // log(`${logPrefix} Ignoring Tab ${evt.type} event to prevent double navigation.`);
    return true;
  }

    // log(`${logPrefix} Processing Tab keydown - implementing cell navigation.`);
     const success = navigateToNextCell(currentLineNum, targetCellIndex, metadataForTargetLine, evt.shiftKey, editorInfo, docManager);
     if (!success) {
      // log(`${logPrefix} Tab navigation failed, cell navigation not possible.`);
     }
     return true;
  }

  if (isEnterKey) {
     // log(`${logPrefix} Enter key pressed. Event type: ${evt.type}`);
    evt.preventDefault();

      if (evt.type !== 'keydown') {
       // log(`${logPrefix} Ignoring Enter ${evt.type} event to prevent double navigation.`);
    return true;
  }

     // log(`${logPrefix} Processing Enter keydown - implementing cell navigation.`);
      const success = navigateToCellBelow(currentLineNum, targetCellIndex, metadataForTargetLine, editorInfo, docManager);
      if (!success) {
       // log(`${logPrefix} Enter navigation failed, cell navigation not possible.`);
      }
      return true; 
  }

      const currentCellTextLength = cellTexts[targetCellIndex]?.length ?? 0;
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex > 0) {
     // log(`${logPrefix} Intercepted Backspace at start of cell ${targetCellIndex}. Preventing default.`);
    evt.preventDefault();
          return true;
      }
      if (isBackspaceKey && relativeCaretPos === 0 && targetCellIndex === 0) {
       // log(`${logPrefix} Intercepted Backspace at start of first cell (line boundary). Preventing merge.`);
        evt.preventDefault();
        return true;
      }
  if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex < metadataForTargetLine.cols - 1) {
     // log(`${logPrefix} Intercepted Delete at end of cell ${targetCellIndex}. Preventing default.`);
          evt.preventDefault();
          return true;
      }
      if (isDeleteKey && relativeCaretPos === currentCellTextLength && targetCellIndex === metadataForTargetLine.cols - 1) {
       // log(`${logPrefix} Intercepted Delete at end of last cell (line boundary). Preventing merge.`);
        evt.preventDefault();
        return true;
      }

  const isInternalBackspace = isBackspaceKey && relativeCaretPos > 0;
  const isInternalDelete = isDeleteKey && relativeCaretPos < currentCellTextLength;

  if ((isInternalBackspace && relativeCaretPos === 1 && targetCellIndex > 0) ||
      (isInternalDelete && relativeCaretPos === 0 && targetCellIndex > 0)) {
   // log(`${logPrefix} Attempt to erase protected delimiter – operation blocked.`);
    evt.preventDefault();
    return true;
  }

  if (isTypingKey || isInternalBackspace || isInternalDelete) {
    if (isTypingKey && relativeCaretPos === 0 && targetCellIndex > 0) {
     // log(`${logPrefix} Caret at forbidden position 0 (just after delimiter). Auto-advancing to position 1.`);
      const safePosAbs = cellStartCol + 1;
      editorInfo.ace_performSelectionChange([currentLineNum, safePosAbs], [currentLineNum, safePosAbs], false);
      editorInfo.ace_updateBrowserSelectionFromRep();
      relativeCaretPos = 1;
     // log(`${logPrefix} Caret moved to safe position. New relativeCaretPos=${relativeCaretPos}`);
    }
    const currentCol = cellStartCol + relativeCaretPos;
   // log(`${logPrefix} Handling INTERNAL key='${evt.key}' Type='${evt.type}' at Line=${currentLineNum}, Col=${currentCol} (CellIndex=${targetCellIndex}, RelativePos=${relativeCaretPos}).`);
   // log(`${logPrefix} [caretTrace] Initial rep.selStart for internal edit: Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);

    if (evt.type !== 'keydown') {
       // log(`${logPrefix} Ignoring non-keydown event type ('${evt.type}') for handled key.`);
        return false; 
    }

   // log(`${logPrefix} Preventing default browser action for keydown event.`);
    evt.preventDefault();

    let newAbsoluteCaretCol = -1;
    let repBeforeEdit = null;

    try {
        repBeforeEdit = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] rep.selStart before ace_performDocumentReplaceRange: Line=${repBeforeEdit.selStart[0]}, Col=${repBeforeEdit.selStart[1]}`);

    if (isTypingKey) {
            const insertPos = [currentLineNum, currentCol];
           // log(`${logPrefix} -> Inserting text '${evt.key}' at [${insertPos}]`);
            editorInfo.ace_performDocumentReplaceRange(insertPos, insertPos, evt.key);
            newAbsoluteCaretCol = currentCol + 1;

        } else if (isInternalBackspace) {
            const delRangeStart = [currentLineNum, currentCol - 1];
            const delRangeEnd = [currentLineNum, currentCol];
           // log(`${logPrefix} -> Deleting (Backspace) range [${delRangeStart}]-[${delRangeEnd}]`);
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol - 1;

        } else if (isInternalDelete) {
            const delRangeStart = [currentLineNum, currentCol];
            const delRangeEnd = [currentLineNum, currentCol + 1];
           // log(`${logPrefix} -> Deleting (Delete) range [${delRangeStart}]-[${delRangeEnd}]`);
            editorInfo.ace_performDocumentReplaceRange(delRangeStart, delRangeEnd, '');
            newAbsoluteCaretCol = currentCol;
        }
        const repAfterReplace = editorInfo.ace_getRep();
       // log(`${logPrefix} [caretTrace] rep.selStart after ace_performDocumentReplaceRange: Line=${repAfterReplace.selStart[0]}, Col=${repAfterReplace.selStart[1]}`);


       // log(`${logPrefix} -> Re-applying tbljson line attribute...`);

       // log(`${logPrefix} DEBUG: Before calculating attrStringToApply - trustedLastClick=${trustedLastClick}, reportedLineNum=${reportedLineNum}, currentLineNum=${currentLineNum}`);
       // log(`${logPrefix} DEBUG: lineAttrString value:`, lineAttrString ? `"${lineAttrString}"` : 'null/undefined');

        const applyHelper = editorInfo.ep_data_tables_applyMeta; 
        if (applyHelper && typeof applyHelper === 'function' && repBeforeEdit) { 
             const attrStringToApply = (trustedLastClick || reportedLineNum === currentLineNum) ? lineAttrString : null;

            // log(`${logPrefix} DEBUG: Calculated attrStringToApply:`, attrStringToApply ? `"${attrStringToApply}"` : 'null/undefined');
            // log(`${logPrefix} DEBUG: Condition result: (${trustedLastClick} || ${reportedLineNum} === ${currentLineNum}) = ${trustedLastClick || reportedLineNum === currentLineNum}`);

             applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, repBeforeEdit, editorInfo, attrStringToApply, docManager);
            // log(`${logPrefix} -> tbljson line attribute re-applied (using rep before edit).`);
                } else {
             console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute (helper or repBeforeEdit missing).`);
             const currentRepFallback = editorInfo.ace_getRep();
             if (applyHelper && typeof applyHelper === 'function' && currentRepFallback) {
                // log(`${logPrefix} -> Retrying attribute application with current rep...`);
                 applyHelper(currentLineNum, metadataForTargetLine.tblId, metadataForTargetLine.row, metadataForTargetLine.cols, currentRepFallback, editorInfo, null, docManager);
                // log(`${logPrefix} -> tbljson line attribute re-applied (using current rep fallback).`);
            } else {
                  console.error(`${logPrefix} -> FAILED to re-apply tbljson attribute even with fallback rep.`);
             }
        }

        if (newAbsoluteCaretCol >= 0) {
             const newCaretPos = [currentLineNum, newAbsoluteCaretCol];
            // log(`${logPrefix} -> Setting selection immediately to:`, newCaretPos);
            // log(`${logPrefix} [caretTrace] rep.selStart before ace_performSelectionChange: Line=${editorInfo.ace_getRep().selStart[0]}, Col=${editorInfo.ace_getRep().selStart[1]}`);
             try {
                editorInfo.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                const repAfterSelectionChange = editorInfo.ace_getRep();
               // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_performSelectionChange: Line=${repAfterSelectionChange.selStart[0]}, Col=${repAfterSelectionChange.selStart[1]}`);
               // log(`${logPrefix} -> Selection set immediately.`);

                editorInfo.ace_fastIncorp(1); 
                const repAfterFastIncorp = editorInfo.ace_getRep();
               // log(`${logPrefix} [caretTrace] [selection] rep.selStart after ace_fastIncorp: Line=${repAfterFastIncorp.selStart[0]}, Col=${repAfterFastIncorp.selStart[1]}`);
               // log(`${logPrefix} -> Requested sync hint (fastIncorp 1).`);

                const targetCaretPosForReassert = [currentLineNum, newAbsoluteCaretCol];
               // log(`${logPrefix} [caretTrace] Attempting to re-assert selection post-fastIncorp to [${targetCaretPosForReassert[0]}, ${targetCaretPosForReassert[1]}]`);
                editorInfo.ace_performSelectionChange(targetCaretPosForReassert, targetCaretPosForReassert, false);
                const repAfterReassert = editorInfo.ace_getRep();
               // log(`${logPrefix} [caretTrace] [selection] rep.selStart after re-asserting selection: Line=${repAfterReassert.selStart[0]}, Col=${repAfterReassert.selStart[1]}`);

                const newRelativePos = newAbsoluteCaretCol - cellStartCol;
                editor.ep_data_tables_last_clicked = {
                    lineNum: currentLineNum, 
                    tblId: metadataForTargetLine.tblId,
                    cellIndex: targetCellIndex,
                    relativePos: newRelativePos
                };
               // log(`${logPrefix} -> Updated stored click/caret info:`, editor.ep_data_tables_last_clicked);
               // log(`${logPrefix} [caretTrace] Updated ep_data_tables_last_clicked. Line=${editor.ep_data_tables_last_clicked.lineNum}, Cell=${editor.ep_data_tables_last_clicked.cellIndex}, RelPos=${editor.ep_data_tables_last_clicked.relativePos}`);


            } catch (selError) {
                 console.error(`${logPrefix} -> ERROR setting selection immediately:`, selError);
             }
        } else {
           // log(`${logPrefix} -> Warning: newAbsoluteCaretCol not set, skipping selection update.`);
            }

        } catch (error) {
       // log(`${logPrefix} ERROR during manual key handling:`, error);
            console.error('[ep_data_tables] Error processing key event update:', error);
    return true;
  }

    const endLogTime = Date.now();
   // log(`${logPrefix} END (Handled Internal Edit Manually) Key='${evt.key}' Type='${evt.type}' -> Returned true. Duration: ${endLogTime - startLogTime}ms`);
    return true;

  }


  const endLogTimeFinal = Date.now();
 // log(`${logPrefix} END (Fell Through / Unhandled Case) Key='${evt.key}' Type='${evt.type}'. Allowing default. Duration: ${endLogTimeFinal - startLogTime}ms`);
 // log(`${logPrefix} [caretTrace] Final rep.selStart at end of aceKeyEvent (if unhandled): Line=${rep.selStart[0]}, Col=${rep.selStart[1]}`);
  return false;
};
