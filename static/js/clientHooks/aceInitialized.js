'use strict';

const {
  ATTR_TABLE_JSON,
  ATTR_CELL,
  DELIMITER,
  INPUTTYPE_REPLACEMENT_TYPES,
  getTableLineMetadata,
  navigateToCellBelow,
  normalizeSoftWhitespace,
  isAndroidUA,
  isIOSUA,
  rand,
} = require('./shared');

let isResizing = false;
let resizeStartX = 0;
let resizeCurrentX = 0;
let resizeTargetTable = null;
let resizeTargetColumn = -1;
let resizeOriginalWidths = [];
let resizeTableMetadata = null;
let resizeLineNum = -1;
let resizeOverlay = null;
let suppressNextBeforeInputInsertTextOnce = false;
let isAndroidChromeComposition = false;
let handledCurrentComposition = false;
let suppressBeforeInputInsertTextDuringComposition = false;

const aceInitialized = (_hook, ctx) => {
  const logPrefix = '[ep_data_tables:aceInitialized]';
 // log(`${logPrefix} START`, { hook_name: h, context: ctx });
  const ed = ctx.editorInfo;
  const docManager = ctx.documentAttributeManager;

 // log(`${logPrefix} Attaching ep_data_tables_applyMeta helper to editorInfo.`);
  ed.ep_data_tables_applyMeta = applyTableLineMetadataAttribute;
 // log(`${logPrefix}: Attached applyTableLineMetadataAttribute helper to ed.ep_data_tables_applyMeta successfully.`);

 // log(`${logPrefix} Storing documentAttributeManager reference on editorInfo.`);
  ed.ep_data_tables_docManager = docManager;
 // log(`${logPrefix}: Stored documentAttributeManager reference as ed.ep_data_tables_docManager.`);

 // log(`${logPrefix} Preparing to attach paste and resize listeners via ace_callWithAce.`);
  ed.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = '[ep_data_tables:aceInitialized:callWithAceForListeners]';
   // log(`${callWithAceLogPrefix} Entered ace_callWithAce callback for listeners.`);

    if (!ace || !ace.editor) {
      console.error(`${callWithAceLogPrefix} ERROR: ace or ace.editor is not available. Cannot attach listeners.`);
     // log(`${callWithAceLogPrefix} Aborting listener attachment due to missing ace.editor.`);
      return;
    }
    const editor = ace.editor;
   // log(`${callWithAceLogPrefix} ace.editor obtained successfully.`);

   // log(`${logPrefix} Storing editor reference on editorInfo.`);
    ed.ep_data_tables_editor = editor;
   // log(`${logPrefix}: Stored editor reference as ed.ep_data_tables_editor.`);

    let $inner;
    try {
     // log(`${callWithAceLogPrefix} Attempting to find inner iframe body for listener attachment.`);
      const $iframeOuter = $('iframe[name="ace_outer"]');
      if ($iframeOuter.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find outer iframe (ace_outer).`);
       // log(`${callWithAceLogPrefix} Failed to find ace_outer.`);
        return;
      }
     // log(`${callWithAceLogPrefix} Found ace_outer:`, $iframeOuter);

      const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
      if ($iframeInner.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find inner iframe (ace_inner).`);
       // log(`${callWithAceLogPrefix} Failed to find ace_inner within ace_outer.`);
        return;
      }
     // log(`${callWithAceLogPrefix} Found ace_inner:`, $iframeInner);

      const innerDocBody = $iframeInner.contents().find('body');
      if (innerDocBody.length === 0) {
        console.error(`${callWithAceLogPrefix} ERROR: Could not find body element in inner iframe.`);
       // log(`${callWithAceLogPrefix} Failed to find body in ace_inner.`);
        return;
      }
      $inner = $(innerDocBody[0]);
     // log(`${callWithAceLogPrefix} Successfully found inner iframe body:`, $inner);

      const mobileSuggestionBlocker = (evt) => {
        const t = evt && evt.inputType || '';
        const dataStr = (evt && typeof evt.data === 'string') ? evt.data : '';
        const isProblem = (
          t === 'insertReplacementText' ||
          t === 'insertFromComposition' ||
          (t === 'insertText' && !evt.isComposing && (!dataStr || dataStr.length > 1))
        );
        if (!isProblem) return;

        try {
          const repQuick = ed.ace_getRep && ed.ace_getRep();
          if (!repQuick || !repQuick.selStart) return;
          const lineNumQuick = repQuick.selStart[0];
          let metaStrQuick = docManager && docManager.getAttributeOnLine
            ? docManager.getAttributeOnLine(lineNumQuick, ATTR_TABLE_JSON)
            : null;
          let metaQuick = null;
          if (metaStrQuick) { try { metaQuick = JSON.parse(metaStrQuick); } catch (_) {} }
          if (!metaQuick) metaQuick = getTableLineMetadata(lineNumQuick, ed, docManager);
          if (!metaQuick || typeof metaQuick.cols !== 'number') return;
        } catch (_) { return; }

        evt._epDataTablesHandled = true;
        if (evt.originalEvent) evt.originalEvent._epDataTablesHandled = true;
        evt.preventDefault();
        if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

        const capturedInputType = t;
        const capturedData = dataStr;

        setTimeout(() => {
          try {
            ed.ace_callWithAce((aceInstance) => {
              aceInstance.ace_fastIncorp(10);
              const rep = aceInstance.ace_getRep();
              if (!rep || !rep.selStart || !rep.selEnd) return;

              const selStart = [...rep.selStart];
              const selEnd = [...rep.selEnd];
              const lineNum = selStart[0];

              let metaStr = docManager && docManager.getAttributeOnLine
                ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON)
                : null;
              let tableMetadata = null;
              if (metaStr) { try { tableMetadata = JSON.parse(metaStr); } catch (_) {} }
              if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
              if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined') {
                return;
              }

              const initialHasSelection = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);

              let replacement = typeof capturedData === 'string' ? capturedData : '';
              if (!replacement) {
                if (capturedInputType === 'insertText' && !initialHasSelection) {
                  replacement = ' ';
                } else if (capturedInputType === 'insertReplacementText' || capturedInputType === 'insertFromComposition' || initialHasSelection) {
                  replacement = ' ';
                }
              }

              replacement = normalizeSoftWhitespace(
                (replacement || '')
                  .replace(new RegExp(DELIMITER, 'g'), ' ')
                  .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
              );

              if (!replacement) replacement = ' ';

              const lineEntry = rep.lines.atIndex(lineNum);
              const lineText = lineEntry?.text || '';
              const cells = lineText.split(DELIMITER);
              let currentOffset = 0;
              let targetCellIndex = -1;
              let cellStartCol = 0;
              let cellEndCol = 0;
              for (let i = 0; i < cells.length; i++) {
                const cellLen = cells[i]?.length ?? 0;
                const cellEndThis = currentOffset + cellLen;
                if (selStart[1] >= currentOffset && selStart[1] <= cellEndThis) {
                  targetCellIndex = i;
                  cellStartCol = currentOffset;
                  cellEndCol = cellEndThis;
                  break;
                }
                currentOffset += cellLen + DELIMITER.length;
              }

              if (targetCellIndex === -1) {
                aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, replacement);
                const repAfterFallback = aceInstance.ace_getRep();
                ed.ep_data_tables_applyMeta(
                  lineNum,
                  tableMetadata.tblId,
                  tableMetadata.row,
                  tableMetadata.cols,
                  repAfterFallback,
                  ed,
                  null,
                  docManager
                );
                const fallbackLineEntry = repAfterFallback.lines.atIndex(lineNum);
                const fallbackMaxLen = fallbackLineEntry ? fallbackLineEntry.text.length : 0;
                const fallbackStartCol = Math.min(Math.max(selStart[1], 0), fallbackMaxLen);
                const fallbackEndCol = Math.min(fallbackStartCol + replacement.length, fallbackMaxLen);
                const fallbackCaretPos = [lineNum, fallbackEndCol];
                aceInstance.ace_performSelectionChange(fallbackCaretPos, fallbackCaretPos, false);
                aceInstance.ace_fastIncorp(10);
                return;
              }

              if (selEnd[0] !== selStart[0]) {
                selEnd[0] = selStart[0];
                selEnd[1] = cellEndCol;
              }

              if (selEnd[1] > cellEndCol) {
                selEnd[1] = Math.min(selEnd[1], cellEndCol);
              }

              if (selEnd[1] < selStart[1]) selEnd[1] = selStart[1];

              aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, replacement);

              const repAfter = aceInstance.ace_getRep();
              const lineEntryAfter = repAfter.lines.atIndex(lineNum);
              const maxLen = lineEntryAfter ? lineEntryAfter.text.length : 0;
              const startCol = Math.min(Math.max(selStart[1], 0), maxLen);
              const endCol = Math.min(startCol + replacement.length, maxLen);

              if (endCol > startCol) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [lineNum, startCol],
                  [lineNum, endCol],
                  [[ATTR_CELL, String(targetCellIndex)]]
                );
              }

              ed.ep_data_tables_applyMeta(
                lineNum,
                tableMetadata.tblId,
                tableMetadata.row,
                tableMetadata.cols,
                repAfter,
                ed,
                null,
                docManager
              );

              const newCaretPos = [lineNum, endCol];
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              aceInstance.ace_fastIncorp(10);

              const editor = ed.ep_data_tables_editor;
              if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
                const freshLineText = lineEntryAfter ? lineEntryAfter.text : '';
                const freshCells = freshLineText.split(DELIMITER);
                let freshOffset = 0;
                for (let i = 0; i < targetCellIndex; i++) {
                  freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
                }
                const newRelativePos = newCaretPos[1] - freshOffset;
                editor.ep_data_tables_last_clicked = {
                  lineNum,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
                };
              }
            }, 'mobileSuggestionBlocker', true);
          } catch (e) {
            console.error('[ep_data_tables:mobileSuggestionBlocker] Error applying predictive text:', e);
          }
        }, 0);
      };

      // IME/autocorrect diagnostics: capture-phase logging and newline soft-normalization for table lines
      const logIMEEvent = (rawEvt, tag) => {
        try {
          const e = rawEvt && (rawEvt.originalEvent || rawEvt);
          const rep = ed.ace_getRep && ed.ace_getRep();
          const selStart = rep && rep.selStart;
          const lineNum = selStart ? selStart[0] : -1;
          let isTableLine = false;
          if (lineNum >= 0) {
            let s = docManager && docManager.getAttributeOnLine ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON) : null;
            if (!s) {
              const meta = getTableLineMetadata(lineNum, ed, docManager);
              isTableLine = !!meta && typeof meta.cols === 'number';
            } else {
              isTableLine = true;
            }
          }
          if (!isTableLine) return;
          const payload = {
            tag,
            type: e && e.type,
            inputType: e && e.inputType,
            data: typeof (e && e.data) === 'string' ? e.data : null,
            isComposing: !!(e && e.isComposing),
            key: e && e.key,
            code: e && e.code,
            which: e && e.which,
            keyCode: e && e.keyCode,
          };
          console.debug('[ep_data_tables:ime-diag]', payload);
        } catch (_) {}
      };

      const softBreakNormalizer = (rawEvt) => {
        try {
          const e = rawEvt && (rawEvt.originalEvent || rawEvt);
          if (!e || e._epDataTablesNormalized) return;
          const t = e.inputType || '';
          const dataStr = typeof e.data === 'string' ? e.data : '';
          const hasLF = /[\r\n]/.test(dataStr);
          const isSoftBreak = t === 'insertParagraph' || t === 'insertLineBreak' || hasLF;
          if (!isSoftBreak) return;

          const rep = ed.ace_getRep && ed.ace_getRep();
          if (!rep || !rep.selStart) return;
          const lineNum = rep.selStart[0];
          let metaStr = docManager && docManager.getAttributeOnLine ? docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON) : null;
          let meta = null;
          if (metaStr) { try { meta = JSON.parse(metaStr); } catch (_) {} }
          if (!meta) meta = getTableLineMetadata(lineNum, ed, docManager);
          if (!meta || typeof meta.cols !== 'number') return;

          e.preventDefault();
          if (typeof e.stopImmediatePropagation === 'function') e.stopImmediatePropagation();
          e._epDataTablesNormalized = true;
          setTimeout(() => {
            try {
              ed.ace_callWithAce((aceInstance) => {
                aceInstance.ace_fastIncorp(10);
                const rep2 = aceInstance.ace_getRep();
                const start = rep2.selStart;
                const end = rep2.selEnd;
                // Replace the attempted soft-break with a single space
                aceInstance.ace_performDocumentReplaceRange(start, end, ' ');
              }, 'softBreakNormalizer', true);
            } catch (err) { console.error('[ep_data_tables:softBreakNormalizer] error', err); }
          }, 0);
        } catch (_) {}
      };

      if ($inner && $inner.length > 0 && $inner[0].addEventListener) {
        const el = $inner[0];
        ['beforeinput','input','textInput','compositionstart','compositionupdate','compositionend','keydown','keyup'].forEach((t) => {
          el.addEventListener(t, (ev) => logIMEEvent(ev, 'capture'), true);
        });
        el.addEventListener('beforeinput', softBreakNormalizer, true);
      }

      if ($inner && $inner.length > 0 && $inner[0].addEventListener) {
        $inner[0].addEventListener('beforeinput', mobileSuggestionBlocker, true);
      }

      if (isAndroidUA && isAndroidUA() && $inner && $inner.length > 0 && $inner[0].addEventListener) {
        $inner[0].addEventListener('textInput', (evt) => {
          const s = typeof evt.data === 'string' ? evt.data : '';
          if (s && s.length > 1) {
            mobileSuggestionBlocker({
              inputType: 'insertText',
              isComposing: false,
              data: s,
              preventDefault: () => { try { evt.preventDefault(); } catch (_) {} },
              stopImmediatePropagation: () => { try { evt.stopImmediatePropagation(); } catch (_) {} },
            });
          }
        }, true);
      }

      try {
        const disableAuto = (el) => {
          if (!el) return;
          el.setAttribute('autocorrect', 'off');
          el.setAttribute('autocomplete', 'off');
          el.setAttribute('autocapitalize', 'off');
          el.setAttribute('spellcheck', 'false');
        };
        disableAuto(innerDocBody[0] || innerDocBody);
      } catch (_) {}
    } catch (e) {
      console.error(`${callWithAceLogPrefix} ERROR: Exception while trying to find inner iframe body:`, e);
     // log(`${callWithAceLogPrefix} Exception details:`, { message: e.message, stack: e.stack });
      return;
    }

    if (!$inner || $inner.length === 0) {
      console.error(`${callWithAceLogPrefix} ERROR: $inner is not valid after attempting to find iframe body. Cannot attach listeners.`);
     // log(`${callWithAceLogPrefix} $inner is invalid. Aborting.`);
      return;
    }

   // log(`${callWithAceLogPrefix} Attaching cut event listener to $inner (inner iframe body).`);
    $inner.on('cut', (evt) => {
      const cutLogPrefix = '[ep_data_tables:cutHandler]';
      console.log(`${cutLogPrefix} CUT EVENT TRIGGERED. Event object:`, evt);

      console.log(`${cutLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        console.warn(`${cutLogPrefix} WARNING: Could not get representation or selection. Allowing default cut.`);
        console.warn(`${cutLogPrefix} Could not get rep or selStart.`);
        return;
      }
      console.log(`${cutLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
      console.log(`${cutLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);
      const hasSelectionInRep = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);
      if (!hasSelectionInRep) {
        console.log(`${cutLogPrefix} No selection detected in rep; deferring decision until table-line check.`);
      }

      if (selStart[0] !== selEnd[0]) {
        console.warn(`${cutLogPrefix} WARNING: Selection spans multiple lines. Preventing cut to protect table structure.`);
        evt.preventDefault();
        return;
      }

      console.log(`${cutLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (lineAttrString) {
        try {
        tableMetadata = JSON.parse(lineAttrString);
        } catch {}
      }

      if (!tableMetadata) {
        tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      }

      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        console.log(`${cutLogPrefix} Line ${lineNum} is NOT a recognised table line. Allowing default cut.`);
        return;
      }

      console.log(`${cutLogPrefix} Line ${lineNum} IS a table line. Metadata:`, tableMetadata);

      if (!hasSelectionInRep) {
        console.log(`${cutLogPrefix} Preventing default CUT on table line with collapsed selection to protect delimiters.`);
        evt.preventDefault();
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;

      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;

        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      /* allow "…cell content + delimiter" selections */
      const wouldClampStart = targetCellIndex > 0 && selStart[1] === cellStartCol - DELIMITER.length;
      const wouldClampEnd = targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length;

      console.log(`[ep_data_tables:cut-handler] Cut selection analysis:`, {
        targetCellIndex,
        selStartCol: selStart[1],
        selEndCol: selEnd[1],
        cellStartCol,
        cellEndCol,
        delimiterLength: DELIMITER.length,
        expectedLeadingDelimiterPos: cellStartCol - DELIMITER.length,
        expectedTrailingDelimiterPos: cellEndCol + DELIMITER.length,
        wouldClampStart,
        wouldClampEnd
      });

      if (wouldClampStart) {
        console.log(`[ep_data_tables:cut-handler] CLAMPING cut selection start from ${selStart[1]} to ${cellStartCol}`);
        selStart[1] = cellStartCol;
      }

      if (wouldClampEnd) {
        console.log(`[ep_data_tables:cut-handler] CLAMPING cut selection end from ${selEnd[1]} to ${cellEndCol}`);
        selEnd[1] = cellEndCol;
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        console.warn(`${cutLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing cut to protect table structure.`);
        evt.preventDefault();
        return;
      }

      console.log(`${cutLogPrefix} Selection is entirely within cell ${targetCellIndex}. Intercepting cut to preserve table structure.`);
      evt.preventDefault();

      try {
        const selectedText = lineText.substring(selStart[1], selEnd[1]);
        console.log(`${cutLogPrefix} Selected text to cut: "${selectedText}"`);

        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(selectedText).then(() => {
            console.log(`${cutLogPrefix} Successfully copied to clipboard via Navigator API.`);
          }).catch((err) => {
            console.warn(`${cutLogPrefix} Failed to copy to clipboard via Navigator API:`, err);
          });
        } else {
          console.log(`${cutLogPrefix} Using fallback clipboard method.`);
          const textArea = document.createElement('textarea');
          textArea.value = selectedText;
          document.body.appendChild(textArea);
          textArea.select();
          try {
            document.execCommand('copy');
            console.log(`${cutLogPrefix} Successfully copied to clipboard via execCommand fallback.`);
          } catch (err) {
            console.warn(`${cutLogPrefix} Failed to copy to clipboard via fallback:`, err);
          }
          document.body.removeChild(textArea);
        }

        console.log(`${cutLogPrefix} Performing deletion via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${cutLogPrefix}[ace_callWithAceOps]`;
          console.log(`${callAceLogPrefix} Entered ace_callWithAce for cut operations. selStart:`, selStart, `selEnd:`, selEnd);

          console.log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to delete selected text.`);
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');
          console.log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
            console.log(`${callAceLogPrefix} Cell ${targetCellIndex} became empty after cut – inserting single space to preserve structure.`);
            const insertPos = [lineNum, selStart[1]];
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            const attrStart = insertPos;
            const attrEnd   = [insertPos[0], insertPos[1] + 1];
            aceInstance.ace_performDocumentApplyAttributesToRange(
              attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

          console.log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
          const repAfterCut = aceInstance.ace_getRep();
          console.log(`${callAceLogPrefix} Fetched rep after cut for applyMeta. Line ${lineNum} text now: "${repAfterCut.lines.atIndex(lineNum).text}"`);

          ed.ep_data_tables_applyMeta(
            lineNum,
            tableMetadata.tblId,
            tableMetadata.row,
            tableMetadata.cols,
            repAfterCut,
            ed,
            null,
            docManager
          );
          console.log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

          const newCaretPos = [lineNum, selStart[1]];
          console.log(`${callAceLogPrefix} Setting caret position to: [${newCaretPos}].`);
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
          console.log(`${callAceLogPrefix} Selection change successful.`);

          console.log(`${callAceLogPrefix} Cut operations within ace_callWithAce completed successfully.`);
        }, 'tableCutTextOperations', true);

        console.log(`${cutLogPrefix} Cut operation completed successfully.`);
      } catch (error) {
        console.error(`${cutLogPrefix} ERROR during cut operation:`, error);
        console.log(`${cutLogPrefix} Cut operation failed. Error details:`, { message: error.message, stack: error.stack });
      }
    });

   // log(`${callWithAceLogPrefix} Attaching beforeinput event listener to $inner (inner iframe body).`);
    $inner.on('beforeinput', (evt) => {
      const deleteLogPrefix = '[ep_data_tables:beforeinputDeleteHandler]';
     // log(`${deleteLogPrefix} BEFOREINPUT EVENT TRIGGERED. inputType: "${evt.originalEvent.inputType}", event object:`, evt);

      if (!evt.originalEvent.inputType || !evt.originalEvent.inputType.startsWith('delete')) {
       // log(`${deleteLogPrefix} Not a deletion event (inputType: "${evt.originalEvent.inputType}"). Allowing default.`);
        return;
      }

     // log(`${deleteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
       // log(`${deleteLogPrefix} WARNING: Could not get representation or selection. Allowing default delete.`);
        console.warn(`${deleteLogPrefix} Could not get rep or selStart.`);
        return;
      }
     // log(`${deleteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
     // log(`${deleteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      const isAndroidChrome = isAndroidUA();
      const inputType = (evt.originalEvent && evt.originalEvent.inputType) || '';

      const isCollapsed = (selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);
      if (isCollapsed && isAndroidChrome && (inputType === 'deleteContentBackward' || inputType === 'deleteContentForward')) {
        let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
        let tableMetadata = null;
        if (lineAttrString) { try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {} }
        if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number') {
          return;
        }

        const lineText = rep.lines.atIndex(lineNum)?.text || '';
        const cells = lineText.split(DELIMITER);
        let currentOffset = 0;
        let targetCellIndex = -1;
        let cellStartCol = 0;
        let cellEndCol = 0;
        for (let i = 0; i < cells.length; i++) {
          const cellLength = cells[i]?.length ?? 0;
          const cellEndColThisIteration = currentOffset + cellLength;
          if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
            targetCellIndex = i;
            cellStartCol = currentOffset;
            cellEndCol = cellEndColThisIteration;
            break;
          }
          currentOffset += cellLength + DELIMITER.length;
        }

        if (targetCellIndex === -1) return;

        const isBackward = inputType === 'deleteContentBackward';
        const caretCol = selStart[1];

        if ((isBackward && caretCol === cellStartCol) || (!isBackward && caretCol === cellEndCol)) {
          evt.preventDefault();
          return;
        }

        evt.preventDefault();
        try {
          ed.ace_callWithAce((aceInstance) => {
            const delStart = isBackward ? [lineNum, caretCol - 1] : [lineNum, caretCol];
            const delEnd   = isBackward ? [lineNum, caretCol]     : [lineNum, caretCol + 1];
            aceInstance.ace_performDocumentReplaceRange(delStart, delEnd, '');

            const repAfter = aceInstance.ace_getRep();
            ed.ep_data_tables_applyMeta(
              lineNum,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfter,
              ed,
              null,
              docManager
            );

            const newCaretCol = isBackward ? caretCol - 1 : caretCol;
            const newCaretPos = [lineNum, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            aceInstance.ace_fastIncorp(10);

            if (ed.ep_data_tables_editor && ed.ep_data_tables_editor.ep_data_tables_last_clicked && ed.ep_data_tables_editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
              const newRelativePos = newCaretCol - cellStartCol;
              ed.ep_data_tables_editor.ep_data_tables_last_clicked = {
                lineNum: lineNum,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos,
              };
            }
          }, 'tableCollapsedDeleteHandler', true);
        } catch (error) {
          console.error(`${deleteLogPrefix} ERROR handling collapsed delete:`, error);
        }
        return;
      }

      if (isCollapsed) {
        return;
      }

      if (selStart[0] !== selEnd[0]) {
       // log(`${deleteLogPrefix} WARNING: Selection spans multiple lines. Preventing delete to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${deleteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (lineAttrString) {
        try {
          tableMetadata = JSON.parse(lineAttrString);
        } catch {}
      }

      if (!tableMetadata) {
        tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      }

      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
       // log(`${deleteLogPrefix} Line ${lineNum} is NOT a recognised table line. Allowing default delete.`);
        return;
      }

     // log(`${deleteLogPrefix} Line ${lineNum} IS a table line. Metadata:`, tableMetadata);

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;

      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;

        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      /* allow "…cell content + delimiter" selections */
      const wouldClampStart = targetCellIndex > 0 && selStart[1] === cellStartCol - DELIMITER.length;
      const wouldClampEnd = targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length;

      console.log(`[ep_data_tables:beforeinput-delete] Delete selection analysis:`, {
        targetCellIndex,
        selStartCol: selStart[1],
        selEndCol: selEnd[1],
        cellStartCol,
        cellEndCol,
        delimiterLength: DELIMITER.length,
        expectedLeadingDelimiterPos: cellStartCol - DELIMITER.length,
        expectedTrailingDelimiterPos: cellEndCol + DELIMITER.length,
        wouldClampStart,
        wouldClampEnd
      });

      if (wouldClampStart) {
        console.log(`[ep_data_tables:beforeinput-delete] CLAMPING delete selection start from ${selStart[1]} to ${cellStartCol}`);
        selStart[1] = cellStartCol;
      }

      if (wouldClampEnd) {
        console.log(`[ep_data_tables:beforeinput-delete] CLAMPING delete selection end from ${selEnd[1]} to ${cellEndCol}`);
        selEnd[1] = cellEndCol;
      }

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
       // log(`${deleteLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing delete to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${deleteLogPrefix} Selection is entirely within cell ${targetCellIndex}. Intercepting delete to preserve table structure.`);
      evt.preventDefault();

      try {
       // log(`${deleteLogPrefix} Performing deletion via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
          const callAceLogPrefix = `${deleteLogPrefix}[ace_callWithAceOps]`;
         // log(`${callAceLogPrefix} Entered ace_callWithAce for delete operations. selStart:`, selStart, `selEnd:`, selEnd);

         // log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to delete selected text.`);
          aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, '');
         // log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

          const repAfterDeletion = aceInstance.ace_getRep();
          const lineTextAfterDeletion = repAfterDeletion.lines.atIndex(lineNum).text;
          const cellsAfterDeletion = lineTextAfterDeletion.split(DELIMITER);
          const cellTextAfterDeletion = cellsAfterDeletion[targetCellIndex] || '';

          if (cellTextAfterDeletion.length === 0) {
           // log(`${callAceLogPrefix} Cell ${targetCellIndex} became empty after delete – inserting single space to preserve structure.`);
            const insertPos = [lineNum, selStart[1]];
            aceInstance.ace_performDocumentReplaceRange(insertPos, insertPos, ' ');

            const attrStart = insertPos;
            const attrEnd   = [insertPos[0], insertPos[1] + 1];
            aceInstance.ace_performDocumentApplyAttributesToRange(
              attrStart, attrEnd, [[ATTR_CELL, String(targetCellIndex)]],
            );
          }

         // log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
          const repAfterDelete = aceInstance.ace_getRep();
         // log(`${callAceLogPrefix} Fetched rep after delete for applyMeta. Line ${lineNum} text now: "${repAfterDelete.lines.atIndex(lineNum).text}"`);

          ed.ep_data_tables_applyMeta(
            lineNum,
            tableMetadata.tblId,
            tableMetadata.row,
            tableMetadata.cols,
            repAfterDelete,
            ed,
            null,
            docManager
          );
         // log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

          const newCaretAbsoluteCol = (cellTextAfterDeletion.length === 0) ? selStart[1] + 1 : selStart[1];
          const newCaretPos = [lineNum, newCaretAbsoluteCol];
         // log(`${callAceLogPrefix} Setting caret position to: [${newCaretPos}].`);
          aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
         // log(`${callAceLogPrefix} Selection change successful.`);

         // log(`${callAceLogPrefix} Delete operations within ace_callWithAce completed successfully.`);
        }, 'tableDeleteTextOperations', true);

       // log(`${deleteLogPrefix} Delete operation completed successfully.`);
      } catch (error) {
        console.error(`${deleteLogPrefix} ERROR during delete operation:`, error);
       // log(`${deleteLogPrefix} Delete operation failed. Error details:`, { message: error.message, stack: error.stack });
      }
    });

    $inner.on('beforeinput', (evt) => {
      const insertLogPrefix = '[ep_data_tables:beforeinputInsertHandler]';
      const inputType = evt.originalEvent && evt.originalEvent.inputType || '';

      if (!inputType || !inputType.startsWith('insert')) return;

      if ((evt && evt._epDataTablesHandled) || (evt.originalEvent && evt.originalEvent._epDataTablesHandled)) return;

      if (!isAndroidUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
      }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;
      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;
        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      if (inputType === 'insertParagraph' || inputType === 'insertLineBreak') {
        evt.preventDefault();
        evt.stopPropagation();
        if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();
        setTimeout(() => {
          try {
            ed.ace_callWithAce((aceInstance) => {
              aceInstance.ace_fastIncorp(10);
              const freshRep = aceInstance.ace_getRep();
              const freshSelStart = freshRep.selStart;
              const freshSelEnd = freshRep.selEnd;
              aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, ' ');

              const afterRep = aceInstance.ace_getRep();
              const maxLen = Math.max(0, afterRep.lines.atIndex(lineNum)?.text?.length || 0);
              const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
              const endCol = Math.min(startCol + 1, maxLen);
              if (endCol > startCol) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [lineNum, startCol], [lineNum, endCol], [[ATTR_CELL, String(targetCellIndex)]]
                );
              }

              ed.ep_data_tables_applyMeta(
                lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols,
                afterRep, ed, null, docManager
              );

              const newCaretPos = [lineNum, endCol];
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              aceInstance.ace_fastIncorp(10);
            }, 'iosSoftBreakToSpace', true);
          } catch (e) {
            console.error(`${autoLogPrefix} ERROR replacing soft break:`, e);
          }
        }, 0);
        return;
      }

      if (targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }

      if (inputType === 'insertParagraph' || inputType === 'insertLineBreak') {
        evt.preventDefault();
        try {
          navigateToCellBelow(lineNum, targetCellIndex, tableMetadata, ed, docManager);
        } catch (e) { console.error(`${insertLogPrefix} Error navigating on line break:`, e); }
        return;
      }

      if (suppressBeforeInputInsertTextDuringComposition && inputType === 'insertText') {
        evt.preventDefault();
        return;
      }

      if (suppressNextBeforeInputInsertTextOnce && inputType === 'insertText') {
        suppressNextBeforeInputInsertTextOnce = false;
        evt.preventDefault();
        return;
      }

      if (isAndroidChromeComposition) return;

      const rawData = evt.originalEvent && typeof evt.originalEvent.data === 'string' ? evt.originalEvent.data : '';
      if (!rawData) return;

      let insertedText = normalizeSoftWhitespace(
        rawData
          .replace(new RegExp(DELIMITER, 'g'), ' ')
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
          .replace(/\s+/g, ' ')
      );

      if (insertedText.length === 0) {
        evt.preventDefault();
        return;
      }

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        setTimeout(() => {
          ed.ace_callWithAce((aceInstance) => {
            aceInstance.ace_fastIncorp(10);
            const freshRep = aceInstance.ace_getRep();
            const freshSelStart = freshRep.selStart;
            const freshSelEnd = freshRep.selEnd;

            aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

            const repAfterReplace = aceInstance.ace_getRep();
            const freshLineIndex = freshSelStart[0];
            const freshLineEntry = repAfterReplace.lines.atIndex(freshLineIndex);
            const maxLen = Math.max(0, (freshLineEntry && freshLineEntry.text) ? freshLineEntry.text.length : 0);
            const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
            const endColRaw = startCol + insertedText.length;
            const endCol = Math.min(endColRaw, maxLen);
            if (endCol > startCol) {
              aceInstance.ace_performDocumentApplyAttributesToRange(
                [freshLineIndex, startCol], [freshLineIndex, endCol], [[ATTR_CELL, String(targetCellIndex)]]
              );
            }

            ed.ep_data_tables_applyMeta(
              freshLineIndex,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );

            const newCaretCol = endCol;
            const newCaretPos = [freshLineIndex, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            aceInstance.ace_fastIncorp(10);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
              const freshLineText = (freshLineEntry && freshLineEntry.text) || '';
              const freshCells = freshLineText.split(DELIMITER);
              let freshOffset = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
              }
              const newRelativePos = newCaretCol - freshOffset;
              editor.ep_data_tables_last_clicked = {
                lineNum: freshLineIndex,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos,
              };
            }
          }, 'tableInsertTextOperations', true);
        }, 0);
      } catch (error) {
        console.error(`${insertLogPrefix} ERROR during insert handling:`, error);
      }
    });

    $inner.on('beforeinput', (evt) => {
      const genericLogPrefix = '[ep_data_tables:beforeinputInsertTextGeneric]';
      const inputType = (evt.originalEvent && evt.originalEvent.inputType) || '';
      if (!inputType || !inputType.startsWith('insert')) return;

      if (evt._epDataTablesNormalized || (evt.originalEvent && evt.originalEvent._epDataTablesNormalized)) return;
      if (isAndroidUA() || isIOSUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
      }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') return;
      console.debug(`${genericLogPrefix} event`, { inputType, data: evt.originalEvent && evt.originalEvent.data });

      const rawData = evt.originalEvent && typeof evt.originalEvent.data === 'string' ? evt.originalEvent.data : ' ';

      const insertedText = normalizeSoftWhitespace(
        rawData
          .replace(/[\u00A0\r\n\t]/g, ' ') // NBSP sanitized back to space for stability
          .replace(new RegExp(DELIMITER, 'g'), ' ')
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
          .replace(/\s+/g, ' ')
      );

      if (!insertedText) { evt.preventDefault(); return; }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;
      for (let i = 0; i < cells.length; i++) {
        const len = cells[i]?.length ?? 0;
        const end = currentOffset + len;
        if (selStart[1] >= currentOffset && selStart[1] <= end) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = end;
          break;
        }
        currentOffset += len + DELIMITER.length;
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) { evt.preventDefault(); console.debug(`${genericLogPrefix} abort: selection outside cell`, { selStart, selEnd, cellStartCol, cellEndCol }); return; }

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        ed.ace_callWithAce((ace) => {
          ace.ace_fastIncorp(10);
          const freshRep = ace.ace_getRep();
          const freshSelStart = freshRep.selStart;
          const freshSelEnd = freshRep.selEnd;

          ace.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

          const afterRep = ace.ace_getRep();
          const lineEntry = afterRep.lines.atIndex(lineNum);
          const maxLen = lineEntry ? lineEntry.text.length : 0;
          const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
          const endCol = Math.min(startCol + insertedText.length, maxLen);
          if (endCol > startCol) {
            ace.ace_performDocumentApplyAttributesToRange([lineNum, startCol], [lineNum, endCol], [[ATTR_CELL, String(targetCellIndex)]]);
          }

          ed.ep_data_tables_applyMeta(lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols, afterRep, ed, null, docManager);

          const newCaretPos = [lineNum, endCol];
          ace.ace_performSelectionChange(newCaretPos, newCaretPos, false);
        }, 'tableGenericInsertText', true);
      } catch (e) {
        console.error(`${genericLogPrefix} ERROR handling generic insertText:`, e);
      }
    });

    $inner.on('beforeinput', (evt) => {
      const autoLogPrefix = '[ep_data_tables:beforeinputAutoReplaceHandler]';
      const inputType = (evt.originalEvent && evt.originalEvent.inputType) || '';

      if ((evt && evt._epDataTablesHandled) || (evt.originalEvent && evt.originalEvent._epDataTablesHandled)) return;

      if (!isIOSUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      const dataStr = (evt.originalEvent && typeof evt.originalEvent.data === 'string') ? evt.originalEvent.data : '';
      const hasSelection = !(selStart[0] === selEnd[0] && selStart[1] === selEnd[1]);
      const looksLikeIOSAutoReplace = inputType === 'insertText' && dataStr.length > 1;
      const insertTextNull = inputType === 'insertText' && dataStr === '' && !hasSelection;
      const shouldHandle = INPUTTYPE_REPLACEMENT_TYPES.has(inputType) || looksLikeIOSAutoReplace || (inputType === 'insertText' && (hasSelection || insertTextNull));
      if (!shouldHandle) return;

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) {
        try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {}
      }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;
      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;
        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      if (targetCellIndex !== -1 && selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }

      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
        evt.preventDefault();
        return;
      }

      let insertedText = dataStr;
      if (!insertedText) {
        if (insertTextNull) {
          evt.preventDefault();
          evt.stopPropagation();
          if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

          setTimeout(() => {
            try {
              ed.ace_callWithAce((aceInstance) => {
                aceInstance.ace_fastIncorp(10);
                const freshRep = aceInstance.ace_getRep();
                const freshSelStart = freshRep.selStart;
                const freshSelEnd = freshRep.selEnd;
                aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, ' ');

                const afterRep = aceInstance.ace_getRep();
                const maxLen = Math.max(0, afterRep.lines.atIndex(lineNum)?.text?.length || 0);
                const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
                const endCol = Math.min(startCol + 1, maxLen);
                if (endCol > startCol) {
                  aceInstance.ace_performDocumentApplyAttributesToRange(
                    [lineNum, startCol], [lineNum, endCol], [[ATTR_CELL, String(targetCellIndex)]]
                  );
                }

                ed.ep_data_tables_applyMeta(
                  lineNum, tableMetadata.tblId, tableMetadata.row, tableMetadata.cols,
                  afterRep, ed, null, docManager
                );

                const newCaretPos = [lineNum, endCol];
                aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
                aceInstance.ace_fastIncorp(10);
              }, 'iosPredictiveCommit', true);
            } catch (e) {
              console.error(`${autoLogPrefix} ERROR fixing predictive commit:`, e);
            }
          }, 0);
          return;
        } else {
          if (INPUTTYPE_REPLACEMENT_TYPES.has(inputType) || hasSelection) {
            evt.preventDefault();
            evt.stopPropagation();
            if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();
          }
          return;
        }
      }

      insertedText = normalizeSoftWhitespace(
        insertedText
          .replace(new RegExp(DELIMITER, 'g'), ' ')
          .replace(/[\u200B\u200C\u200D\uFEFF]/g, '')
      );

      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
        setTimeout(() => {
          ed.ace_callWithAce((aceInstance) => {
            aceInstance.ace_fastIncorp(10);
            const freshRep = aceInstance.ace_getRep();
            const freshSelStart = freshRep.selStart;
            const freshSelEnd = freshRep.selEnd;

            aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

            const repAfterReplace = aceInstance.ace_getRep();
            const freshLineIndex = freshSelStart[0];
            const freshLineEntry = repAfterReplace.lines.atIndex(freshLineIndex);
            const maxLen = Math.max(0, (freshLineEntry && freshLineEntry.text) ? freshLineEntry.text.length : 0);
            const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
            const endColRaw = startCol + insertedText.length;
            const endCol = Math.min(endColRaw, maxLen);
            if (endCol > startCol) {
              aceInstance.ace_performDocumentApplyAttributesToRange(
                [freshLineIndex, startCol], [freshLineIndex, endCol], [[ATTR_CELL, String(targetCellIndex)]]
              );
            }

            ed.ep_data_tables_applyMeta(
              freshLineIndex,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );

            const newCaretCol = endCol;
            const newCaretPos = [freshLineIndex, newCaretCol];
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
            aceInstance.ace_fastIncorp(10);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
              const freshLineText = (freshLineEntry && freshLineEntry.text) || '';
              const freshCells = freshLineText.split(DELIMITER);
              let freshOffset = 0;
              for (let i = 0; i < targetCellIndex; i++) {
                freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
              }
              const newRelativePos = newCaretCol - freshOffset;
              editor.ep_data_tables_last_clicked = {
                lineNum: freshLineIndex,
                tblId: tableMetadata.tblId,
                cellIndex: targetCellIndex,
                relativePos: newRelativePos < 0 ? 0 : newRelativePos,
              };
            }
          }, 'tableAutoReplaceTextOperations', true);
        }, 0);
      } catch (error) {
        console.error(`${autoLogPrefix} ERROR during auto-replace handling:`, error);
      }
    });

    $inner.on('compositionstart', (evt) => {
      if (!isAndroidUA()) return;
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const lineNum = rep.selStart[0];
      let meta = null; let s = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      if (s) { try { meta = JSON.parse(s); } catch (_) {} }
      if (!meta) meta = getTableLineMetadata(lineNum, ed, docManager);
      if (!meta || typeof meta.cols !== 'number') return;
      isAndroidChromeComposition = true;
      handledCurrentComposition = false;
      suppressBeforeInputInsertTextDuringComposition = false;
    });
    $inner.on('compositionupdate', (evt) => {
      const compLogPrefix = '[ep_data_tables:compositionHandler]';

      if (!isAndroidUA()) return;

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) return;
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];

      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;
      if (lineAttrString) { try { tableMetadata = JSON.parse(lineAttrString); } catch (_) {} }
      if (!tableMetadata) tableMetadata = getTableLineMetadata(lineNum, ed, docManager);
      if (!tableMetadata || typeof tableMetadata.cols !== 'number') return;

      const d = evt.originalEvent && typeof evt.originalEvent.data === 'string' ? evt.originalEvent.data : '';
      if (evt.type === 'compositionupdate') {
        const isWhitespaceOnly = d && normalizeSoftWhitespace(d).trim() === '';
        if (!isWhitespaceOnly) return;

        const lineText = rep.lines.atIndex(lineNum)?.text || '';
        const cells = lineText.split(DELIMITER);
        let currentOffset = 0;
        let targetCellIndex = -1;
        let cellStartCol = 0;
        let cellEndCol = 0;
        for (let i = 0; i < cells.length; i++) {
          const cellLength = cells[i]?.length ?? 0;
          const cellEndColThisIteration = currentOffset + cellLength;
          if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
            targetCellIndex = i;
            cellStartCol = currentOffset;
            cellEndCol = cellEndColThisIteration;
            break;
          }
          currentOffset += cellLength + DELIMITER.length;
        }
        if (targetCellIndex === -1 || selEnd[1] > cellEndCol) return;

        evt.preventDefault();
        evt.stopPropagation();
        if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

        let insertedText = d
          .replace(/[\u00A0\r\n\t]/g, ' ')
          .replace(/\s+/g, ' ');
        if (insertedText.length === 0) insertedText = ' ';

        try {
          setTimeout(() => {
            ed.ace_callWithAce((aceInstance) => {
              aceInstance.ace_fastIncorp(10);
              const freshRep = aceInstance.ace_getRep();
              const freshSelStart = freshRep.selStart;
              const freshSelEnd = freshRep.selEnd;
              aceInstance.ace_performDocumentReplaceRange(freshSelStart, freshSelEnd, insertedText);

              const repAfterReplace = aceInstance.ace_getRep();
              const freshLineIndex = freshSelStart[0];
              const freshLineEntry = repAfterReplace.lines.atIndex(freshLineIndex);
              const maxLen = Math.max(0, (freshLineEntry && freshLineEntry.text) ? freshLineEntry.text.length : 0);
              const startCol = Math.min(Math.max(freshSelStart[1], 0), maxLen);
              const endColRaw = startCol + insertedText.length;
              const endCol = Math.min(endColRaw, maxLen);
              if (endCol > startCol) {
                aceInstance.ace_performDocumentApplyAttributesToRange(
                  [freshLineIndex, startCol], [freshLineIndex, endCol], [[ATTR_CELL, String(targetCellIndex)]]
                );
              }

              ed.ep_data_tables_applyMeta(
                freshLineIndex,
                tableMetadata.tblId,
                tableMetadata.row,
                tableMetadata.cols,
                repAfterReplace,
                ed,
                null,
                docManager
              );

              const newCaretCol = endCol;
              const newCaretPos = [freshLineIndex, newCaretCol];
              aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
              aceInstance.ace_fastIncorp(10);
              if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
                const freshLineText = (freshLineEntry && freshLineEntry.text) || '';
                const freshCells = freshLineText.split(DELIMITER);
                let freshOffset = 0;
                for (let i = 0; i < targetCellIndex; i++) {
                  freshOffset += (freshCells[i]?.length ?? 0) + DELIMITER.length;
                }
                const newRelativePos = newCaretCol - freshOffset;
                editor.ep_data_tables_last_clicked = {
                  lineNum: freshLineIndex,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
                };
              }
            }, 'tableCompositionSpaceInsert', true);
          }, 0);
          suppressBeforeInputInsertTextDuringComposition = true;
        } catch (error) {
          console.error(`${compLogPrefix} ERROR inserting space during composition:`, error);
        }
      }
    });

    $inner.on('compositionend', () => {
      if (isAndroidChromeComposition) {
        isAndroidChromeComposition = false;
        handledCurrentComposition = false;
        suppressBeforeInputInsertTextDuringComposition = false;
      }
    });

   // log(`${callWithAceLogPrefix} Attaching drag and drop event listeners to $inner (inner iframe body).`);

    $inner.on('drop', (evt) => {
      const dropLogPrefix = '[ep_data_tables:dropHandler]';
     // log(`${dropLogPrefix} DROP EVENT TRIGGERED. Event object:`, evt);

      const targetEl = evt.target;
      if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('table.dataTable')) {
        evt.preventDefault();
        evt.stopPropagation();
        if (evt.originalEvent && evt.originalEvent.dataTransfer) {
          try { evt.originalEvent.dataTransfer.dropEffect = 'none'; } catch (_) {}
        }
        console.warn('[ep_data_tables] Drop prevented on table to protect structure.');
        return;
      }

     // log(`${dropLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
       // log(`${dropLogPrefix} WARNING: Could not get representation or selection. Allowing default drop.`);
        return;
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];
     // log(`${dropLogPrefix} Current line number: ${lineNum}.`);

     // log(`${dropLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let isTableLine = !!lineAttrString;

      if (!isTableLine) {
        const metadataFallback = getTableLineMetadata(lineNum, ed, docManager);
        isTableLine = !!metadataFallback;
      }

      if (isTableLine) {
     // log(`${dropLogPrefix} Line ${lineNum} IS a table line. Preventing drop to protect table structure.`);
      evt.preventDefault();
      evt.stopPropagation();
      console.warn('[ep_data_tables] Drop operation prevented to protect table structure. Please use copy/paste within table cells.');
      }
    });

    $inner.on('dragover', (evt) => {
      const dragLogPrefix = '[ep_data_tables:dragoverHandler]';

      const targetEl = evt.target;
      if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('table.dataTable')) {
        if (evt.originalEvent && evt.originalEvent.dataTransfer) {
          try { evt.originalEvent.dataTransfer.dropEffect = 'none'; } catch (_) {}
        }
        evt.preventDefault();
        evt.stopPropagation();
        return;
      }

      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
        return;
      }

      const selStart = rep.selStart;
      const lineNum = selStart[0];

     // log(`${dragLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let isTableLine = !!lineAttrString;

      if (!isTableLine) {
        isTableLine = !!getTableLineMetadata(lineNum, ed, docManager);
      }

      if (isTableLine) {
       // log(`${dragLogPrefix} Preventing dragover on table line ${lineNum} to control drop handling.`);
        evt.preventDefault();
      }
    });

    $inner.on('dragenter', (evt) => {
      const targetEl = evt.target;
      if (targetEl && typeof targetEl.closest === 'function' && targetEl.closest('table.dataTable')) {
        if (evt.originalEvent && evt.originalEvent.dataTransfer) {
          try { evt.originalEvent.dataTransfer.dropEffect = 'none'; } catch (_) {}
        }
        evt.preventDefault();
        evt.stopPropagation();
      }
    });

   // log(`${callWithAceLogPrefix} Attaching paste event listener to $inner (inner iframe body).`);
    $inner.on('paste', (evt) => {
      const pasteLogPrefix = '[ep_data_tables:pasteHandler]';
     // log(`${pasteLogPrefix} PASTE EVENT TRIGGERED. Event object:`, evt);

     // log(`${pasteLogPrefix} Getting current editor representation (rep).`);
      const rep = ed.ace_getRep();
      if (!rep || !rep.selStart) {
       // log(`${pasteLogPrefix} WARNING: Could not get representation or selection. Allowing default paste.`);
        console.warn(`${pasteLogPrefix} Could not get rep or selStart.`);
        return;
      }
     // log(`${pasteLogPrefix} Rep obtained. selStart:`, rep.selStart, `selEnd:`, rep.selEnd);
      const selStart = rep.selStart;
      const selEnd = rep.selEnd;
      const lineNum = selStart[0];
     // log(`${pasteLogPrefix} Current line number: ${lineNum}. Column start: ${selStart[1]}, Column end: ${selEnd[1]}.`);

      if (selStart[0] !== selEnd[0]) {
       // log(`${pasteLogPrefix} WARNING: Selection spans multiple lines. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${pasteLogPrefix} Checking if line ${lineNum} is a table line by fetching '${ATTR_TABLE_JSON}' attribute.`);
      let lineAttrString = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
      let tableMetadata = null;

      if (!lineAttrString) {
       // log(`${pasteLogPrefix} No '${ATTR_TABLE_JSON}' attribute found. Checking if this is a block-styled table row via DOM reconstruction.`);
        const fallbackMeta = getTableLineMetadata(lineNum, ed, docManager);
        if (fallbackMeta) {
          tableMetadata = fallbackMeta;
          lineAttrString = JSON.stringify(fallbackMeta);
         // log(`${pasteLogPrefix} Block-styled table row detected. Reconstructed metadata:`, fallbackMeta);
        }
      }

      if (!lineAttrString) {
       // log(`${pasteLogPrefix} Line ${lineNum} is NOT a table line (no '${ATTR_TABLE_JSON}' attribute found and no DOM reconstruction possible). Allowing default paste.`);
        return;
      }
     // log(`${pasteLogPrefix} Line ${lineNum} IS a table line. Attribute string: "${lineAttrString}".`);

      try {
       // log(`${pasteLogPrefix} Parsing table metadata from attribute string.`);
        if (!tableMetadata) {
          tableMetadata = JSON.parse(lineAttrString);
        }
       // log(`${pasteLogPrefix} Parsed table metadata:`, tableMetadata);
        if (!tableMetadata || typeof tableMetadata.cols !== 'number' || typeof tableMetadata.tblId === 'undefined' || typeof tableMetadata.row === 'undefined') {
         // log(`${pasteLogPrefix} WARNING: Invalid or incomplete table metadata on line ${lineNum}. Allowing default paste. Metadata:`, tableMetadata);
          console.warn(`${pasteLogPrefix} Invalid table metadata for line ${lineNum}.`);
          return;
        }
       // log(`${pasteLogPrefix} Table metadata validated successfully: tblId=${tableMetadata.tblId}, row=${tableMetadata.row}, cols=${tableMetadata.cols}.`);
      } catch(e) {
        console.error(`${pasteLogPrefix} ERROR parsing table metadata for line ${lineNum}:`, e);
       // log(`${pasteLogPrefix} Metadata parse error. Allowing default paste. Error details:`, { message: e.message, stack: e.stack });
        return;
      }

      const lineText = rep.lines.atIndex(lineNum)?.text || '';
      const cells = lineText.split(DELIMITER);
      let currentOffset = 0;
      let targetCellIndex = -1;
      let cellStartCol = 0;
      let cellEndCol = 0;

      for (let i = 0; i < cells.length; i++) {
        const cellLength = cells[i]?.length ?? 0;
        const cellEndColThisIteration = currentOffset + cellLength;

        if (selStart[1] >= currentOffset && selStart[1] <= cellEndColThisIteration) {
          targetCellIndex = i;
          cellStartCol = currentOffset;
          cellEndCol = cellEndColThisIteration;
          break;
        }
        currentOffset += cellLength + DELIMITER.length;
      }

      /* allow "…cell content + delimiter" selections */
      if (targetCellIndex !== -1 &&
          selEnd[1] === cellEndCol + DELIMITER.length) {
        selEnd[1] = cellEndCol;
      }
      if (targetCellIndex === -1 || selEnd[1] > cellEndCol) {
       // log(`${pasteLogPrefix} WARNING: Selection spans cell boundaries or is outside cells. Preventing paste to protect table structure.`);
        evt.preventDefault();
        return;
      }

     // log(`${pasteLogPrefix} Accessing clipboard data.`);
      const clipboardData = evt.originalEvent.clipboardData || window.clipboardData;
      if (!clipboardData) {
       // log(`${pasteLogPrefix} WARNING: No clipboard data found. Allowing default paste.`);
        return;
      }
     // log(`${pasteLogPrefix} Clipboard data object obtained:`, clipboardData);

      const types = clipboardData.types || [];
      if (types.includes('text/html') && clipboardData.getData('text/html')) {
       // log(`${pasteLogPrefix} Detected text/html in clipboard – deferring to other plugins and default paste.`);
        return;
      }

     // log(`${pasteLogPrefix} Getting 'text/plain' from clipboard.`);
      const pastedTextRaw = clipboardData.getData('text/plain');
     // log(`${pasteLogPrefix} Pasted text raw: "${pastedTextRaw}" (Type: ${typeof pastedTextRaw})`);

      let pastedText = pastedTextRaw
        .replace(/(\r\n|\n|\r)/gm, " ")
        .replace(new RegExp(DELIMITER, 'g'), ' ')
        .replace(/\t/g, " ")
        .replace(/\s+/g, " ")
        .trim();

     // log(`${pasteLogPrefix} Pasted text after sanitization: "${pastedText}"`);

      if (typeof pastedText !== 'string' || pastedText.length === 0) {
       // log(`${pasteLogPrefix} No plain text in clipboard or text is empty (after sanitization). Allowing default paste.`);
        const types = clipboardData.types;
       // log(`${pasteLogPrefix} Clipboard types available:`, types);
        if (types && types.includes('text/html')) {
           // log(`${pasteLogPrefix} Clipboard also contains HTML:`, clipboardData.getData('text/html'));
        }
        return;
      }
     // log(`${pasteLogPrefix} Plain text obtained from clipboard: "${pastedText}". Length: ${pastedText.length}.`);

      const currentCellText = cells[targetCellIndex] || '';
      const selectionLength = selEnd[1] - selStart[1];
      const newCellLength = currentCellText.length - selectionLength + pastedText.length;

      const MAX_CELL_LENGTH = 8000;
      if (newCellLength > MAX_CELL_LENGTH) {
       // log(`${pasteLogPrefix} WARNING: Paste would exceed maximum cell length (${newCellLength} > ${MAX_CELL_LENGTH}). Truncating paste.`);
        const truncatedPaste = pastedText.substring(0, MAX_CELL_LENGTH - (currentCellText.length - selectionLength));
        if (truncatedPaste.length === 0) {
         // log(`${pasteLogPrefix} Paste would be completely truncated. Preventing paste.`);
          evt.preventDefault();
          return;
        }
       // log(`${pasteLogPrefix} Using truncated paste: "${truncatedPaste}"`);
        pastedText = truncatedPaste;
      }

     // log(`${pasteLogPrefix} INTERCEPTING paste of plain text into table line ${lineNum}. PREVENTING DEFAULT browser action.`);
      evt.preventDefault();
      evt.stopPropagation();
      if (typeof evt.stopImmediatePropagation === 'function') evt.stopImmediatePropagation();

      try {
       // log(`${pasteLogPrefix} Preparing to perform paste operations via ed.ace_callWithAce.`);
        ed.ace_callWithAce((aceInstance) => {
            const callAceLogPrefix = `${pasteLogPrefix}[ace_callWithAceOps]`;
           // log(`${callAceLogPrefix} Entered ace_callWithAce for paste operations. selStart:`, selStart, `selEnd:`, selEnd);

           // log(`${callAceLogPrefix} Original line text from initial rep: "${rep.lines.atIndex(lineNum).text}". SelStartCol: ${selStart[1]}, SelEndCol: ${selEnd[1]}.`);

           // log(`${callAceLogPrefix} Calling aceInstance.ace_performDocumentReplaceRange to insert text: "${pastedText}".`);
            aceInstance.ace_performDocumentReplaceRange(selStart, selEnd, pastedText);
           // log(`${callAceLogPrefix} ace_performDocumentReplaceRange successful.`);

           // log(`${callAceLogPrefix} Preparing to re-apply tbljson attribute to line ${lineNum}.`);
            const repAfterReplace = aceInstance.ace_getRep();
           // log(`${callAceLogPrefix} Fetched rep after replace for applyMeta. Line ${lineNum} text now: "${repAfterReplace.lines.atIndex(lineNum).text}"`);

            ed.ep_data_tables_applyMeta(
              lineNum,
              tableMetadata.tblId,
              tableMetadata.row,
              tableMetadata.cols,
              repAfterReplace,
              ed,
              null,
              docManager
            );
           // log(`${callAceLogPrefix} tbljson attribute re-applied successfully via ep_data_tables_applyMeta.`);

            const newCaretCol = selStart[1] + pastedText.length;
            const newCaretPos = [lineNum, newCaretCol];
           // log(`${callAceLogPrefix} New calculated caret position: [${newCaretPos}]. Setting selection.`);
            aceInstance.ace_performSelectionChange(newCaretPos, newCaretPos, false);
           // log(`${callAceLogPrefix} Selection change successful.`);

           // log(`${callAceLogPrefix} Requesting fastIncorp(10) for sync.`);
            aceInstance.ace_fastIncorp(10);
           // log(`${callAceLogPrefix} fastIncorp requested.`);

            if (editor && editor.ep_data_tables_last_clicked && editor.ep_data_tables_last_clicked.tblId === tableMetadata.tblId) {
               const newRelativePos = newCaretCol - cellStartCol;
               editor.ep_data_tables_last_clicked = {
                  lineNum: lineNum,
                  tblId: tableMetadata.tblId,
                  cellIndex: targetCellIndex,
                  relativePos: newRelativePos < 0 ? 0 : newRelativePos,
               };
              // log(`${callAceLogPrefix} Updated stored click/caret info:`, editor.ep_data_tables_last_clicked);
            }

           // log(`${callAceLogPrefix} Paste operations within ace_callWithAce completed successfully.`);
        }, 'tablePasteTextOperations', true);
       // log(`${pasteLogPrefix} ed.ace_callWithAce for paste operations was called.`);

      } catch (error) {
        console.error(`${pasteLogPrefix} CRITICAL ERROR during paste handling operation:`, error);
       // log(`${pasteLogPrefix} Error details:`, { message: error.message, stack: error.stack });
       // log(`${pasteLogPrefix} Paste handling FAILED. END OF HANDLER.`);
      }
    });
   // log(`${callWithAceLogPrefix} Paste event listener attached.`);

   // log(`${callWithAceLogPrefix} Attaching column resize listeners...`);

    const $iframeOuter = $('iframe[name="ace_outer"]');
    const $iframeInner = $iframeOuter.contents().find('iframe[name="ace_inner"]');
    const innerDoc = $iframeInner.contents();
    const outerDoc = $iframeOuter.contents();

   // log(`${callWithAceLogPrefix} Found iframe documents: outer=${outerDoc.length}, inner=${innerDoc.length}`);

    $inner.on('mousedown', '.ep-data_tables-resize-handle', (evt) => {
      const resizeLogPrefix = '[ep_data_tables:resizeMousedown]';
     // log(`${resizeLogPrefix} Resize handle mousedown detected`);

      if (evt.button !== 0) {
       // log(`${resizeLogPrefix} Ignoring non-left mouse button: ${evt.button}`);
        return;
      }

      const target = evt.target;
      const $target = $(target);
      const isImageRelated = $target.closest('.inline-image, .image-placeholder, .image-inner').length > 0;
      const isImageResizeHandle = $target.hasClass('image-resize-handle') || $target.closest('.image-resize-handle').length > 0;

      if (isImageRelated || isImageResizeHandle) {
       // log(`${resizeLogPrefix} Click detected on image-related element or image resize handle, ignoring for table resize`);
        return;
      }

      evt.preventDefault();
      evt.stopPropagation();

      const handle = evt.target;
      const columnIndex = parseInt(handle.getAttribute('data-column'), 10);
      const table = handle.closest('table.dataTable');
      const lineNode = table.closest('div.ace-line');

     // log(`${resizeLogPrefix} Parsed resize target: columnIndex=${columnIndex}, table=${!!table}, lineNode=${!!lineNode}`);

      if (table && lineNode && !isNaN(columnIndex)) {
        const tblId = table.getAttribute('data-tblId');
        const rep = ed.ace_getRep();

        if (!rep || !rep.lines) {
          console.error(`${resizeLogPrefix} Cannot get editor representation`);
          return;
        }

        const lineNum = rep.lines.indexOfKey(lineNode.id);

       // log(`${resizeLogPrefix} Table info: tblId=${tblId}, lineNum=${lineNum}`);

        if (tblId && lineNum !== -1) {
          try {
            const lineAttrString = docManager.getAttributeOnLine(lineNum, 'tbljson');
            if (lineAttrString) {
              const metadata = JSON.parse(lineAttrString);
              if (metadata.tblId === tblId) {
               // log(`${resizeLogPrefix} Starting resize with metadata:`, metadata);
                startColumnResize(table, columnIndex, evt.clientX, metadata, lineNum);
               // log(`${resizeLogPrefix} Started resize for column ${columnIndex}`);

               // log(`${resizeLogPrefix} Global resize state: isResizing=${isResizing}, targetTable=${!!resizeTargetTable}, targetColumn=${resizeTargetColumn}`);
              } else {
               // log(`${resizeLogPrefix} Table ID mismatch: ${metadata.tblId} vs ${tblId}`);
              }
            } else {
             // log(`${resizeLogPrefix} No table metadata found for line ${lineNum}, trying DOM reconstruction...`);

              const rep = ed.ace_getRep();
              if (rep && rep.lines) {
                const lineEntry = rep.lines.atIndex(lineNum);
                if (lineEntry && lineEntry.lineNode) {
                  const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
                  if (tableInDOM) {
                    const domTblId = tableInDOM.getAttribute('data-tblId');
                    const domRow = tableInDOM.getAttribute('data-row');
                    if (domTblId === tblId && domRow !== null) {
                      const domCells = tableInDOM.querySelectorAll('td');
                      if (domCells.length > 0) {
                        const columnWidths = [];
                        domCells.forEach(cell => {
                          const style = cell.getAttribute('style') || '';
                          const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                          if (widthMatch) {
                            columnWidths.push(parseFloat(widthMatch[1]));
                          } else {
                            columnWidths.push(100 / domCells.length);
                          }
                        });

                        const reconstructedMetadata = {
                          tblId: domTblId,
                          row: parseInt(domRow, 10),
                          cols: domCells.length,
                          columnWidths: columnWidths
                        };
                       // log(`${resizeLogPrefix} Reconstructed metadata from DOM:`, reconstructedMetadata);

                        startColumnResize(table, columnIndex, evt.clientX, reconstructedMetadata, lineNum);
                       // log(`${resizeLogPrefix} Started resize for column ${columnIndex} using reconstructed metadata`);

                       // log(`${resizeLogPrefix} Global resize state: isResizing=${isResizing}, targetTable=${!!resizeTargetTable}, targetColumn=${resizeTargetColumn}`);
    } else {
                       // log(`${resizeLogPrefix} DOM table found but no cells detected`);
                      }
                    } else {
                     // log(`${resizeLogPrefix} DOM table found but tblId mismatch or missing row: domTblId=${domTblId}, domRow=${domRow}`);
                    }
                  } else {
                   // log(`${resizeLogPrefix} No table found in DOM for line ${lineNum}`);
                  }
                } else {
                 // log(`${resizeLogPrefix} Could not get line entry or lineNode for line ${lineNum}`);
                }
              } else {
               // log(`${resizeLogPrefix} Could not get rep or rep.lines for DOM reconstruction`);
              }
            }
          } catch (e) {
            console.error(`${resizeLogPrefix} Error getting table metadata:`, e);
          }
        } else {
         // log(`${resizeLogPrefix} Invalid line number (${lineNum}) or table ID (${tblId})`);
        }
      } else {
       // log(`${resizeLogPrefix} Invalid resize target:`, { table: !!table, lineNode: !!lineNode, columnIndex });
      }
    });

    const setupGlobalHandlers = () => {
      const mouseupLogPrefix = '[ep_data_tables:resizeMouseup]';
      const mousemoveLogPrefix = '[ep_data_tables:resizeMousemove]';

      const handleMousemove = (evt) => {
        if (isResizing) {
          evt.preventDefault();
          updateColumnResize(evt.clientX);
        }
      };

      const handleMouseup = (evt) => {
       // log(`${mouseupLogPrefix} Mouseup detected on ${evt.target.tagName || 'unknown'}. isResizing: ${isResizing}`);

        if (isResizing) {
         // log(`${mouseupLogPrefix} Processing resize completion...`);
          evt.preventDefault();
          evt.stopPropagation();

          setTimeout(() => {
           // log(`${mouseupLogPrefix} Executing finishColumnResize after delay...`);
            finishColumnResize(ed, docManager);
           // log(`${mouseupLogPrefix} Resize completion finished.`);
          }, 50);
        } else {
         // log(`${mouseupLogPrefix} Not in resize mode, ignoring mouseup.`);
        }
      };

     // log(`${callWithAceLogPrefix} Attaching global mousemove/mouseup handlers to multiple contexts...`);

      $(document).on('mousemove', handleMousemove);
      $(document).on('mouseup', handleMouseup);
     // log(`${callWithAceLogPrefix} Attached to main document`);

      if (outerDoc.length > 0) {
        outerDoc.on('mousemove', handleMousemove);
        outerDoc.on('mouseup', handleMouseup);
       // log(`${callWithAceLogPrefix} Attached to outer iframe document`);
      }

      if (innerDoc.length > 0) {
        innerDoc.on('mousemove', handleMousemove);
        innerDoc.on('mouseup', handleMouseup);
       // log(`${callWithAceLogPrefix} Attached to inner iframe document`);
      }

      $inner.on('mousemove', handleMousemove);
      $inner.on('mouseup', handleMouseup);
     // log(`${callWithAceLogPrefix} Attached to inner iframe body`);

      const failsafeMouseup = (evt) => {
        if (isResizing) {
         // log(`${mouseupLogPrefix} FAILSAFE: Detected mouse event during resize: ${evt.type}`);
          if (evt.type === 'mouseup' || evt.type === 'mousedown' || evt.type === 'click') {
           // log(`${mouseupLogPrefix} FAILSAFE: Triggering resize completion due to ${evt.type}`);
            setTimeout(() => {
              if (isResizing) {
                finishColumnResize(ed, docManager);
              }
            }, 100);
          }
        }
      };

      document.addEventListener('mouseup', failsafeMouseup, true);
      document.addEventListener('mousedown', failsafeMouseup, true);
      document.addEventListener('click', failsafeMouseup, true);
     // log(`${callWithAceLogPrefix} Attached failsafe event handlers`);

      const preventTableDrag = (evt) => {
        const target = evt.target;
        const inTable = target && typeof target.closest === 'function' && target.closest('table.dataTable');
        if (inTable) {
         // log('[ep_data_tables:dragPrevention] Preventing drag operation originating from inside table');
          evt.preventDefault();
          evt.stopPropagation();
          if (evt.originalEvent && evt.originalEvent.dataTransfer) {
            try { evt.originalEvent.dataTransfer.effectAllowed = 'none'; } catch (_) {}
          }
          return false;
        }
      };

      $inner.on('dragstart', preventTableDrag);
      $inner.on('drag', preventTableDrag);
      $inner.on('dragend', preventTableDrag);
     // log(`${callWithAceLogPrefix} Attached drag prevention handlers to inner body`);

      if (innerDoc.length > 0) {
        innerDoc.on('dragstart', preventTableDrag);
        innerDoc.on('drag', preventTableDrag);
        innerDoc.on('dragend', preventTableDrag);
      }
      if (outerDoc.length > 0) {
        outerDoc.on('dragstart', preventTableDrag);
        outerDoc.on('drag', preventTableDrag);
        outerDoc.on('dragend', preventTableDrag);
      }
      $(document).on('dragstart', preventTableDrag);
      $(document).on('drag', preventTableDrag);
      $(document).on('dragend', preventTableDrag);
    };

    setupGlobalHandlers();

   // log(`${callWithAceLogPrefix} Column resize listeners attached successfully.`);

  }, 'tablePasteAndResizeListeners', true);
 // log(`${logPrefix} ace_callWithAce for listeners setup completed.`);

  function applyTableLineMetadataAttribute (lineNum, tblId, rowIndex, numCols, rep, editorInfo, attributeString = null, documentAttributeManager = null) {
    const funcName = 'applyTableLineMetadataAttribute';
   // log(`${logPrefix}:${funcName}: START - Applying METADATA attribute to line ${lineNum}`, {tblId, rowIndex, numCols});

    let finalMetadata;

    if (attributeString) {
      try {
        const providedMetadata = JSON.parse(attributeString);
        if (providedMetadata.columnWidths && Array.isArray(providedMetadata.columnWidths) && providedMetadata.columnWidths.length === numCols) {
          finalMetadata = providedMetadata;
         // log(`${logPrefix}:${funcName}: Using provided metadata with existing columnWidths`);
        } else {
          finalMetadata = providedMetadata;
         // log(`${logPrefix}:${funcName}: Provided metadata missing columnWidths, attempting DOM extraction`);
           }
         } catch (e) {
       // log(`${logPrefix}:${funcName}: Error parsing provided attributeString, will reconstruct:`, e);
        finalMetadata = null;
      }
    }

    if (!finalMetadata || !finalMetadata.columnWidths) {
      let columnWidths = null;

      try {
        const lineEntry = rep.lines.atIndex(lineNum);
        if (lineEntry && lineEntry.lineNode) {
          const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
          if (tableInDOM) {
            const domTblId = tableInDOM.getAttribute('data-tblId');
            if (domTblId === tblId) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    columnWidths.push(100 / numCols);
                  }
                });
               // log(`${logPrefix}:${funcName}: Extracted column widths from DOM: ${columnWidths.map(w => w.toFixed(1) + '%').join(', ')}`);
              }
            }
          }
             }
           } catch (e) {
       // log(`${logPrefix}:${funcName}: Error extracting column widths from DOM:`, e);
      }

      finalMetadata = finalMetadata || {
        tblId: tblId,
        row: rowIndex,
        cols: numCols
      };

      if (columnWidths && columnWidths.length === numCols) {
        finalMetadata.columnWidths = columnWidths;
      }
    }

    const finalAttributeString = JSON.stringify(finalMetadata);
   // log(`${logPrefix}:${funcName}: Final metadata attribute string: ${finalAttributeString}`);

    try {
       const lineEntry = rep.lines.atIndex(lineNum);
       if (!lineEntry) {
          // log(`${logPrefix}:${funcName}: ERROR - Could not find line entry for line number ${lineNum}`);
           return;
       }
       const lineLength = Math.max(1, lineEntry.text.length);
      // log(`${logPrefix}:${funcName}: Line ${lineNum} text length: ${lineLength}`);

       const attributes = [[ATTR_TABLE_JSON, finalAttributeString]];
       const start = [lineNum, 0];
       const end = [lineNum, lineLength];

      // log(`${logPrefix}:${funcName}: Applying tbljson attribute to range [${start}]-[${end}]`);
       editorInfo.ace_performDocumentApplyAttributesToRange(start, end, attributes);
      // log(`${logPrefix}:${funcName}: Successfully applied tbljson attribute to line ${lineNum}`);

    } catch(e) {
        console.error(`[ep_data_tables] ${logPrefix}:${funcName}: Error applying metadata attribute on line ${lineNum}:`, e);
    }
  }

  /** Insert a fresh rows×cols blank table at the caret */
  ed.ace_createTableViaAttributes = (rows = 2, cols = 2) => {
    const funcName = 'ace_createTableViaAttributes';
   // log(`${funcName}: START - Refactored Phase 4 (Get Selection Fix)`, { rows, cols });
    rows = Math.max(1, rows); cols = Math.max(1, cols);
   // log(`${funcName}: Ensuring minimum 1 row, 1 col.`);

    const tblId   = rand();
   // log(`${funcName}: Generated table ID: ${tblId}`);
    const initialCellContent = ' ';
    const lineTxt = Array.from({ length: cols }).fill(initialCellContent).join(DELIMITER);
   // log(`${funcName}: Constructed initial line text for ${cols} cols: "${lineTxt}"`);
    const block = Array.from({ length: rows }).fill(lineTxt).join('\n') + '\n';
   // log(`${funcName}: Constructed block for ${rows} rows:\n${block}`);

   // log(`${funcName}: Getting current representation and selection...`);
    const currentRepInitial = ed.ace_getRep(); 
    if (!currentRepInitial || !currentRepInitial.selStart || !currentRepInitial.selEnd) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation or selection via ace_getRep(). Aborting.`);
       // log(`${funcName}: END - Error getting initial rep/selection`);
        return;
    }
    const start = currentRepInitial.selStart;
    const end = currentRepInitial.selEnd;
    const initialStartLine = start[0];
   // log(`${funcName}: Current selection from initial rep:`, { start, end });

   // log(`${funcName}: Phase 2 - Inserting text block...`);
    ed.ace_performDocumentReplaceRange(start, end, block);
   // log(`${funcName}: Inserted block of delimited text lines.`);
   // log(`${funcName}: Requesting text sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20);
   // log(`${funcName}: Text sync requested.`);

   // log(`${funcName}: Phase 3 - Applying metadata attributes to ${rows} inserted lines...`);
    const currentRep = ed.ace_getRep();
    if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get updated rep after text insertion. Cannot apply attributes reliably.`);
       // log(`${funcName}: END - Error getting updated rep`);
        return; 
    }
   // log(`${funcName}: Fetched updated rep for attribute application.`);

    for (let r = 0; r < rows; r++) {
      const lineNumToApply = initialStartLine + r;
     // log(`${funcName}: -> Processing row ${r} on line ${lineNumToApply}`);

      const lineEntry = currentRep.lines.atIndex(lineNumToApply);
      if (!lineEntry) {
       // log(`${funcName}: Could not find line entry for ${lineNumToApply}, skipping attribute application.`);
        continue;
      }
      const lineText = lineEntry.text || '';
      const cells = lineText.split(DELIMITER);
      let offset = 0;

      for (let c = 0; c < cols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) {
          const cellStart = [lineNumToApply, offset];
          const cellEnd = [lineNumToApply, offset + cellContent.length];
         // log(`${funcName}: Applying ${ATTR_CELL} attribute to Line ${lineNumToApply} Col ${c} Range ${offset}-${offset + cellContent.length}`);
          ed.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < cols - 1) {
          offset += DELIMITER.length;
        }
      }

      applyTableLineMetadataAttribute(lineNumToApply, tblId, r, cols, currentRep, ed, null, null); 
    }
   // log(`${funcName}: Finished applying metadata attributes.`);
   // log(`${funcName}: Requesting attribute sync (ace_fastIncorp 20)...`);
    ed.ace_fastIncorp(20);
   // log(`${funcName}: Attribute sync requested.`);

   // log(`${funcName}: Phase 4 - Setting final caret position...`);
    const finalCaretLine = initialStartLine + rows;
    const finalCaretPos = [finalCaretLine, 0];
   // log(`${funcName}: Target caret position:`, finalCaretPos);
    try {
      ed.ace_performSelectionChange(finalCaretPos, finalCaretPos, false);
      // log(`${funcName}: Successfully set caret position.`);
    } catch(e) {
       console.error(`[ep_data_tables] ${funcName}: Error setting caret position after table creation:`, e);
      // log(`[ep_data_tables] ${funcName}: Error details:`, { message: e.message, stack: e.stack });
    }

   // log(`${funcName}: END - Refactored Phase 4`);
  };

  ed.ace_doDatatableOptions = (action) => {
    const funcName = 'ace_doDatatableOptions';
   // log(`${funcName}: START - Processing action: ${action}`);

    const editor = ed.ep_data_tables_editor;
    if (!editor) {
      console.error(`[ep_data_tables] ${funcName}: Could not get editor reference.`);
      return;
    }

    const lastClick = editor.ep_data_tables_last_clicked;
    if (!lastClick || !lastClick.tblId) {
     // log(`${funcName}: No table selected. Please click on a table cell first.`);
      console.warn('[ep_data_tables] No table selected. Please click on a table cell first.');
      return;
    }

   // log(`${funcName}: Operating on table ${lastClick.tblId}, clicked line ${lastClick.lineNum}, cell ${lastClick.cellIndex}`);

    try {
      const currentRep = ed.ace_getRep();
      if (!currentRep || !currentRep.lines) {
        console.error(`[ep_data_tables] ${funcName}: Could not get current representation.`);
        return;
      }

      const docManager = ed.ep_data_tables_docManager;
      if (!docManager) {
        console.error(`[ep_data_tables] ${funcName}: Could not get document attribute manager from stored reference.`);
        return;
      }

     // log(`${funcName}: Successfully obtained documentAttributeManager from stored reference.`);

      const tableLines = [];
      const totalLines = currentRep.lines.length();

      for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
        try {
          let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

          if (!lineAttrString) {
            const lineEntry = currentRep.lines.atIndex(lineIndex);
            if (lineEntry && lineEntry.lineNode) {
              const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
              if (tableInDOM) {
                const domTblId = tableInDOM.getAttribute('data-tblId');
                const domRow = tableInDOM.getAttribute('data-row');
                if (domTblId && domRow !== null) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length > 0) {
                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length
                    };
                    lineAttrString = JSON.stringify(reconstructedMetadata);
                   // log(`${funcName}: Reconstructed metadata from DOM for line ${lineIndex}: ${lineAttrString}`);
                  }
                }
              }
            }
          }

          if (lineAttrString) {
            const lineMetadata = JSON.parse(lineAttrString);
            if (lineMetadata.tblId === lastClick.tblId) {
              const lineEntry = currentRep.lines.atIndex(lineIndex);
              if (lineEntry) {
                tableLines.push({
                  lineIndex,
                  row: lineMetadata.row,
                  cols: lineMetadata.cols,
                  lineText: lineEntry.text,
                  metadata: lineMetadata
                });
              }
            }
          }
        } catch (e) {
          continue;
        }
      }

      if (tableLines.length === 0) {
       // log(`${funcName}: No table lines found for table ${lastClick.tblId}`);
        return;
      }

      tableLines.sort((a, b) => a.row - b.row);
     // log(`${funcName}: Found ${tableLines.length} table lines`);

      const numRows = tableLines.length;
      const numCols = tableLines[0].cols;

      let targetRowIndex = -1;

      targetRowIndex = tableLines.findIndex(line => line.lineIndex === lastClick.lineNum);

      if (targetRowIndex === -1) {
       // log(`${funcName}: Direct line number match failed, searching by DOM structure...`);
        const clickedLineEntry = currentRep.lines.atIndex(lastClick.lineNum);
        if (clickedLineEntry && clickedLineEntry.lineNode) {
          const clickedTable = clickedLineEntry.lineNode.querySelector('table.dataTable[data-tblId="' + lastClick.tblId + '"]');
          if (clickedTable) {
            const clickedRowAttr = clickedTable.getAttribute('data-row');
            if (clickedRowAttr !== null) {
              const clickedRowNum = parseInt(clickedRowAttr, 10);
              targetRowIndex = tableLines.findIndex(line => line.row === clickedRowNum);
             // log(`${funcName}: Found target row by DOM attribute matching: row ${clickedRowNum}, index ${targetRowIndex}`);
            }
          }
        }
      }

      if (targetRowIndex === -1) {
       // log(`${funcName}: Warning: Could not find target row, defaulting to row 0`);
        targetRowIndex = 0;
      }

      const targetColIndex = lastClick.cellIndex || 0;

     // log(`${funcName}: Table dimensions: ${numRows} rows x ${numCols} cols. Target: row ${targetRowIndex}, col ${targetColIndex}`);

      let newNumCols = numCols;
      let success = false;

      switch (action) {
        case 'addTblRowA':
         // log(`${funcName}: Inserting row above row ${targetRowIndex}`);
          success = addTableRowAboveWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;

        case 'addTblRowB':
         // log(`${funcName}: Inserting row below row ${targetRowIndex}`);
          success = addTableRowBelowWithText(tableLines, targetRowIndex, numCols, lastClick.tblId, ed, docManager);
          break;

        case 'addTblColL':
         // log(`${funcName}: Inserting column left of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnLeftWithText(tableLines, targetColIndex, ed, docManager);
          break;

        case 'addTblColR':
         // log(`${funcName}: Inserting column right of column ${targetColIndex}`);
          newNumCols = numCols + 1;
          success = addTableColumnRightWithText(tableLines, targetColIndex, ed, docManager);
          break;

        case 'delTblRow':
          const rowConfirmMessage = `Are you sure you want to delete Row ${targetRowIndex + 1} and all content within?`;
          if (!confirm(rowConfirmMessage)) {
           // log(`${funcName}: Row deletion cancelled by user`);
            return;
          }
         // log(`${funcName}: Deleting row ${targetRowIndex}`);
          success = deleteTableRowWithText(tableLines, targetRowIndex, ed, docManager);
          break;

        case 'delTblCol':
          const colConfirmMessage = `Are you sure you want to delete Column ${targetColIndex + 1} and all content within?`;
          if (!confirm(colConfirmMessage)) {
           // log(`${funcName}: Column deletion cancelled by user`);
            return;
          }
         // log(`${funcName}: Deleting column ${targetColIndex}`);
          newNumCols = numCols - 1;
          success = deleteTableColumnWithText(tableLines, targetColIndex, ed, docManager);
          break;

        default:
         // log(`${funcName}: Unknown action: ${action}`);
          return;
      }

      if (!success) {
        console.error(`[ep_data_tables] ${funcName}: Table operation failed for action: ${action}`);
        return;
      }

     // log(`${funcName}: Table operation completed successfully with text and metadata synchronization`);

    } catch (error) {
      console.error(`[ep_data_tables] ${funcName}: Error during table operation:`, error);
     // log(`${funcName}: Error details:`, { message: error.message, stack: error.stack });
    }
  };

  function addTableRowAboveWithText(tableLines, targetRowIndex, numCols, tblId, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      const newLineText = Array.from({ length: numCols }).fill(' ').join(DELIMITER);
      const insertLineIndex = targetLine.lineIndex;

      editorInfo.ace_performDocumentReplaceRange([insertLineIndex, 0], [insertLineIndex, 0], newLineText + '\n');

      const rep = editorInfo.ace_getRep();
      const cells = newLineText.split(DELIMITER);
      let offset = 0;
      for (let c = 0; c < numCols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) {
          const cellStart = [insertLineIndex, offset];
          const cellEnd = [insertLineIndex, offset + cellContent.length];
          editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < numCols - 1) {
          offset += DELIMITER.length;
        }
      }

      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        try {
          const rep = editorInfo.ace_getRep();
          const lineEntry = rep.lines.atIndex(targetLine.lineIndex + 1);
          if (lineEntry && lineEntry.lineNode) {
            const tableInDOM = lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${tblId}"]`);
            if (tableInDOM) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    columnWidths.push(100 / numCols);
                  }
                });
               // log('[ep_data_tables] addTableRowAbove: Extracted column widths from DOM:', columnWidths);
              }
            }
          }
        } catch (e) {
          console.error('[ep_data_tables] addTableRowAbove: Error extracting column widths from DOM:', e);
        }
      }

      for (let i = targetRowIndex; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1;
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };

        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      const newMetadata = { tblId, row: targetLine.metadata.row, cols: numCols, columnWidths };
      applyTableLineMetadataAttribute(insertLineIndex, tblId, targetLine.metadata.row, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding row above with text:', e);
      return false;
    }
  }

  function addTableRowBelowWithText(tableLines, targetRowIndex, numCols, tblId, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];
      const newLineText = Array.from({ length: numCols }).fill(' ').join(DELIMITER);
      const insertLineIndex = targetLine.lineIndex + 1;

      editorInfo.ace_performDocumentReplaceRange([insertLineIndex, 0], [insertLineIndex, 0], newLineText + '\n');

      const rep = editorInfo.ace_getRep();
      const cells = newLineText.split(DELIMITER);
      let offset = 0;
      for (let c = 0; c < numCols; c++) {
        const cellContent = (c < cells.length) ? cells[c] || '' : '';
        if (cellContent.length > 0) {
          const cellStart = [insertLineIndex, offset];
          const cellEnd = [insertLineIndex, offset + cellContent.length];
          editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
        }
        offset += cellContent.length;
        if (c < numCols - 1) {
          offset += DELIMITER.length;
        }
      }

      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        try {
          const rep = editorInfo.ace_getRep();
          const lineEntry = rep.lines.atIndex(targetLine.lineIndex);
          if (lineEntry && lineEntry.lineNode) {
            const tableInDOM = lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${tblId}"]`);
            if (tableInDOM) {
              const domCells = tableInDOM.querySelectorAll('td');
              if (domCells.length === numCols) {
                columnWidths = [];
                domCells.forEach(cell => {
                  const style = cell.getAttribute('style') || '';
                  const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                  if (widthMatch) {
                    columnWidths.push(parseFloat(widthMatch[1]));
                  } else {
                    columnWidths.push(100 / numCols);
                  }
                });
                // log('[ep_data_tables] addTableRowBelow: Extracted column widths from DOM:', columnWidths);
              }
            }
          }
        } catch (e) {
          console.error('[ep_data_tables] addTableRowBelow: Error extracting column widths from DOM:', e);
        }
      }

      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex + 1;
        const newRowIndex = tableLines[i].metadata.row + 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };

        applyTableLineMetadataAttribute(lineToUpdate, tblId, newRowIndex, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      const newMetadata = { tblId, row: targetLine.metadata.row + 1, cols: numCols, columnWidths };
      applyTableLineMetadataAttribute(insertLineIndex, tblId, targetLine.metadata.row + 1, numCols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding row below with text:', e);
      return false;
    }
  }

  function addTableColumnLeftWithText(tableLines, targetColIndex, editorInfo, docManager) {
    const funcName = 'addTableColumnLeftWithText';
    try {
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);

        let insertPos = 0;
        for (let i = 0; i < targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0) + DELIMITER.length;
        }

        const textToInsert = ' ' + DELIMITER;
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];

        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);

        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (lineEntry) {
          const newLineText = lineEntry.text || '';
          const newCells = newLineText.split(DELIMITER);
          let offset = 0;

          for (let c = 0; c < tableLine.cols + 1; c++) {
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) {
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
             // log(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL} attribute to Line ${tableLine.lineIndex} Col ${c} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }

        const newColCount = tableLine.cols + 1;
        const equalWidth = 100 / newColCount;
        const normalizedWidths = Array(newColCount).fill(equalWidth);
      // log(`[ep_data_tables] addTableColumnLeft: Reset all column widths to equal distribution: ${newColCount} columns at ${equalWidth.toFixed(1)}% each`);

        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols + 1, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols + 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding column left with text:', e);
      return false;
    }
  }

  function addTableColumnRightWithText(tableLines, targetColIndex, editorInfo, docManager) {
    const funcName = 'addTableColumnRightWithText';
    try {
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);

        let insertPos = 0;
        for (let i = 0; i <= targetColIndex; i++) {
          insertPos += (cells[i]?.length ?? 0);
          if (i < targetColIndex) insertPos += DELIMITER.length;
        }

        const textToInsert = DELIMITER + ' ';
        const insertStart = [tableLine.lineIndex, insertPos];
        const insertEnd = [tableLine.lineIndex, insertPos];

        editorInfo.ace_performDocumentReplaceRange(insertStart, insertEnd, textToInsert);

        const rep = editorInfo.ace_getRep();
        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (lineEntry) {
          const newLineText = lineEntry.text || '';
          const newCells = newLineText.split(DELIMITER);
          let offset = 0;

          for (let c = 0; c < tableLine.cols + 1; c++) {
            const cellContent = (c < newCells.length) ? newCells[c] || '' : '';
            if (cellContent.length > 0) {
              const cellStart = [tableLine.lineIndex, offset];
              const cellEnd = [tableLine.lineIndex, offset + cellContent.length];
             // log(`[ep_data_tables] ${funcName}: Applying ${ATTR_CELL} attribute to Line ${tableLine.lineIndex} Col ${c} Range ${offset}-${offset + cellContent.length}`);
              editorInfo.ace_performDocumentApplyAttributesToRange(cellStart, cellEnd, [[ATTR_CELL, String(c)]]);
            }
            offset += cellContent.length;
            if (c < newCells.length - 1) {
              offset += DELIMITER.length;
            }
          }
        }

        const newColCount = tableLine.cols + 1;
        const equalWidth = 100 / newColCount;
        const normalizedWidths = Array(newColCount).fill(equalWidth);
       // log(`[ep_data_tables] addTableColumnRight: Reset all column widths to equal distribution: ${newColCount} columns at ${equalWidth.toFixed(1)}% each`);

        const newMetadata = { ...tableLine.metadata, cols: tableLine.cols + 1, columnWidths: normalizedWidths };
        applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, tableLine.cols + 1, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error adding column right with text:', e);
      return false;
    }
  }

  function deleteTableRowWithText(tableLines, targetRowIndex, editorInfo, docManager) {
    try {
      const targetLine = tableLines[targetRowIndex];

      if (targetRowIndex === 0) {
       // log('[ep_data_tables] Deleting first row (row 0) - inserting blank line to prevent table from getting stuck');
        const insertStart = [targetLine.lineIndex, 0];
        editorInfo.ace_performDocumentReplaceRange(insertStart, insertStart, '\n');

        const deleteStart = [targetLine.lineIndex + 1, 0];
        const deleteEnd = [targetLine.lineIndex + 2, 0];
        editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      } else {
      const deleteStart = [targetLine.lineIndex, 0];
      const deleteEnd = [targetLine.lineIndex + 1, 0];
      editorInfo.ace_performDocumentReplaceRange(deleteStart, deleteEnd, '');
      }

      let columnWidths = targetLine.metadata.columnWidths;
      if (!columnWidths) {
        try {
          const rep = editorInfo.ace_getRep();
          for (const tableLine of tableLines) {
            if (tableLine.lineIndex !== targetLine.lineIndex) {
              const lineEntry = rep.lines.atIndex(tableLine.lineIndex >= targetLine.lineIndex ? tableLine.lineIndex - 1 : tableLine.lineIndex);
              if (lineEntry && lineEntry.lineNode) {
                const tableInDOM = lineEntry.lineNode.querySelector(`table.dataTable[data-tblId="${targetLine.metadata.tblId}"]`);
                if (tableInDOM) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length === targetLine.metadata.cols) {
                    columnWidths = [];
                    domCells.forEach(cell => {
                      const style = cell.getAttribute('style') || '';
                      const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                      if (widthMatch) {
                        columnWidths.push(parseFloat(widthMatch[1]));
                      } else {
                        columnWidths.push(100 / targetLine.metadata.cols);
                      }
                    });
                   // log('[ep_data_tables] deleteTableRow: Extracted column widths from DOM:', columnWidths);
                    break;
                  }
                }
              }
            }
          }
        } catch (e) {
          console.error('[ep_data_tables] deleteTableRow: Error extracting column widths from DOM:', e);
        }
      }

      for (let i = targetRowIndex + 1; i < tableLines.length; i++) {
        const lineToUpdate = tableLines[i].lineIndex - 1;
        const newRowIndex = tableLines[i].metadata.row - 1;
        const newMetadata = { ...tableLines[i].metadata, row: newRowIndex, columnWidths };

        applyTableLineMetadataAttribute(lineToUpdate, tableLines[i].metadata.tblId, newRowIndex, tableLines[i].cols, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error deleting row with text:', e);
      return false;
    }
  }
  function deleteTableColumnWithText(tableLines, targetColIndex, editorInfo, docManager) {
    try {
      for (const tableLine of tableLines) {
        const lineText = tableLine.lineText;
        const cells = lineText.split(DELIMITER);

        if (targetColIndex >= cells.length) {
         // log(`[ep_data_tables] Warning: Target column ${targetColIndex} doesn't exist in line with ${cells.length} columns`);
          continue;
        }

        let deleteStart = 0;
        let deleteEnd = 0;

        for (let i = 0; i < targetColIndex; i++) {
          deleteStart += (cells[i]?.length ?? 0) + DELIMITER.length;
        }

        deleteEnd = deleteStart + (cells[targetColIndex]?.length ?? 0);

        if (targetColIndex === 0 && cells.length > 1) {
          deleteEnd += DELIMITER.length;
        } else if (targetColIndex > 0) {
          deleteStart -= DELIMITER.length;
        }

       // log(`[ep_data_tables] Deleting column ${targetColIndex} from line ${tableLine.lineIndex}: chars ${deleteStart}-${deleteEnd} from "${lineText}"`);

        const rangeStart = [tableLine.lineIndex, deleteStart];
        const rangeEnd = [tableLine.lineIndex, deleteEnd];

        editorInfo.ace_performDocumentReplaceRange(rangeStart, rangeEnd, '');

        const newColCount = tableLine.cols - 1;
        if (newColCount > 0) {
          const equalWidth = 100 / newColCount;
          const normalizedWidths = Array(newColCount).fill(equalWidth);
         // log(`[ep_data_tables] deleteTableColumn: Reset all column widths to equal distribution: ${newColCount} columns at ${equalWidth.toFixed(1)}% each`);

          const newMetadata = { ...tableLine.metadata, cols: newColCount, columnWidths: normalizedWidths };
          applyTableLineMetadataAttribute(tableLine.lineIndex, tableLine.metadata.tblId, tableLine.metadata.row, newColCount, editorInfo.ace_getRep(), editorInfo, JSON.stringify(newMetadata), docManager);
        }
      }

      editorInfo.ace_fastIncorp(10);
      return true;
    } catch (e) {
      console.error('[ep_data_tables] Error deleting column with text:', e);
      return false;
    }
  }


 // log('aceInitialized: END - helpers defined.');
};


const startColumnResize = (table, columnIndex, startX, metadata, lineNum) => {
  const funcName = 'startColumnResize';
 // log(`${funcName}: Starting resize for column ${columnIndex}`);

  isResizing = true;
  resizeStartX = startX;
  resizeCurrentX = startX;
  resizeTargetTable = table;
  resizeTargetColumn = columnIndex;
  resizeTableMetadata = metadata;
  resizeLineNum = lineNum;

  const numCols = metadata.cols;
  resizeOriginalWidths = metadata.columnWidths ? [...metadata.columnWidths] : Array(numCols).fill(100 / numCols);

 // log(`${funcName}: Original widths:`, resizeOriginalWidths);

  createResizeOverlay(table, columnIndex);

  document.body.style.userSelect = 'none';
  document.body.style.webkitUserSelect = 'none';
  document.body.style.mozUserSelect = 'none';
  document.body.style.msUserSelect = 'none';
};

const createResizeOverlay = (table, columnIndex) => {
  if (resizeOverlay) {
    resizeOverlay.remove();
  }

  const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
  if ($innerIframe.length === 0) {
    console.error('[ep_data_tables] createResizeOverlay: Could not find inner iframe');
    return;
  }

  const innerDocBody = $innerIframe.contents().find('body')[0];
  const padOuter = $('iframe[name="ace_outer"]').contents().find('body');

  if (!innerDocBody || padOuter.length === 0) {
    console.error('[ep_data_tables] createResizeOverlay: Could not find required container elements');
          return;
      }

  const tblId = table.getAttribute('data-tblId');
  if (!tblId) {
    console.error('[ep_data_tables] createResizeOverlay: No tblId found on table');
    return;
  }

  const allTableRows = innerDocBody.querySelectorAll(`table.dataTable[data-tblId="${tblId}"]`);
  if (allTableRows.length === 0) {
    console.error('[ep_data_tables] createResizeOverlay: No table rows found for tblId:', tblId);
    return;
  }

  let minTop = Infinity;
  let maxBottom = -Infinity;
  let tableLeft = 0;
  let tableWidth = 0;

  Array.from(allTableRows).forEach((tableRow, index) => {
    const rect = tableRow.getBoundingClientRect();
    minTop = Math.min(minTop, rect.top);
    maxBottom = Math.max(maxBottom, rect.bottom);

    if (index === 0) {
      tableLeft = rect.left;
      tableWidth = rect.width;
    }
  });

  const totalTableHeight = maxBottom - minTop;

 // log(`createResizeOverlay: Found ${allTableRows.length} table rows, total height: ${totalTableHeight}px`);

  let innerBodyRect, innerIframeRect, outerBodyRect;
  let scrollTopInner, scrollLeftInner, scrollTopOuter, scrollLeftOuter;

  try {
    innerBodyRect = innerDocBody.getBoundingClientRect();
    innerIframeRect = $innerIframe[0].getBoundingClientRect();
    outerBodyRect = padOuter[0].getBoundingClientRect();
    scrollTopInner = innerDocBody.scrollTop;
    scrollLeftInner = innerDocBody.scrollLeft;
    scrollTopOuter = padOuter.scrollTop();
    scrollLeftOuter = padOuter.scrollLeft();
  } catch (e) {
    console.error('[ep_data_tables] createResizeOverlay: Error getting container rects/scrolls:', e);
    return;
  }

  const tableTopRelInner = minTop - innerBodyRect.top + scrollTopInner;
  const tableLeftRelInner = tableLeft - innerBodyRect.left + scrollLeftInner;

  const innerFrameTopRelOuter = innerIframeRect.top - outerBodyRect.top + scrollTopOuter;
  const innerFrameLeftRelOuter = innerIframeRect.left - outerBodyRect.left + scrollLeftOuter;

  const overlayTopOuter = innerFrameTopRelOuter + tableTopRelInner;
  const overlayLeftOuter = innerFrameLeftRelOuter + tableLeftRelInner;

  const outerPadding = window.getComputedStyle(padOuter[0]);
  const outerPaddingTop = parseFloat(outerPadding.paddingTop) || 0;
  const outerPaddingLeft = parseFloat(outerPadding.paddingLeft) || 0;

  const MANUAL_OFFSET_TOP = 6;
  const MANUAL_OFFSET_LEFT = 39;

  const finalOverlayTop = overlayTopOuter + outerPaddingTop + MANUAL_OFFSET_TOP;
  const finalOverlayLeft = overlayLeftOuter + outerPaddingLeft + MANUAL_OFFSET_LEFT;

  const tds = table.querySelectorAll('td');
  const tds_array = Array.from(tds);
  let linePosition = 0;

  if (columnIndex < tds_array.length) {
    const currentTd = tds_array[columnIndex];
    const currentTdRect = currentTd.getBoundingClientRect();
    const currentRelativeLeft = currentTdRect.left - tableLeft;
    const currentWidth = currentTdRect.width;
    linePosition = currentRelativeLeft + currentWidth;
  }

  resizeOverlay = document.createElement('div');
  resizeOverlay.className = 'ep-data_tables-resize-overlay';
  resizeOverlay.style.cssText = `
    position: absolute;
    left: ${finalOverlayLeft}px;
    top: ${finalOverlayTop}px;
    width: ${tableWidth}px;
    height: ${totalTableHeight}px;
    pointer-events: none;
    z-index: 1000;
    background: transparent;
    box-sizing: border-box;
  `;

  const resizeLine = document.createElement('div');
  resizeLine.className = 'resize-line';
  resizeLine.style.cssText = `
    position: absolute;
    left: ${linePosition}px;
    top: 0;
    width: 2px;
    height: 100%;
    background: #1a73e8;
    z-index: 1001;
  `;
  resizeOverlay.appendChild(resizeLine);

  padOuter.append(resizeOverlay);

 // log('createResizeOverlay: Created Google Docs style blue line overlay spanning entire table height');
};

const updateColumnResize = (currentX) => {
  if (!isResizing || !resizeTargetTable || !resizeOverlay) return;

  resizeCurrentX = currentX;
  const deltaX = currentX - resizeStartX;

  const tblId = resizeTargetTable.getAttribute('data-tblId');
  if (!tblId) return;

  const $innerIframe = $('iframe[name="ace_outer"]').contents().find('iframe[name="ace_inner"]');
  const innerDocBody = $innerIframe.contents().find('body')[0];
  const firstTableRow = innerDocBody.querySelector(`table.dataTable[data-tblId="${tblId}"]`);

  if (!firstTableRow) return;

  const tableRect = firstTableRow.getBoundingClientRect();
  const deltaPercent = (deltaX / tableRect.width) * 100;

  const newWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;

  if (nextColumn < newWidths.length) {
    const transfer = Math.min(deltaPercent, newWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(newWidths[currentColumn] - 5));

    newWidths[currentColumn] += actualTransfer;
    newWidths[nextColumn] -= actualTransfer;

    const resizeLine = resizeOverlay.querySelector('.resize-line');
    if (resizeLine) {
      const newColumnWidth = (newWidths[currentColumn] / 100) * tableRect.width;

      const tds = firstTableRow.querySelectorAll('td');
      const tds_array = Array.from(tds);

      if (currentColumn < tds_array.length) {
        const currentTd = tds_array[currentColumn];
        const currentTdRect = currentTd.getBoundingClientRect();
        const currentRelativeLeft = currentTdRect.left - tableRect.left;

        const newLinePosition = currentRelativeLeft + newColumnWidth;
        resizeLine.style.left = newLinePosition + 'px';
      }
    }
  }
};

const finishColumnResize = (editorInfo, docManager) => {
  if (!isResizing || !resizeTargetTable) {
   // log('finishColumnResize: Not in resize mode');
                      return;
                  }

  const funcName = 'finishColumnResize';
 // log(`${funcName}: Finishing resize`);

  const tableRect = resizeTargetTable.getBoundingClientRect();
  const deltaX = resizeCurrentX - resizeStartX;
  const deltaPercent = (deltaX / tableRect.width) * 100;

 // log(`${funcName}: Mouse moved ${deltaX}px (${deltaPercent.toFixed(1)}%)`);

  const finalWidths = [...resizeOriginalWidths];
  const currentColumn = resizeTargetColumn;
  const nextColumn = currentColumn + 1;

  if (nextColumn < finalWidths.length) {
    const transfer = Math.min(deltaPercent, finalWidths[nextColumn] - 5);
    const actualTransfer = Math.max(transfer, -(finalWidths[currentColumn] - 5));

    finalWidths[currentColumn] += actualTransfer;
    finalWidths[nextColumn] -= actualTransfer;

   // log(`${funcName}: Transferred ${actualTransfer.toFixed(1)}% from column ${nextColumn} to column ${currentColumn}`);
  }

  const totalWidth = finalWidths.reduce((sum, width) => sum + width, 0);
  if (totalWidth > 0) {
    finalWidths.forEach((width, index) => {
      finalWidths[index] = (width / totalWidth) * 100;
    });
  }

 // log(`${funcName}: Final normalized widths:`, finalWidths.map(w => w.toFixed(1) + '%'));

  if (resizeOverlay) {
    resizeOverlay.remove();
    resizeOverlay = null;
  }

  document.body.style.userSelect = '';
  document.body.style.webkitUserSelect = '';
  document.body.style.mozUserSelect = '';
  document.body.style.msUserSelect = '';

  isResizing = false;

  editorInfo.ace_callWithAce((ace) => {
    const callWithAceLogPrefix = `${funcName}[ace_callWithAce]`;
   // log(`${callWithAceLogPrefix}: Finding and updating all table rows with tblId: ${resizeTableMetadata.tblId}`);

    try {
      const rep = ace.ace_getRep();
      if (!rep || !rep.lines) {
        console.error(`${callWithAceLogPrefix}: Invalid rep`);
        return;
      }

      const tableLines = [];
      const totalLines = rep.lines.length();

      for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
        try {
          let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

          if (lineAttrString) {
            const lineMetadata = JSON.parse(lineAttrString);
            if (lineMetadata.tblId === resizeTableMetadata.tblId) {
              tableLines.push({
                lineIndex,
                metadata: lineMetadata
              });
              }
          } else {
            const lineEntry = rep.lines.atIndex(lineIndex);
            if (lineEntry && lineEntry.lineNode) {
              const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
              if (tableInDOM) {
                const domTblId = tableInDOM.getAttribute('data-tblId');
                const domRow = tableInDOM.getAttribute('data-row');
                if (domTblId === resizeTableMetadata.tblId && domRow !== null) {
                  const domCells = tableInDOM.querySelectorAll('td');
                  if (domCells.length > 0) {
                    const columnWidths = [];
                    domCells.forEach(cell => {
                      const style = cell.getAttribute('style') || '';
                      const widthMatch = style.match(/width:\s*([0-9.]+)%/);
                      if (widthMatch) {
                        columnWidths.push(parseFloat(widthMatch[1]));
                      } else {
                        columnWidths.push(100 / domCells.length);
                      }
                    });

                    const reconstructedMetadata = {
                      tblId: domTblId,
                      row: parseInt(domRow, 10),
                      cols: domCells.length,
                      columnWidths: columnWidths
                    };
                   // log(`${callWithAceLogPrefix}: Reconstructed metadata from DOM for line ${lineIndex}:`, reconstructedMetadata);
                    tableLines.push({
                      lineIndex,
                      metadata: reconstructedMetadata
                    });
                  }
                }
              }
            }
                  }
              } catch (e) {
          continue;
        }
      }

     // log(`${callWithAceLogPrefix}: Found ${tableLines.length} table lines to update`);

      for (const tableLine of tableLines) {
        const updatedMetadata = { ...tableLine.metadata, columnWidths: finalWidths };
        const updatedMetadataString = JSON.stringify(updatedMetadata);

        const lineEntry = rep.lines.atIndex(tableLine.lineIndex);
        if (!lineEntry) {
          console.error(`${callWithAceLogPrefix}: Could not get line entry for line ${tableLine.lineIndex}`);
          continue;
        }

        const lineLength = Math.max(1, lineEntry.text.length);
        const rangeStart = [tableLine.lineIndex, 0];
        const rangeEnd = [tableLine.lineIndex, lineLength];

       // log(`${callWithAceLogPrefix}: Updating line ${tableLine.lineIndex} (row ${tableLine.metadata.row}) with new column widths`);

        ace.ace_performDocumentApplyAttributesToRange(rangeStart, rangeEnd, [
          [ATTR_TABLE_JSON, updatedMetadataString]
        ]);
      }

     // log(`${callWithAceLogPrefix}: Successfully applied updated column widths to all ${tableLines.length} table rows`);

    } catch (error) {
      console.error(`${callWithAceLogPrefix}: Error applying updated metadata:`, error);
     // log(`${callWithAceLogPrefix}: Error details:`, { message: error.message, stack: error.stack });
    }
  }, 'applyTableResizeToAllRows', true);

 // log(`${funcName}: Column width update initiated for all table rows via ace_callWithAce`);

  resizeStartX = 0;
  resizeCurrentX = 0;
  resizeTargetTable = null;
  resizeTargetColumn = -1;
  resizeOriginalWidths = [];
  resizeTableMetadata = null;
  resizeLineNum = -1;

 // log(`${funcName}: Resize complete - state reset`);
};



module.exports = aceInitialized;
