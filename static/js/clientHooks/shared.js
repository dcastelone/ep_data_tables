'use strict';

const ATTR_TABLE_JSON = 'tbljson';
const ATTR_CELL = 'td';
const ATTR_CLASS_PREFIX = 'tbljson-';
const log = (...m) => console.debug('[ep_data_tables:client_hooks]', ...m);
const DELIMITER = '\u241F';
const HIDDEN_DELIM = DELIMITER;
const INPUTTYPE_REPLACEMENT_TYPES = new Set(['insertReplacementText', 'insertFromComposition']);

const rand = () => Math.random().toString(36).slice(2, 8);

const enc = (s) => btoa(s).replace(/\+/g, '-').replace(/\//g, '_');
const dec = (s) => {
  const str = s.replace(/-/g, '+').replace(/_/g, '/');
  try {
    if (typeof atob === 'function') {
      return atob(str);
    } else if (typeof Buffer === 'function') {
      return Buffer.from(str, 'base64').toString('utf8');
    }
    console.error('[ep_data_tables] Base64 decoding function (atob or Buffer) not found.');
    return null;
  } catch (e) {
    console.error('[ep_data_tables] Error decoding base64 string:', s, e);
    return null;
  }
};

const TBLJSON_CLASS_RE = /\btbljson-([A-Za-z0-9_-]+)/;

const toClassTokens = (classSource) => {
  if (!classSource) return [];
  if (typeof classSource === 'string') {
    return classSource.trim().split(/\s+/).filter(Boolean);
  }
  if (Array.isArray(classSource)) return classSource.filter(Boolean);
  if (typeof classSource === 'object' && typeof classSource.length === 'number') {
    return Array.from(classSource, (token) => `${token}`).filter(Boolean);
  }
  return [];
};

const extractEncodedTbljson = (classSource) => {
  for (const token of toClassTokens(classSource)) {
    const match = TBLJSON_CLASS_RE.exec(token);
    if (match) return match[1];
  }
  return null;
};

const isWellFormedMetadata = (metadata) => (
  metadata &&
  typeof metadata.tblId !== 'undefined' &&
  typeof metadata.row !== 'undefined' &&
  typeof metadata.cols === 'number'
);

const decodeTbljsonClass = (classSource) => {
  const encoded = extractEncodedTbljson(classSource);
  if (!encoded) return null;
  const json = dec(encoded);
  if (!json) return null;
  let metadata = null;
  let parseError = null;
  try {
    metadata = JSON.parse(json);
  } catch (err) {
    parseError = err;
  }
  return {
    encoded,
    json,
    metadata,
    isWellFormed: isWellFormedMetadata(metadata),
    error: parseError,
  };
};

const normalizeSoftWhitespace = (str) => (
  (str || '')
    .replace(/[\u00A0\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
);

function isAndroidUA() {
  const ua = (navigator.userAgent || '').toLowerCase();
  const isAndroid = ua.includes('android');
  const isIOS = ua.includes('iphone') || ua.includes('ipad') || ua.includes('ipod') || ua.includes('crios');
  return isAndroid && !isIOS;
}

function isIOSUA() {
  const ua = (navigator.userAgent || '').toLowerCase();
  return ua.includes('iphone') || ua.includes('ipad') || ua.includes('ipod') || ua.includes('ios') || ua.includes('crios') || ua.includes('fxios') || ua.includes('edgios');
}

function findTbljsonElement(element) {
  if (!element) return null;
  if (extractEncodedTbljson(element.classList || element.className || '')) {
    return element;
  }
  const {children} = element;
  if (!children || !children.length) return null;
  for (const child of children) {
    const found = findTbljsonElement(child);
    if (found) return found;
  }
  return null;
}

function getTableLineMetadata(lineNum, editorInfo, docManager) {
  const funcName = 'getTableLineMetadata';
  try {
    const attribs = docManager.getAttributeOnLine(lineNum, ATTR_TABLE_JSON);
    if (attribs) {
      try {
        const metadata = JSON.parse(attribs);
        if (metadata && metadata.tblId) {
          return metadata;
        }
      } catch (e) {
        // ignore malformed metadata
      }
    }

    const rep = editorInfo.ace_getRep();
    const lineEntry = rep.lines.atIndex(lineNum);
    const lineNode = lineEntry?.lineNode;

    if (!lineNode) {
      return null;
    }

    const tbljsonElement = findTbljsonElement(lineNode);
    if (tbljsonElement) {
      const info = decodeTbljsonClass(tbljsonElement.classList || tbljsonElement.className || '');
      if (info) {
        if (info.metadata && info.metadata.tblId) {
          return info.metadata;
        }
        if (info.error) {
          console.error(`${funcName}: Failed to decode/parse tbljson class on line ${lineNum}:`, info.error);
        }
      }
    }

    return null;
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error getting metadata for line ${lineNum}:`, e);
    return null;
  }
}

function findLineForTableRow(tblId, targetRow, editorInfo, docManager) {
  const funcName = 'findLineForTableRow';

  try {
    const rep = editorInfo.ace_getRep();
    if (!rep || !rep.lines) {
      return -1;
    }

    const totalLines = rep.lines.length();
    for (let lineIndex = 0; lineIndex < totalLines; lineIndex++) {
      try {
        let lineAttrString = docManager.getAttributeOnLine(lineIndex, ATTR_TABLE_JSON);

        if (!lineAttrString) {
          const lineEntry = rep.lines.atIndex(lineIndex);
          if (lineEntry && lineEntry.lineNode) {
            const tableInDOM = lineEntry.lineNode.querySelector('table.dataTable[data-tblId]');
            if (tableInDOM) {
              const domTblId = tableInDOM.getAttribute('data-tblId');
              const domRow = tableInDOM.getAttribute('data-row');
              if (domTblId === tblId && domRow !== null && parseInt(domRow, 10) === targetRow) {
                return lineIndex;
              }
            }
          }
        }

        if (lineAttrString) {
          const lineMetadata = JSON.parse(lineAttrString);
          if (lineMetadata.tblId === tblId && lineMetadata.row === targetRow) {
            return lineIndex;
          }
        }
      } catch (e) {
        continue;
      }
    }

    return -1;
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error searching for line:`, e);
    return -1;
  }
}

function navigateToCell(targetLineNum, targetCellIndex, editorInfo, docManager) {
  const funcName = 'navigateToCell';
  let targetPos;

  try {
    const rep = editorInfo.ace_getRep();
    if (!rep || !rep.lines) {
      return false;
    }

    const lineEntry = rep.lines.atIndex(targetLineNum);
    if (!lineEntry) {
      return false;
    }

    const lineText = lineEntry.text || '';
    const cells = lineText.split(DELIMITER);

    if (targetCellIndex >= cells.length) {
      return false;
    }

    let targetCol = 0;
    for (let i = 0; i < targetCellIndex; i++) {
      targetCol += (cells[i]?.length ?? 0) + DELIMITER.length;
    }
    const targetCellContent = cells[targetCellIndex] || '';
    targetCol += targetCellContent.length;

    const clampedTargetCol = Math.min(targetCol, lineText.length);
    targetPos = [targetLineNum, clampedTargetCol];

    try {
      const editor = editorInfo.ep_data_tables_editor;
      const tableMetadata = getTableLineMetadata(targetLineNum, editorInfo, docManager);

      if (editor && tableMetadata) {
        editor.ep_data_tables_last_clicked = {
          lineNum: targetLineNum,
          tblId: tableMetadata.tblId,
          cellIndex: targetCellIndex,
          relativePos: targetCellContent.length,
        };
      }
    } catch (e) {
      // ignore
    }

    try {
      editorInfo.ace_performSelectionChange(targetPos, targetPos, false);
      editorInfo.ace_updateBrowserSelectionFromRep();
      editorInfo.ace_focus();
    } catch (e) {
      console.error(`[ep_data_tables] ${funcName}: Error during direct navigation update:`, e);
      return false;
    }
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during cell navigation:`, e);
    return false;
  }

  return true;
}

function navigateToNextCell(currentLineNum, currentCellIndex, tableMetadata, shiftKey, editorInfo, docManager) {
  const funcName = 'navigateToNextCell';

  try {
    let targetRow = tableMetadata.row;
    let targetCol = currentCellIndex;

    if (shiftKey) {
      targetCol--;
      if (targetCol < 0) {
        targetRow--;
        targetCol = tableMetadata.cols - 1;
      }
    } else {
      targetCol++;
      if (targetCol >= tableMetadata.cols) {
        targetRow++;
        targetCol = 0;
      }
    }

    const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);
    if (targetLineNum === -1) {
      return false;
    }

    return navigateToCell(targetLineNum, targetCol, editorInfo, docManager);
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during navigation:`, e);
    return false;
  }
}

function navigateToCellBelow(currentLineNum, currentCellIndex, tableMetadata, editorInfo, docManager) {
  const funcName = 'navigateToCellBelow';

  try {
    const targetRow = tableMetadata.row + 1;
    const targetCol = currentCellIndex;

    const targetLineNum = findLineForTableRow(tableMetadata.tblId, targetRow, editorInfo, docManager);

    if (targetLineNum !== -1) {
      return navigateToCell(targetLineNum, targetCol, editorInfo, docManager);
    }

    const rep = editorInfo.ace_getRep();
    const lineTextLength = rep.lines.atIndex(currentLineNum).text.length;
    const endOfLinePos = [currentLineNum, lineTextLength];

    editorInfo.ace_performSelectionChange(endOfLinePos, endOfLinePos, false);
    editorInfo.ace_performDocumentReplaceRange(endOfLinePos, endOfLinePos, '\n');

    editorInfo.ace_updateBrowserSelectionFromRep();
    editorInfo.ace_focus();

    const editor = editorInfo.editor;
    if (editor) editor.ep_data_tables_last_clicked = null;

    return true;
  } catch (e) {
    console.error(`[ep_data_tables] ${funcName}: Error during navigation:`, e);
    return false;
  }
}

function escapeHtml(text = '') {
  const strText = String(text);
  const map = {
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  };
  return strText.replace(/[&<>"'']/g, (m) => map[m]);
}

function buildTableFromDelimitedHTML(metadata, innerHTMLSegments) {
  const funcName = 'buildTableFromDelimitedHTML';

  if (!metadata || typeof metadata.tblId === 'undefined' || typeof metadata.row === 'undefined') {
    console.error(`[ep_data_tables] ${funcName}: Invalid or missing metadata. Aborting.`);
    return '<table class="dataTable dataTable-error"><tbody><tr><td>Error: Missing table metadata</td></tr></tbody></table>';
  }

  const numCols = innerHTMLSegments.length;
  const columnWidths = metadata.columnWidths || Array(numCols).fill(100 / numCols);

  while (columnWidths.length < numCols) {
    columnWidths.push(100 / numCols);
  }
  if (columnWidths.length > numCols) {
    columnWidths.splice(numCols);
  }

  const tdStyle = 'padding: 5px 7px; word-wrap:break-word; vertical-align: top; border: 1px solid #000; position: relative;';

  let encodedTbljsonClass = '';
  try {
    encodedTbljsonClass = `tbljson-${enc(JSON.stringify(metadata))}`;
  } catch (_) { encodedTbljsonClass = ''; }

  const cellsHtml = innerHTMLSegments.map((segment, index) => {
    const textOnly = (segment || '').replace(/<[^>]*>/g, '').replace(/&nbsp;/ig, ' ').trim();
    let modifiedSegment = segment || '';
    const isEmpty = !segment || textOnly === '';
    if (isEmpty) {
      const cellClass = encodedTbljsonClass ? `${encodedTbljsonClass} tblCell-${index}` : `tblCell-${index}`;
      modifiedSegment = `<span class="${cellClass}">&nbsp;</span>`;
    }
    if (index > 0) {
      const delimSpan = `<span class="ep-data_tables-delim" contenteditable="false">${HIDDEN_DELIM}</span>`;
      modifiedSegment = modifiedSegment.replace(/^(<span[^>]*>)/i, `$1${delimSpan}`);
      if (!/^<span[^>]*>/i.test(modifiedSegment)) modifiedSegment = `${delimSpan}${modifiedSegment}`;
    }

    const caretAnchorSpan = '<span class="ep-data_tables-caret-anchor" contenteditable="false"></span>';
    const anchorInjected = modifiedSegment.replace(/<\/span>\s*$/i, `${caretAnchorSpan}</span>`);
    modifiedSegment = (anchorInjected !== modifiedSegment)
      ? anchorInjected
      : (isEmpty
          ? `<span class="${encodedTbljsonClass ? `${encodedTbljsonClass} ` : ''}tblCell-${index}">${modifiedSegment}${caretAnchorSpan}</span>`
          : `${modifiedSegment}${caretAnchorSpan}`);

    try {
      const requiredCellClass = `tblCell-${index}`;
      const leadingDelimMatch = modifiedSegment.match(/^\s*<span[^>]*\bep-data_tables-delim\b[^>]*>[\s\S]*?<\/span>\s*/i);
      const head = leadingDelimMatch ? leadingDelimMatch[0] : '';
      const tail = leadingDelimMatch ? modifiedSegment.slice(head.length) : modifiedSegment;

      const openSpanMatch = tail.match(/^\s*<span([^>]*)>/i);
      if (!openSpanMatch) {
        const baseClasses = `${encodedTbljsonClass ? `${encodedTbljsonClass} ` : ''}${requiredCellClass}`;
        modifiedSegment = `${head}<span class="${baseClasses}">${tail}</span>`;
      } else {
        const fullOpen = openSpanMatch[0];
        const attrs = openSpanMatch[1] || '';
        const classMatch = /\bclass\s*=\s*"([^"]*)"/i.exec(attrs);
        let classList = classMatch ? classMatch[1].split(/\s+/).filter(Boolean) : [];
        classList = classList.filter((c) => !/^tblCell-\d+$/.test(c));
        classList.push(requiredCellClass);
        const unique = Array.from(new Set(classList));
        const newClassAttr = ` class="${unique.join(' ')}"`;
        const attrsWithoutClass = classMatch ? attrs.replace(/\s*class\s*=\s*"[^"]*"/i, '') : attrs;
        const rebuiltOpen = `<span${newClassAttr}${attrsWithoutClass}>`;
        const afterOpen = tail.slice(fullOpen.length);
        const cleanedTail = afterOpen.replace(/(<span[^>]*class=")([^"]*)(")/ig, (m, p1, classes, p3) => {
          const filtered = classes.split(/\s+/).filter((c) => c && !/^tblCell-\d+$/.test(c)).join(' ');
          return p1 + filtered + p3;
        });
        modifiedSegment = head + rebuiltOpen + cleanedTail;
      }
    } catch (_) { /* ignore normalization errors */ }

    const widthPercent = columnWidths[index] || (100 / numCols);
    const cellStyle = `${tdStyle} width: ${widthPercent}%;`;

    const isLastColumn = index === innerHTMLSegments.length - 1;
    const resizeHandle = !isLastColumn
      ? `<div class="ep-data_tables-resize-handle" data-column="${index}" style="position: absolute; top: 0; right: -2px; width: 4px; height: 100%; cursor: col-resize; background: transparent; z-index: 10;"></div>`
      : '';

    const tdContent = `<td style="${cellStyle}" data-column="${index}" draggable="false">${modifiedSegment}${resizeHandle}</td>`;
    return tdContent;
  }).join('');

  const firstRowClass = metadata.row === 0 ? ' dataTable-first-row' : '';

  const tableHtml = `<table class="dataTable${firstRowClass}" writingsuggestions="false" data-tblId="${metadata.tblId}" data-row="${metadata.row}" style="width:100%; border-collapse: collapse; table-layout: fixed;" draggable="false"><tbody><tr>${cellsHtml}</tr></tbody></table>`;
  return tableHtml;
}

function _getLineNumberOfElement(element) {
  let currentElement = element;
  let count = 0;
  while (currentElement = currentElement.previousElementSibling) {
    count++;
  }
  return count;
}

module.exports = {
  ATTR_TABLE_JSON,
  ATTR_CELL,
  ATTR_CLASS_PREFIX,
  DELIMITER,
  HIDDEN_DELIM,
  INPUTTYPE_REPLACEMENT_TYPES,
  TBLJSON_CLASS_RE,
  buildTableFromDelimitedHTML,
  decodeTbljsonClass,
  enc,
  dec,
  extractEncodedTbljson,
  findLineForTableRow,
  findTbljsonElement,
  getTableLineMetadata,
  isAndroidUA,
  isIOSUA,
  isWellFormedMetadata,
  log,
  navigateToCell,
  navigateToCellBelow,
  navigateToNextCell,
  normalizeSoftWhitespace,
  rand,
  toClassTokens,
  escapeHtml,
  _getLineNumberOfElement,
};
