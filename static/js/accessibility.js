'use strict';

(() => {
  const tableModel = typeof window !== 'undefined' && window.epDataTablesTableModel
    ? window.epDataTablesTableModel
    : require('./table_model');
  const {headerRowCount, tableCaption} = tableModel;

  const CONTEXT_CLASS = 'ep-data_tables-a11y-context';
  const pendingRefreshes = new WeakMap();
  const pendingDetachedTables = new WeakSet();

  const toDomElement = (node) => {
    if (!node) return null;
    if (node.jquery) return node[0] || null;
    return node;
  };

  const interpolate = (template, values) => String(template).replace(
      /\{\{\s*([^}\s]+)\s*\}\}/g, (_match, key) => values[key] == null ? '' : values[key]);

  const message = (key, values, fallback, translate) => {
    try {
      const localized = translate && translate(key, values);
      if (localized && localized !== key) return localized;
    } catch (_) {
      // A missing early translation must never block editor rendering.
    }
    return interpolate(fallback, values);
  };

  const defaultTranslate = (key, values) => {
    if (typeof html10n === 'undefined' || !html10n ||
        typeof html10n.get !== 'function') return null;
    return html10n.get(key, values);
  };

  const tableId = (table) => table && typeof table.getAttribute === 'function'
    ? table.getAttribute('data-tblId') || table.getAttribute('data-tblid') || ''
    : '';

  const isEditableDocument = (document) => {
    const body = document && document.body;
    return Boolean(body && (
      body.isContentEditable ||
      body.getAttribute?.('contenteditable') === 'true' ||
      String(document.designMode || '').toLowerCase() === 'on'
    ));
  };

  const cellText = (cell) => String(cell && cell.textContent || '')
      .replace(/\u241F/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

  const clearGeneratedContext = (table) => {
    if (!table || typeof table.querySelectorAll !== 'function') return;
    for (const context of table.querySelectorAll(`.${CONTEXT_CLASS}`)) context.remove();
    for (const cell of table.querySelectorAll('[data-ep-data-tables-description-id]')) {
      const generatedId = cell.getAttribute('data-ep-data-tables-description-id');
      const describedBy = String(cell.getAttribute('aria-describedby') || '')
          .split(/\s+/).filter((id) => id && id !== generatedId);
      if (describedBy.length) cell.setAttribute('aria-describedby', describedBy.join(' '));
      else cell.removeAttribute('aria-describedby');
      cell.removeAttribute('data-ep-data-tables-description-id');
    }
  };

  const refreshLogicalTableAccessibility = (root, logicalTableId, options = {}) => {
    if (!root || typeof root.querySelectorAll !== 'function' || !logicalTableId) return 0;
    const tables = Array.from(root.querySelectorAll(
        'table.dataTable[data-tblId], table.dataTable[data-tblid]'))
        .filter((table) => tableId(table) === logicalTableId)
        .sort((a, b) => Number(a.getAttribute('data-row')) - Number(b.getAttribute('data-row')));
    if (!tables.length) return 0;

    for (const table of tables) clearGeneratedContext(table);

    const translate = options.translate || defaultTranslate;
    const rowCount = Math.max(
        tables.length,
        ...tables.map((table) => (Number(table.getAttribute('data-row')) || 0) + 1));
    const cellSelector = 'tbody > tr > td, tbody > tr > th';
    const columnCounts = tables.map((table) => table.querySelectorAll(cellSelector).length);
    const columnCount = Math.max(0, ...columnCounts);
    const isFirstRow = (table) => Number(table.getAttribute('data-row')) === 0;
    const firstTable = tables.find(isFirstRow) || tables[0];
    const headerCells = Array.from(firstTable.querySelectorAll(
        'tbody > tr > th[scope="col"], tbody > tr > [role="columnheader"]'));
    const headerLabels = headerCells.map(cellText);
    const caption = tables.map((table) => table.getAttribute('data-ep-data-tables-caption') || '')
        .find(Boolean) || '';
    const baseLabel = caption || message(
        'ep_data_tables.a11y.dataTable', {}, 'Data table', translate);

    tables.forEach((table) => {
      const rowIndex = (Number(table.getAttribute('data-row')) || 0) + 1;
      table.setAttribute('aria-rowcount', String(rowCount));
      table.setAttribute('aria-colcount', String(columnCount));
      table.setAttribute('aria-label', message(
          'ep_data_tables.a11y.tableRow',
          {table: baseLabel, row: rowIndex, rows: rowCount},
          '{{table}}, row {{row}} of {{rows}}', translate));
      const row = table.querySelector('tbody > tr');
      if (row) row.setAttribute('aria-rowindex', String(rowIndex));

      const cells = Array.from(table.querySelectorAll('tbody > tr > td, tbody > tr > th'));
      cells.forEach((cell, columnIndex) => {
        cell.setAttribute('aria-colindex', String(columnIndex + 1));
        if (cell.tagName && cell.tagName.toLowerCase() === 'th') return;
        const document = cell.ownerDocument;
        if (!document || typeof document.createElement !== 'function') return;
        const column = headerLabels[columnIndex] || message(
            'ep_data_tables.a11y.columnNumber', {column: columnIndex + 1},
            'Column {{column}}', translate);
        const context = document.createElement('span');
        context.className = CONTEXT_CLASS;
        context.hidden = true;
        context.contentEditable = 'false';
        const safeTableId = logicalTableId.replace(/[^A-Za-z0-9_-]/g, '-');
        context.id = `ep-data-tables-context-${safeTableId}-${rowIndex}-${columnIndex + 1}`;
        context.textContent = message(
            'ep_data_tables.a11y.cellContext',
            {column, row: rowIndex, rows: rowCount},
            '{{column}}, row {{row}} of {{rows}}', translate);
        cell.appendChild(context);
        const describedBy = String(cell.getAttribute('aria-describedby') || '')
            .split(/\s+/).filter(Boolean);
        if (!describedBy.includes(context.id)) describedBy.push(context.id);
        cell.setAttribute('aria-describedby', describedBy.join(' '));
        cell.setAttribute('data-ep-data-tables-description-id', context.id);
      });
    });
    return tables.length;
  };

  const queueDocumentRefresh = (document, logicalTableId, options) => {
    let pending = pendingRefreshes.get(document);
    if (!pending) {
      pending = new Map();
      pendingRefreshes.set(document, pending);
    }
    pending.set(logicalTableId, options);
    if (pending.scheduled) return;
    pending.scheduled = true;
    const scheduleFrame = document.defaultView &&
        typeof document.defaultView.requestAnimationFrame === 'function'
      ? (callback) => document.defaultView.requestAnimationFrame(callback)
      : (callback) => setTimeout(callback, 0);
    scheduleFrame(() => {
      pending.scheduled = false;
      const work = Array.from(pending.entries()).filter(([key]) => key !== 'scheduled');
      pending.clear();
      for (const [id, refreshOptions] of work) {
        refreshLogicalTableAccessibility(document, id, refreshOptions);
      }
    });
  };

  const scheduleLogicalTableAccessibility = (table, options = {}) => {
    const tableElement = toDomElement(table);
    const logicalTableId = tableElement && tableId(tableElement);
    if (!tableElement || !logicalTableId || pendingDetachedTables.has(tableElement)) return;

    // Etherpad calls acePostWriteDomLineHTML while a line can still belong to a
    // detached construction document. Resolve ownerDocument after attachment;
    // capturing it here can permanently refresh the wrong document.
    pendingDetachedTables.add(tableElement);
    const refreshAfterAttachment = (attempt = 0) => {
      const document = tableElement.ownerDocument;
      if ((!document || tableElement.isConnected === false) && attempt < 4) {
        setTimeout(() => refreshAfterAttachment(attempt + 1), 0);
        return;
      }
      pendingDetachedTables.delete(tableElement);
      if (document) queueDocumentRefresh(document, logicalTableId, options);
    };
    setTimeout(refreshAfterAttachment, 0);
  };

  const enhanceTableMarkup = (table, metadata = {}, options = {}) => {
    const tableElement = toDomElement(table);
    if (!tableElement || typeof tableElement.querySelectorAll !== 'function') return;
    const rowIndex = (Number(metadata.row) || 0) + 1;
    const caption = tableCaption(metadata);
    tableElement.setAttribute('data-ep-data-tables-accessible', 'true');
    tableElement.setAttribute('data-ep-data-tables-header-rows', String(headerRowCount(metadata)));
    if (caption) tableElement.setAttribute('data-ep-data-tables-caption', caption);
    else tableElement.removeAttribute?.('data-ep-data-tables-caption');
    tableElement.setAttribute('aria-label', caption || 'Data table');

    const row = tableElement.querySelector?.('tbody > tr');
    if (row) row.setAttribute('aria-rowindex', String(rowIndex));
    const cells = Array.from(tableElement.querySelectorAll('td, th'));
    cells.forEach((cell, index) => {
      cell.setAttribute('aria-colindex', String(index + 1));
      if (cell.tagName && cell.tagName.toLowerCase() === 'th' && !cell.getAttribute('scope')) {
        cell.setAttribute('scope', metadata.row < headerRowCount(metadata) ? 'col' : 'row');
      }
    });

    const artifacts = tableElement.querySelectorAll(
        '.ep-data_tables-delim, .ep-data_tables-caret-anchor');
    for (const delimiter of artifacts) {
      delimiter.setAttribute('aria-hidden', 'true');
    }
    for (const handle of tableElement.querySelectorAll('.ep-data_tables-resize-handle')) {
      handle.setAttribute('aria-hidden', 'true');
      handle.setAttribute('tabindex', '-1');
    }
    if (options.logicalRefresh !== false && !isEditableDocument(tableElement.ownerDocument)) {
      scheduleLogicalTableAccessibility(tableElement, options);
    }
  };

  const enhanceTableHtml = (html, metadata = {}, options = {}) => {
    if (typeof document === 'undefined') return html;
    const template = document.createElement('template');
    template.innerHTML = html;
    const table = template.content.querySelector('table.dataTable');
    enhanceTableMarkup(table, metadata, options);
    return template.innerHTML;
  };

  const setupMenuAccessibility = (
      {$, menu, gridWrap, gridCells, sizeText, toolbarButton, createTable}) => {
    const $menu = $(menu);
    const $gridWrap = $(gridWrap);
    const $gridCells = $(gridCells);
    const $toolbarButton = $(toolbarButton);

    const syncMenuState = () => {
      const isOpen = $menu.is(':visible');
      $toolbarButton.find('button').attr('aria-expanded', isOpen ? 'true' : 'false');
    };
    const closeMenus = () => {
      $menu.hide();
      $gridWrap.hide();
      $menu.find('#tbl_prop_create_table').attr('aria-expanded', 'false');
      syncMenuState();
      $toolbarButton.find('button').trigger('focus');
    };
    const focusFirstMenuItem = () => {
      const first = $menu.find('[role="menuitem"]').first();
      if (first.length) first.trigger('focus');
    };

    $menu.find('[role="menuitem"]').attr('tabindex', '-1');
    $menu.on('keydown', '[role="menuitem"]', (evt) => {
      const items = $menu.find('[role="menuitem"]');
      const idx = items.index(evt.currentTarget);
      if (evt.key === 'Escape') {
        evt.preventDefault(); closeMenus();
      } else if (evt.key === 'ArrowDown') {
        evt.preventDefault(); items.eq((idx + 1) % items.length).trigger('focus');
      } else if (evt.key === 'ArrowUp') {
        evt.preventDefault(); items.eq((idx - 1 + items.length) % items.length).trigger('focus');
      } else if (evt.key === 'Enter' || evt.key === ' ') {
        evt.preventDefault(); $(evt.currentTarget).trigger('click');
      } else if (evt.key === 'ArrowRight' &&
          evt.currentTarget.id === 'tbl_prop_create_table') {
        evt.preventDefault();
        $gridWrap.show();
        $gridCells.first().attr('tabindex', '0').trigger('focus');
      }
    });
    $toolbarButton.on('keydown', (evt) => {
      if (evt.key === 'ArrowDown' || evt.key === 'Enter' || evt.key === ' ') {
        setTimeout(focusFirstMenuItem, 0);
      } else if (evt.key === 'Escape') { closeMenus(); }
    });

    const updateGridSelection = (cell) => {
      const $cell = $(cell);
      const row = $cell.parent().index();
      const col = $cell.index();
      $gridCells.removeClass('selected').attr('aria-selected', 'false').attr('tabindex', '-1');
      for (let r = 0; r <= row; r++) {
        for (let c = 0; c <= col; c++) {
          $('#new-table-size-selector tr').eq(r).find('td').eq(c)
              .addClass('selected').attr('aria-selected', 'true');
        }
      }
      $cell.attr('tabindex', '0');
      sizeText.text(`${col + 1} × ${row + 1}`);
    };

    $gridCells.attr('role', 'gridcell').attr('aria-selected', 'false').attr('tabindex', '-1');
    $gridCells.first().attr('tabindex', '0');
    $gridCells.on('focus', (evt) => updateGridSelection(evt.currentTarget));
    $gridWrap.on('keydown', 'td', (evt) => {
      const $cell = $(evt.currentTarget);
      const row = $cell.parent().index();
      const col = $cell.index();
      const rows = $('#new-table-size-selector tr').length;
      const cols = $('#new-table-size-selector tr').first().find('td').length;
      let nextRow = row;
      let nextCol = col;
      if (evt.key === 'Escape') { evt.preventDefault(); closeMenus(); return; }
      if (evt.key === 'Enter' || evt.key === ' ') {
        evt.preventDefault();
        const [selectedCols, selectedRows] = sizeText.text().split(/\s*[X×]\s*/)
            .map((value) => parseInt(value, 10));
        createTable(selectedRows, selectedCols); closeMenus(); return;
      }
      if (evt.key === 'ArrowRight') nextCol = Math.min(cols - 1, col + 1);
      else if (evt.key === 'ArrowLeft') nextCol = Math.max(0, col - 1);
      else if (evt.key === 'ArrowDown') nextRow = Math.min(rows - 1, row + 1);
      else if (evt.key === 'ArrowUp') nextRow = Math.max(0, row - 1);
      else return;
      evt.preventDefault();
      $('#new-table-size-selector tr').eq(nextRow).find('td').eq(nextCol).trigger('focus');
    });
    return {closeMenus, syncMenuState};
  };

  const api = {
    CONTEXT_CLASS,
    clearGeneratedContext,
    enhanceTableHtml,
    enhanceTableMarkup,
    isEditableDocument,
    refreshLogicalTableAccessibility,
    scheduleLogicalTableAccessibility,
    setupMenuAccessibility,
    toDomElement,
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.epDataTablesAccessibility = api;
})();
