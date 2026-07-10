'use strict';

const toDomElement = (node) => {
  if (!node) return null;
  if (node.jquery) return node[0] || null;
  return node;
};

const enhanceTableMarkup = (table, metadata = {}) => {
  const tableElement = toDomElement(table);
  if (!tableElement || typeof tableElement.querySelectorAll !== 'function') return;
  tableElement.setAttribute('data-ep-data-tables-accessible', 'true');
  tableElement.setAttribute('aria-label', `Table row ${(Number(metadata.row) || 0) + 1}`);

  const cells = Array.from(tableElement.querySelectorAll('td, th'));
  cells.forEach((cell, index) => {
    cell.setAttribute('aria-colindex', String(index + 1));
  });

  for (const delimiter of tableElement.querySelectorAll('.ep-data_tables-delim, .ep-data_tables-caret-anchor')) {
    delimiter.setAttribute('aria-hidden', 'true');
  }

  for (const handle of tableElement.querySelectorAll('.ep-data_tables-resize-handle')) {
    handle.setAttribute('aria-hidden', 'true');
    handle.setAttribute('tabindex', '-1');
  }
};

const enhanceTableHtml = (html, metadata = {}) => {
  if (typeof document === 'undefined') return html;
  const template = document.createElement('template');
  template.innerHTML = html;
  const table = template.content.querySelector('table.dataTable');
  enhanceTableMarkup(table, metadata);
  return template.innerHTML;
};

const setupMenuAccessibility = ({$, menu, gridWrap, gridCells, sizeText, toolbarButton, createTable}) => {
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
      evt.preventDefault();
      closeMenus();
    } else if (evt.key === 'ArrowDown') {
      evt.preventDefault();
      items.eq((idx + 1) % items.length).trigger('focus');
    } else if (evt.key === 'ArrowUp') {
      evt.preventDefault();
      items.eq((idx - 1 + items.length) % items.length).trigger('focus');
    } else if (evt.key === 'Enter' || evt.key === ' ') {
      evt.preventDefault();
      $(evt.currentTarget).find('a,button').first().trigger('click');
    } else if (evt.key === 'ArrowRight' && evt.currentTarget.id === 'tbl_prop_create_table') {
      evt.preventDefault();
      $gridWrap.show();
      $gridCells.first().attr('tabindex', '0').trigger('focus');
    }
  });

  $toolbarButton.on('keydown', (evt) => {
    if (evt.key === 'ArrowDown' || evt.key === 'Enter' || evt.key === ' ') {
      setTimeout(focusFirstMenuItem, 0);
    } else if (evt.key === 'Escape') {
      closeMenus();
    }
  });

  const updateGridSelection = (cell) => {
    const $cell = $(cell);
    const row = $cell.parent().index();
    const col = $cell.index();
    $gridCells.removeClass('selected').attr('aria-selected', 'false').attr('tabindex', '-1');
    for (let r = 0; r <= row; r++) {
      for (let c = 0; c <= col; c++) {
        const selected = $('#new-table-size-selector tr').eq(r).find('td').eq(c);
        selected.addClass('selected').attr('aria-selected', 'true');
      }
    }
    $cell.attr('tabindex', '0');
    sizeText.text(`${col + 1} X ${row + 1}`);
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
    if (evt.key === 'Escape') {
      evt.preventDefault();
      closeMenus();
      return;
    }
    if (evt.key === 'Enter' || evt.key === ' ') {
      evt.preventDefault();
      const [selectedCols, selectedRows] = sizeText.text().split(' X ').map((n) => parseInt(n, 10));
      createTable(selectedRows, selectedCols);
      closeMenus();
      return;
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

module.exports = {
  enhanceTableHtml,
  enhanceTableMarkup,
  setupMenuAccessibility,
  toDomElement,
};
