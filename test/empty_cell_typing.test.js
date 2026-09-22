'use strict';

const assert = require('node:assert/strict');
const {test} = require('node:test');
const hooks = require('../static/js/client_hooks');

// Exercise the actual published keyboard hook with core's row-boundary contract.
const typeAtCellStart = (cells, column) => {
  const metadata = JSON.stringify({tblId: 'legacy', row: 16, cols: cells.length});
  const line = {text: cells.join('\u241f')};
  const start = cells.slice(0, column).reduce((n, cell) => n + cell.length + 1, 0);
  const rep = {selStart: [0, start], selEnd: [0, start], lines: {atIndex: () => line}};
  const failures = [];
  const editorInfo = {
    editor: {},
    ace_getRep: () => rep,
    ace_performSelectionChange: (from, to) => { rep.selStart = from; rep.selEnd = to; },
    ace_updateBrowserSelectionFromRep: () => {},
    ace_fastIncorp: () => {},
    ep_data_tables_applyMeta: () => {},
    ace_performDocumentReplaceRange: (from, to, text) => {
      assert.ok(from[1] >= 0 && to[1] <= line.text.length,
          'insertion must stay within the stored row');
      line.text = line.text.slice(0, from[1]) + text + line.text.slice(to[1]);
    },
  };
  const error = console.error;
  console.error = (...args) => failures.push(args);
  let prevented = false;
  try {
    hooks.aceKeyEvent('', {
      evt: {type: 'keydown', key: 'X', keyCode: 88, preventDefault: () => { prevented = true; }},
      rep,
      editorInfo,
      documentAttributeManager: {getAttributeOnLine: () => metadata},
    });
  } finally {
    console.error = error;
  }
  assert.deepEqual(failures, [], 'typing must not reach the hook error handler');
  assert.equal(prevented, true);
  return {cells: line.text.split('\u241f'), caret: rep.selStart};
};

test('empty first, middle, and final cells accept a character without crossing boundaries', () => {
  for (const column of [0, 1, 2, 3, 4]) {
    const cells = ['left', 'middle', 'keep', 'safe', 'right'];
    cells[column] = '';
    const expected = cells.slice();
    expected[column] = 'X';
    assert.deepEqual(typeAtCellStart(cells, column).cells, expected);
  }
});

test('an entirely empty five-cell row accepts typing in every cell', () => {
  for (let column = 0; column < 5; column++) {
    const expected = ['', '', '', '', ''];
    expected[column] = 'X';
    assert.deepEqual(typeAtCellStart(['', '', '', '', ''], column).cells, expected);
  }
});

test('existing nonempty-cell start behavior is preserved', () => {
  for (const [before, after] of [[' ', ' X'], ['  ', ' X '], ['AB', 'AXB']]) {
    assert.deepEqual(typeAtCellStart(['left', before], 1).cells, ['left', after]);
  }
  assert.deepEqual(typeAtCellStart(['AB', 'right'], 0).cells, ['XAB', 'right']);
});
