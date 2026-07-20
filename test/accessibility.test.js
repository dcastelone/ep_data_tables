'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const {enhanceTableMarkup, isEditableDocument, toDomElement} = require('../static/js/accessibility');

const element = () => {
  const attributes = new Map();
  return {attributes, setAttribute: (key, value) => attributes.set(key, value)};
};

test('accepts DOM and jQuery-style elements', () => {
  const node = element();
  assert.equal(toDomElement(node), node);
  assert.equal(toDomElement({jquery: true, 0: node}), node);
  assert.equal(toDomElement(null), null);
});

test('adds table, cell, delimiter, and resize accessibility metadata', () => {
  const cells = [element(), element()];
  const delimiter = element();
  const handle = element();
  const table = element();
  table.querySelectorAll = (selector) => ({
    'td, th': cells,
    '.ep-data_tables-delim, .ep-data_tables-caret-anchor': [delimiter],
    '.ep-data_tables-resize-handle': [handle],
  })[selector] || [];

  enhanceTableMarkup(table, {row: 2});
  assert.equal(table.attributes.get('data-ep-data-tables-accessible'), 'true');
  assert.equal(table.attributes.get('aria-label'), 'Data table');
  assert.equal(table.attributes.get('data-ep-data-tables-header-rows'), '1');
  assert.equal(cells[0].attributes.get('aria-colindex'), '1');
  assert.equal(cells[1].attributes.get('aria-colindex'), '2');
  assert.equal(delimiter.attributes.get('aria-hidden'), 'true');
  assert.equal(handle.attributes.get('tabindex'), '-1');
});

test('is a safe no-op for malformed elements', () => {
  assert.doesNotThrow(() => enhanceTableMarkup(null));
  assert.doesNotThrow(() => enhanceTableMarkup({}));
});

test('recognizes editable ACE documents so deferred decoration can be prohibited', () => {
  assert.equal(isEditableDocument({body: {isContentEditable: true}}), true);
  assert.equal(isEditableDocument({body: {
    isContentEditable: false,
    getAttribute: (name) => name === 'contenteditable' ? 'true' : null,
  }}), true);
  assert.equal(isEditableDocument({body: {isContentEditable: false}, designMode: 'on'}), true);
  assert.equal(isEditableDocument({body: {isContentEditable: false}, designMode: 'off'}), false);
});
