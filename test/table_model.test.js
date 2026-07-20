'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const model = require('../static/js/table_model');

test('equal widths are concise and total exactly 100', () => {
  assert.deepEqual(model.equalColumnWidths(3), [33.3333, 33.3333, 33.3334]);
  assert.deepEqual(
      model.equalColumnWidths(6),
      [16.6667, 16.6667, 16.6667, 16.6667, 16.6667, 16.6665]);
  assert.equal(model.equalColumnWidths(7).reduce((sum, value) => sum + value, 0), 100);
});

test('legacy widths render at their original precision without mutation', () => {
  const legacy = [33.333333333333336, 33.333333333333336, 33.33333333333333];
  const metadata = {columnWidths: legacy};
  const rendered = model.columnWidthsForRender(metadata, 3);
  assert.deepEqual(rendered, legacy);
  assert.notEqual(rendered, legacy);
  assert.deepEqual(metadata.columnWidths, legacy);
});

test('intentionally written widths are normalized and bounded', () => {
  const normalized = model.normalizeColumnWidthsForWrite([1, 2, 3], 3);
  assert.deepEqual(normalized, [16.6667, 33.3333, 50]);
  assert.equal(normalized.reduce((sum, value) => sum + value, 0), 100);
  assert.ok(normalized.every((value) => {
    const decimals = String(value).split('.')[1]?.length || 0;
    return decimals <= 4;
  }));
});

test('invalid widths safely fall back to equal columns', () => {
  assert.deepEqual(model.columnWidthsForRender({columnWidths: [50, NaN]}, 2), [50, 50]);
  assert.deepEqual(model.normalizeColumnWidthsForWrite([0, 100], 2), [50, 50]);
  assert.deepEqual(model.equalColumnWidths(0), []);
});

test('metadata writes preserve unknown fields and normalize widths', () => {
  const metadata = model.normalizeMetadataForWrite({
    tblId: 'legacy', row: 4, cols: 3,
    columnWidths: [20.123456789, 30.987654321, 48.88888889],
    futureField: {keep: true},
  }, {row: 5});
  assert.equal(metadata.row, 5);
  assert.deepEqual(metadata.futureField, {keep: true});
  assert.equal(metadata.columnWidths.reduce((sum, value) => sum + value, 0), 100);
  assert.ok(metadata.columnWidths.every((value) => String(value).length < 10));
});

test('collection preserves legacy and future metadata without normalizing it', () => {
  const widths = [33.333333333333336, 66.66666666666666];
  const metadata = model.mergeCollectedMetadata({
    tblId: 'old', row: 0, cols: 2, columnWidths: widths,
    caption: 'Course progress', headerRows: 1, custom: 'preserved',
  }, {tblId: 'old', row: 1, cols: 2});
  assert.deepEqual(metadata.columnWidths, widths);
  assert.equal(metadata.caption, 'Course progress');
  assert.equal(metadata.headerRows, 1);
  assert.equal(metadata.custom, 'preserved');
  assert.equal(metadata.row, 1);
});

test('legacy header defaults are explicit and overridable', () => {
  assert.equal(model.headerRowCount({}), 1);
  assert.equal(model.headerRowCount({headerRows: 0}), 0);
  assert.equal(model.headerRowCount({headerRows: 1}), 1);
  assert.equal(model.headerColumnCount({}), 0);
  assert.equal(model.headerColumnCount({headerColumns: 1}), 1);
  assert.equal(model.tableCaption({caption: '  Assignment   progress  '}), 'Assignment progress');
});

test('table properties are safe, concise, and preserve future metadata', () => {
  const metadata = model.metadataWithTableProperties({
    tblId: 'table-1', row: 2, cols: 3, futureField: {keep: true},
  }, {
    caption: '  Weekly   progress  ',
    headerRows: 0,
    headerColumns: 1,
    columnWidths: [2, 3, 5],
  }, {tblId: 'table-1', row: 2, cols: 3});
  assert.equal(metadata.caption, 'Weekly progress');
  assert.equal(metadata.headerRows, 0);
  assert.equal(metadata.headerColumns, 1);
  assert.deepEqual(metadata.columnWidths, [20, 30, 50]);
  assert.deepEqual(metadata.futureField, {keep: true});
  assert.deepEqual(model.tablePropertiesForEditing(metadata, 3, 4), {
    caption: 'Weekly progress',
    cols: 3,
    columnWidths: [20, 30, 50],
    headerColumns: 1,
    headerRows: 0,
    rows: 4,
  });
});

test('normalization does not accumulate drift across repeated writes', () => {
  for (let columns = 1; columns <= 50; columns++) {
    let widths = Array.from({length: columns}, (_value, index) => index + 1);
    for (let pass = 0; pass < 100; pass++) {
      widths = model.normalizeColumnWidthsForWrite(widths, columns);
      assert.equal(widths.reduce((sum, value) => sum + Math.round(value * 10000), 0), 1000000);
      assert.ok(widths.every((value) => Number.isFinite(value) && value > 0));
    }
  }
});
