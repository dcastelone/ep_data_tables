'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const {buildTimesliderHtml} = require('../static/js/datatables-renderer');

test('legacy first rows render native column headers without changing width precision', () => {
  const widths = [33.333333333333336, 66.66666666666666];
  const html = buildTimesliderHtml(
      {tblId: 'legacy', row: 0, cols: 2, columnWidths: widths},
      ['Activity', 'Status']);
  assert.match(html, /<th scope="col"[^>]*width: 33\.333333333333336%;/);
  assert.match(html, /<th scope="col"[^>]*width: 66\.66666666666666%;/);
  assert.doesNotMatch(html, /<td/);
});

test('data rows remain data cells and retain rich inline markup', () => {
  const html = buildTimesliderHtml(
      {tblId: 'legacy', row: 1, cols: 2, columnWidths: [40, 60]},
      ['<span><strong>Review</strong></span>', '<span><a href="/details">Complete</a></span>']);
  assert.match(html, /<td[^>]*width: 40%;[^>]*>.*<strong>Review<\/strong>/);
  assert.match(html, /<td[^>]*width: 60%;[^>]*>.*<a href="\/details">Complete<\/a>/);
  assert.doesNotMatch(html, /<th/);
});

test('authors can explicitly opt a first row out of header semantics', () => {
  const html = buildTimesliderHtml(
      {tblId: 'layout-like', row: 0, cols: 2, headerRows: 0}, ['Left', 'Right']);
  assert.match(html, /<td/);
  assert.doesNotMatch(html, /<th/);
});

test('row header metadata creates scoped row headers', () => {
  const html = buildTimesliderHtml(
      {tblId: 'schedule', row: 2, cols: 3, headerColumns: 1},
      ['Monday', 'Algebra', 'Biology']);
  assert.match(html, /<th scope="row"[^>]*>.*Monday/);
  assert.equal((html.match(/<td/g) || []).length, 2);
});

test('malformed metadata follows the existing safe error rendering path', () => {
  assert.match(buildTimesliderHtml({}, ['value']), /dataTable-error/);
});
