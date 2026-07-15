'use strict';

const assert = require('node:assert/strict');
const Module = require('node:module');
const test = require('node:test');

const originalLoad = Module._load;
Module._load = function(request, parent, isMain) {
  if (request === 'ep_etherpad-lite/node/utils/Settings') return {toolbar: {}};
  return originalLoad.call(this, request, parent, isMain);
};
const {collectContentPre} = require('../collectContentPre');
Module._load = originalLoad;

const encode = (value) => Buffer.from(value).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

test('round-trips table metadata from a mixed class list', () => {
  const metadata = JSON.stringify({tblId: 'table-1', row: 0, cols: 3, widths: [20, 30, 50]});
  const calls = [];
  const state = {};
  collectContentPre('collectContentPre', {
    cls: `character bold tbljson-${encode(metadata)} tblCell-0`,
    state,
    cc: {doAttrib: (...args) => calls.push(args)},
  });
  assert.deepEqual(calls, [[state, `tbljson::${metadata}`]]);
});

test('ignores content without table metadata', () => {
  let calls = 0;
  for (const cls of [undefined, '', 'character tblCell-0']) {
    collectContentPre('', {cls, state: {}, cc: {doAttrib: () => { calls += 1; }}});
  }
  assert.equal(calls, 0);
});
