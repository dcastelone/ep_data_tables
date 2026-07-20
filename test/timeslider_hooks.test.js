'use strict';

const assert = require('node:assert/strict');
const test = require('node:test');
const {goToRevisionEvent} = require('../static/js/timeslider_hooks');

test('restores canonical timeslider markup through the pre-revision bridge', () => {
  let calls = 0;
  global.window = {epDataTablesBeforeTimesliderRevisionChange: () => calls++};
  try {
    assert.equal(goToRevisionEvent(), undefined);
    assert.equal(calls, 1);
  } finally {
    delete global.window;
  }
});

test('is a safe no-op before the direct timeslider renderer initializes', () => {
  global.window = {};
  try {
    assert.doesNotThrow(() => goToRevisionEvent());
  } finally {
    delete global.window;
  }
});
