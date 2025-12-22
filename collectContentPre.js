'use strict';

// Ensure settings.toolbar exists early to avoid load-order error with ep_font_color
const settings = require('ep_etherpad-lite/node/utils/Settings');
if (!settings.toolbar) settings.toolbar = {};

const log = (...m) => console.log('[ep_data_tables:collectContentPre_SERVER]', ...m);

// Base64 decoding function for URL-safe variant
const dec = (s) => {
  const str = s.replace(/-/g, '+').replace(/_/g, '/');
  try {
    return Buffer.from(str, 'base64').toString('utf8');
  } catch (e) {
    log('ERROR decoding base64:', s, e);
    return null;
  }
};

/**
 * Server-side collectContentPre hook for import/export operations.
 * Processes tbljson-* CSS classes to reapply table metadata attributes.
 * 
 * Context parameters provided by Etherpad:
 * - cc: Content collector object with doAttrib() method
 * - state: Current collection state
 * - tname: Tag name (may be undefined)
 * - styl: Style attribute
 * - cls: Class attribute string
 * 
 * Note: context.node is NOT provided by Etherpad's collectContentPre hook.
 */
exports.collectContentPre = (hookName, context) => {
  const cls = context.cls;
  const state = context.state;

  if (!cls) return;

  // Find tbljson-* class
  const classes = cls.split(' ');
  let encodedJsonMetadata = null;

  for (const c of classes) {
    if (c.startsWith('tbljson-')) {
      encodedJsonMetadata = c.substring(8);
      break;
    }
  }

  if (!encodedJsonMetadata) return;

  // Decode and apply the table metadata attribute
  try {
    const decodedJsonString = dec(encodedJsonMetadata);
    if (!decodedJsonString) {
      log('ERROR: Decoded JSON string is null or empty');
      return;
    }

    const attribToApply = `tbljson::${decodedJsonString}`;
    context.cc.doAttrib(state, attribToApply);
  } catch (e) {
    log('ERROR applying tbljson attribute:', e);
  }
};
