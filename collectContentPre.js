'use strict';

// Using console.log for server-side logging, similar to other Etherpad server-side files.
const log = (...m) => console.log('[ep_tables5:collectContentPre_SERVER]', ...m);

// Base64 decoding function (copied from tableImport.js / client_hooks.js for server-side use)
// Ensure it can handle the URL-safe variant if that's what tableImport.js produces.
const dec = (s) => {
    const str = s.replace(/-/g, '+').replace(/_/g, '/');
    try {
        // Assuming Buffer is available in the Node.js environment where this server hook runs
        return Buffer.from(str, 'base64').toString('utf8');
    } catch (e) {
        log('ERROR decoding base64 string in server-side collectContentPre:', s, e);
        return null;
    }
};

exports.collectContentPre = (hookName, context) => {

  const node = context.node;
  const tname = node?.tagName?.toLowerCase(); // Keep for logging, but don't rely on it for main logic
  const cls = context.cls;
  const state = context.state;

  // Detailed logging for every node encountered by this hook
  let nodeRepresentation = '';
  if (node) {
    if (node.nodeType === 1 /* Node.ELEMENT_NODE */) {
      nodeRepresentation = node.outerHTML ? node.outerHTML.substring(0, 300) : `ElementType: ${tname}`;
    } else if (node.nodeType === 3 /* Node.TEXT_NODE */) {
      nodeRepresentation = `TextNode: "${(node.nodeValue || '').substring(0, 100)}"`;
    }
  }
  log(`PROCESSING Node: ${nodeRepresentation}, TagName: ${tname}, Classes: "${cls || ''}", Current state.attribString: "${state.attribString || ''}"`);

  // MODIFIED CONDITION: Rely primarily on context.cls to find our target class,
  // as context.node and context.tname seem unreliable for our spans.
  if (!cls) { 
    // log('No classes, returning.');
    return;
  }

  const classes = cls.split(' ');
  let encodedJsonMetadata = null;
  let isTblJsonSpan = false;

  for (const c of classes) {
    if (c.startsWith('tbljson-')) {
      encodedJsonMetadata = c.substring(8);
      isTblJsonSpan = true;
      break;
    }
  }

  if (isTblJsonSpan && encodedJsonMetadata) {
    // Added a check for tname === 'span' here for sanity, though cls is primary trigger
    log(`FOUND potential tbljson span based on class. Actual tname: ${tname}. Encoded: ${encodedJsonMetadata.substring(0, 20)}...`);
    let decodedJsonString;
    try {
      decodedJsonString = dec(encodedJsonMetadata);
      if (!decodedJsonString) {
        throw new Error('Decoded JSON string is null or empty.');
      }
      log(`DECODED metadata: ${decodedJsonString.substring(0, 50)}...`);
    } catch (e) {
      log('ERROR decoding/validating metadata from class:', encodedJsonMetadata, e);
      return;
    }

    const attribToApply = `tbljson::${decodedJsonString}`;
    log(`PREPARED attribToApply: "${attribToApply.substring(0, 70)}..."`);

    log('BEFORE cc.doAttrib - state.attribs:', JSON.stringify(state.attribs));
    log('BEFORE cc.doAttrib - state.attribString:', state.attribString);
    try {
      context.cc.doAttrib(state, attribToApply);
      log('AFTER cc.doAttrib - state.attribs:', JSON.stringify(state.attribs)); 
      log('AFTER cc.doAttrib - state.attribString:', state.attribString);
      log('SUCCESSFULLY CALLED cc.doAttrib.');
    } catch (e) {
      log('ERROR calling cc.doAttrib:', e);
    }
  } else if (cls.includes('tbljson-')) {
      log(`WARN: A class string "${cls}" includes 'tbljson-' but didn't parse as expected, or encodedJsonMetadata was null.`);
  }
}; 