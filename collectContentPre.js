'use strict';

const settings = require('ep_etherpad-lite/node/utils/Settings');
if (!settings.toolbar) settings.toolbar = {};

const ATTR_TABLE_JSON = 'tbljson';
const TBLJSON_CLASS_RE = /\btbljson-([A-Za-z0-9_-]+)/;
const LOG_PREFIX = '[ep_data_tables:collectContentPre]';

const toClassTokens = (classSource) => {
  if (!classSource) return [];
  if (typeof classSource === 'string') return classSource.trim().split(/\s+/).filter(Boolean);
  if (Array.isArray(classSource)) return classSource.filter(Boolean);
  if (typeof classSource === 'object' && typeof classSource.length === 'number') {
    return Array.from(classSource, (token) => `${token}`).filter(Boolean);
  }
  return [];
};

const extractEncodedTbljson = (classSource) => {
  for (const token of toClassTokens(classSource)) {
    const match = TBLJSON_CLASS_RE.exec(token);
    if (match) return match[1];
  }
  return null;
};

const decodeTbljsonClass = (classSource) => {
  const encoded = extractEncodedTbljson(classSource);
  if (!encoded) return null;
  const normalized = encoded.replace(/-/g, '+').replace(/_/g, '/');
  let json;
  try {
    json = Buffer.from(normalized, 'base64').toString('utf8');
  } catch (err) {
    console.warn(`${LOG_PREFIX} Failed to decode base64 metadata`, err);
    return null;
  }
  if (!json) return null;
  let metadata = null;
  try {
    metadata = JSON.parse(json);
  } catch (err) {
    console.warn(`${LOG_PREFIX} Decoded metadata is not valid JSON`, err);
  }
  const isWellFormed = (
    metadata &&
    typeof metadata.tblId !== 'undefined' &&
    typeof metadata.row !== 'undefined' &&
    typeof metadata.cols === 'number'
  );
  return {json, metadata, isWellFormed};
};

exports.collectContentPre = (_hookName, context) => {
  const {cls, state, cc} = context;
  if (!cls || !state || !cc) return;

  const info = decodeTbljsonClass(cls);
  if (!info) return;

  if (!info.isWellFormed) {
    console.warn(`${LOG_PREFIX} Applying tbljson attribute with incomplete metadata`, info.metadata);
  }

  try {
    cc.doAttrib(state, `${ATTR_TABLE_JSON}::${info.json}`);
  } catch (err) {
    console.error(`${LOG_PREFIX} Failed to apply ${ATTR_TABLE_JSON} attribute`, err);
  }
};
