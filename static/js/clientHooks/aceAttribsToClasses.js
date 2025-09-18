'use strict';

const {
  ATTR_TABLE_JSON,
  ATTR_CELL,
  enc,
} = require('./shared');

module.exports = (_hook, ctx) => {
  if (ctx.key === ATTR_TABLE_JSON) {
    const rawJsonValue = ctx.value;

    const className = `tbljson-${enc(rawJsonValue)}`;
    return [className];
  }
  if (ctx.key === ATTR_CELL) {
    return [`tblCell-${ctx.value}`];
  }
  return [];
};
