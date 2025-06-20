// index.js – server-side EEJS injections for ep_tables5
const eejs = require('ep_etherpad-lite/node/eejs/');

const log = (...m) => console.debug('[ep_tables5:index]', ...m);
log('Loading EEJS templates...');

const scriptsTag        = eejs.require('ep_tables5/templates/datatablesScripts.ejs');
const editbarButtonsTag = eejs.require('ep_tables5/templates/datatablesEditbarButtons.ejs');
const styleTag          = eejs.require('ep_tables5/templates/styles.ejs'); // Optional base styles
// const timesliderTag  = eejs.require('ep_tables5/templates/datatablesScriptsTimeslider.ejs'); // optional
log('EEJS templates loaded.');

exports.eejsBlock_scripts = (_hook, args) => {
  log('eejsBlock_scripts: START');
  args.content += scriptsTag;
  log('eejsBlock_scripts: Appended scriptsTag.');
  log('eejsBlock_scripts: END');
};

exports.eejsBlock_editbarMenuLeft = (_hook, args) => {
  log('eejsBlock_editbarMenuLeft: START');
  args.content += editbarButtonsTag;
  log('eejsBlock_editbarMenuLeft: Appended editbarButtonsTag.');
  log('eejsBlock_editbarMenuLeft: END');
};

// Reverted: Only prepend styles from template, do not link external CSS file
exports.eejsBlock_styles = (_hook, args) => {
  log('eejsBlock_styles: START');
  // If you rely solely on the inline CSS from client_hooks.js, delete this hook.
  args.content = styleTag + args.content;
  log('eejsBlock_styles: Prepended styleTag (from template).');
  log('eejsBlock_styles: END');
};

/* Uncomment only if you truly need extra JS in the timeslider */

exports.eejsBlock_timesliderScripts = (_hook, args) => {
  log('eejsBlock_timesliderScripts: START');
  args.content +='<script src="/static/plugins/ep_tables5/static/js/datatables-renderer.js"></script>';
  log('eejsBlock_timesliderScripts: Appended timesliderTag.');
  log('eejsBlock_timesliderScripts: END');
};

log('EEJS hooks defined.');
