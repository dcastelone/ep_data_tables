'use strict';

const collectContentPre = require('./clientHooks/collectContentPre');
const aceAttribsToClasses = require('./clientHooks/aceAttribsToClasses');
const acePostWriteDomLineHTML = require('./clientHooks/acePostWriteDomLineHTML');
const aceKeyEvent = require('./clientHooks/aceKeyEvent');
const aceInitialized = require('./clientHooks/aceInitialized');
const aceStartLineAndCharForPoint = require('./clientHooks/aceStartLineAndCharForPoint');
const aceEndLineAndCharForPoint = require('./clientHooks/aceEndLineAndCharForPoint');
const aceSetAuthorStyle = require('./clientHooks/aceSetAuthorStyle');
const aceEditorCSS = require('./clientHooks/aceEditorCSS');
const aceRegisterBlockElements = require('./clientHooks/aceRegisterBlockElements');
const aceUndoRedo = require('./clientHooks/aceUndoRedo');

exports.collectContentPre = collectContentPre;
exports.aceAttribsToClasses = aceAttribsToClasses;
exports.acePostWriteDomLineHTML = acePostWriteDomLineHTML;
exports.aceKeyEvent = aceKeyEvent;
exports.aceInitialized = aceInitialized;
exports.aceStartLineAndCharForPoint = aceStartLineAndCharForPoint;
exports.aceEndLineAndCharForPoint = aceEndLineAndCharForPoint;
exports.aceSetAuthorStyle = aceSetAuthorStyle;
exports.aceEditorCSS = aceEditorCSS;
exports.aceRegisterBlockElements = aceRegisterBlockElements;
exports.aceUndoRedo = aceUndoRedo;
