// ep_tables5 – collector hooks (legacy shim)
// ------------------------------------------------
// In the new attribute‑based design we no longer embed JSON fragments or
// delimiter labels ("tblBreak", "payload", etc.) directly in the line text, so
// most of the old *collectContent* overrides are obsolete.
//
// We keep a *very* thin shim so Etherpad still finds the hook symbols and
// author spans continue to round‑trip unchanged.  Everything else defers to
// Etherpad's built‑in collector.

// *** DISABLED: These hooks interfere with the new client_hooks.js logic ***

/* global $ */

const log = (...m) => console.debug('[ep_tables5:collector_hooks]', ...m);

// ────────────────────────────────────────────────────────────────────────────
// Dummy Hooks - Do Nothing
// ────────────────────────────────────────────────────────────────────────────
exports.collectContentPre = (hook, ctx) => {
    // log('Collector collectContentPre: Dummy Hook - Doing nothing.');
    return; // Allow default processing
};

exports.collectContentLineBreak = (hook, ctx) => {
    // log('Collector collectContentLineBreak: Dummy Hook - Allowing default.');
    return true; // Allow default processing
};

exports.collectContentLineText = (_hook, ctx) => {
    // log('Collector collectContentLineText: Dummy Hook - Returning default text.');
    return ctx.text || ''; // Return whatever text Etherpad found
};

// ────────────────────────────────────────────────────────────────────────────
// 1. Handle TD elements: Prevent default processing for the TD tag itself.
// ────────────────────────────────────────────────────────────────────────────
// exports.collectContentPre = (hook, ctx) => { ... }; // Previous TD logic

// ────────────────────────────────────────────────────────────────────────────
// 2. Handle inner cell span: Explicitly allow default processing.
// (This might be redundant if the TD hook allows children, but added for clarity).
// We might not even need this specific hook if the TD hook works as intended.
// ────────────────────────────────────────────────────────────────────────────
// exports.collectContentPre = (hook, ctx) => { ... }; // Previous inner span logic

// ────────────────────────────────────────────────────────────────────────────
// 3. Line break / Line Text hooks - Keep DISABLED
// ────────────────────────────────────────────────────────────────────────────
// exports.collectContentLineBreak = (hook, ctx) => { ... };
// exports.collectContentLineText = (_hook, ctx) => { ... };
  