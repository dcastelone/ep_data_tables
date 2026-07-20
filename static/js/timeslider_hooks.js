'use strict';

// broadcast.ts invokes goToRevisionEvent synchronously before it composes or
// applies a revision path. Keep this bridge deliberately tiny: the direct
// timeslider renderer owns its DOM state, while the plugin hook only gives it
// the last safe moment to restore Etherpad's canonical line markup.
exports.goToRevisionEvent = () => {
  if (typeof window === 'undefined') return;
  window.epDataTablesBeforeTimesliderRevisionChange?.();
};
