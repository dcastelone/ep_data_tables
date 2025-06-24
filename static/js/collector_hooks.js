const log = (...m) => console.debug('[ep_data_tables:collector_hooks]', ...m);

exports.collectContentPre = (hook, ctx) => {
    return; 
};

exports.collectContentLineBreak = (hook, ctx) => {
    return true; 
};

exports.collectContentLineText = (_hook, ctx) => {
    return ctx.text || ''; 
};