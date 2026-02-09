(function () {
  'use strict';

  function log(message, data) {
    try {
      if (!window.__reportsLog) window.__reportsLog = [];
      window.__reportsLog.push({ time: Date.now(), message: message, data: data || null });
      if (window.__reportsLog.length > 500) {
        window.__reportsLog = window.__reportsLog.slice(window.__reportsLog.length - 500);
      }
      if (window.console && console.log) {
        console.log('[reports-bridge]', message, data || '');
      }
    } catch (e) {
      // ignore
    }
  }

  (function initSystemMessageBridge() {
    try {
      var prev = window.onSystemMessage;
      window.onSystemMessage = function (e) {
        try {
          if (e && e.type === 'reports:debug') {
            log('runtime', e.data || { msg: e.msg });
          }
          if (e && e.type === 'reports:editor-frame') {
            if (e.frameId) {
              window.__reportsEditorFrameId = e.frameId;
              window.__reportsEditorHref = e.href || '';
              log('editor frame', { frameId: e.frameId, href: e.href || '' });
            }
          }
        } catch (err) {
          // ignore
        }
        if (typeof prev === 'function') {
          try { prev(e); } catch (err2) { /* ignore */ }
        }
      };
    } catch (e) {
      // ignore
    }
  })();

  function parseMessage(payload) {
    try {
      return (typeof payload === 'string') ? JSON.parse(payload) : payload;
    } catch (e) {
      return null;
    }
  }

  function buildRunScript(job, debug) {
    var jobStr = JSON.stringify(job || {});
    var debugFlag = debug ? 'true' : 'false';
    return "function(){try{" +
      "var job=JSON.parse(" + JSON.stringify(jobStr) + ");" +
      "job.debug=" + debugFlag + ";" +
      "var msg=JSON.stringify({type:'onExternalPluginMessage',data:{type:'reports:run',job:job}});" +
      "var sent=0;" +
      "try{window.postMessage(msg,'*');sent++;}catch(e){}" +
      "try{if(window.frames&&window.frames.length){for(var i=0;i<window.frames.length;i++){try{window.frames[i].postMessage(msg,'*');sent++;}catch(e){}}}}catch(e){}" +
      "try{if(window.AscDesktopEditor&&window.AscDesktopEditor.sendSystemMessage){window.AscDesktopEditor.sendSystemMessage({type:'reports:debug',data:{msg:'reports:run postMessage sent='+sent+' url='+(window.location&&window.location.href||''),jobId:job.id||''}});}}catch(e){}" +
      "}catch(e){try{if(window.AscDesktopEditor&&window.AscDesktopEditor.sendSystemMessage){window.AscDesktopEditor.sendSystemMessage({type:'reports:debug',data:{msg:'reports:run error '+e,jobId:''}});}}catch(_e){}}}";
  }

  function buildFrameRunScript(job, debug) {
    var jobStr = JSON.stringify(job || {});
    var debugFlag = debug ? 'true' : 'false';
    return "function(){try{" +
      "var job=JSON.parse(" + JSON.stringify(jobStr) + ");" +
      "job.debug=" + debugFlag + ";" +
      "var msg=JSON.stringify({type:'onExternalPluginMessage',data:{type:'reports:run',job:job}});" +
      "var sent=0;" +
      "try{window.postMessage(msg,'*');sent++;}catch(e){}" +
      "try{if(window.frames&&window.frames.length){for(var i=0;i<window.frames.length;i++){try{window.frames[i].postMessage(msg,'*');sent++;}catch(e){}}}}catch(e){}" +
      "try{if(window.AscDesktopEditor&&window.AscDesktopEditor.sendSystemMessage){window.AscDesktopEditor.sendSystemMessage({type:'reports:debug',data:{msg:'reports:run postMessage(frame) sent='+sent+' url='+(window.location&&window.location.href||''),jobId:job.id||''}});}}catch(e){}" +
      "}catch(e){}}";
  }

  function runJob(data) {
    if (!data) return;
    var job = data.job || {};
    var templateId = data.templateId || job.templateId || null;
    log('reportsRun received', { path: data.path, jobId: job.id || '', typeId: data.typeId || 0, debug: !!data.debug });

    try {
      if (window.sdk && data.path) {
        log('create:new', { path: data.path });
        var payload = JSON.stringify({
          template: {
            id: templateId,
            type: data.typeId || 0,
            path: data.path
          }
        });
        log('create:new payload', payload);
        window.sdk.command('create:new', payload);
      }
    } catch (e) {
      log('create:new error', { error: String(e) });
      // ignore
    }

    if (!window.AscDesktopEditor || !window.AscDesktopEditor.CallInAllWindows)
      return;

    var script = buildRunScript(job, !!data.debug);
    var start = Date.now();

    function sendScript() {
      try {
        var frameId = window.__reportsEditorFrameId;
        if (window.AscDesktopEditor && frameId && window.AscDesktopEditor.CallInFrame) {
          var frameScript = buildFrameRunScript(job, !!data.debug);
          log('CallInFrame invoked', { jobId: job.id || '', frameId: frameId });
          window.AscDesktopEditor.CallInFrame(frameId, frameScript);
          return;
        }
        if (typeof script === 'string') {
          log('CallInAllWindows invoked', { jobId: job.id || '' });
          window.AscDesktopEditor.CallInAllWindows(script);
        }
      } catch (e) {
        log('sendScript error', { error: String(e) });
        // ignore
      }
    }

    sendScript();
    var timer = setInterval(function () {
      sendScript();
      if (Date.now() - start > 60000) {
        clearInterval(timer);
      }
    }, 1000);
  }

  window.addEventListener('message', function (evt) {
    var msg = parseMessage(evt.data);
    try {
      if (msg) log('message received', msg);
    } catch (e) { /* ignore */ }
    if (msg && msg.type === 'reports:debug' && msg.data) {
      log('runtime', msg.data);
      return;
    }
    if (!msg || msg.event !== 'reportsRun' || msg.source !== 'reports-ui') return;
    runJob(msg.data);
  }, false);
})();
