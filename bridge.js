(function () {
  'use strict';

  var pendingRuns = Object.create(null);

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
    if (!payload && payload !== 0 && payload !== false) return null;
    if (typeof payload === 'object') return payload;
    var text = String(payload);
    if (!text) return null;
    try {
      return JSON.parse(text);
    } catch (e) {}
    try {
      var div = document.createElement('div');
      div.innerHTML = text;
      var decoded = div.textContent || div.innerText || '';
      if (decoded && decoded !== text) {
        return JSON.parse(decoded);
      }
    } catch (e2) {}
    return null;
  }

  function extractRequestId(raw) {
    var text = '';
    try {
      text = String(raw || '');
      if (!text) return '';
      var m = /"requestId"\s*:\s*"([^"]+)"/.exec(text);
      if (m && m[1]) return m[1];
      var div = document.createElement('div');
      div.innerHTML = text;
      var decoded = div.textContent || div.innerText || '';
      if (decoded && decoded !== text) {
        m = /"requestId"\s*:\s*"([^"]+)"/.exec(decoded);
        if (m && m[1]) return m[1];
      }
    } catch (e) {
      // ignore
    }
    return '';
  }

  function clearRunWatch(requestId) {
    var item = pendingRuns[requestId];
    if (!item) return;
    if (item.timer) clearInterval(item.timer);
    delete pendingRuns[requestId];
  }

  function openViaDocBuilder(path, typeId) {
    if (!path) return;
    if (!window.sdk || typeof window.sdk.command !== 'function') return;
    var fullPath = String(path || '');
    var requestId = 'docbuilder-open-' + Date.now() + '-' + Math.floor(Math.random() * 100000);
    window.sdk.command('docbuilder:open', JSON.stringify({
      requestId: requestId,
      path: fullPath,
      type: typeId || 257
    }));
  }

  function normalizeFilesCheckedValue(value) {
    var out = value;
    if (typeof out === 'string') {
      try {
        out = JSON.parse(out);
      } catch (e) {
        out = (out === 'true' || out === '1');
      }
    }
    return !!out;
  }

  function handleFilesChecked(param) {
    var payload = parseMessage(param);
    if (!payload || typeof payload !== 'object') return;
    var ids = Object.keys(pendingRuns);
    if (!ids.length) return;

    for (var i = 0; i < ids.length; i += 1) {
      var requestId = ids[i];
      var item = pendingRuns[requestId];
      if (!item) continue;
      if (!Object.prototype.hasOwnProperty.call(payload, item.checkKey)) continue;
      var exists = normalizeFilesCheckedValue(payload[item.checkKey]);
      if (!exists) continue;
      if (item.openAfterRun) {
        clearRunWatch(requestId);
        continue;
      }

      clearRunWatch(requestId);
      broadcastToReportsUi({
        event: 'reportsDocBuilderResult',
        source: 'reports-bridge',
        data: {
          ok: true,
          requestId: requestId,
          exitCode: 0,
          via: 'files:checked',
          outputPath: item.outputPath || ''
        }
      });
    }
  }

  function startRunWatch(payload) {
    if (!payload || typeof payload !== 'object') return;
    var requestId = payload.requestId ? String(payload.requestId) : '';
    if (!requestId) return;
    var outPath = '';
    if (payload.argument && typeof payload.argument === 'object' && payload.argument.outputPath) {
      outPath = String(payload.argument.outputPath);
    }
    if (!outPath) return;

    clearRunWatch(requestId);
    var timeoutMs = Number(payload.timeoutMs || 120000);
    if (!Number.isFinite(timeoutMs) || timeoutMs < 1000) timeoutMs = 120000;
    var checkKey = 'reports_db_' + requestId.replace(/[^a-zA-Z0-9_-]/g, '_');
    var started = Date.now();
    var item = {
      requestId: requestId,
      outputPath: outPath,
      checkKey: checkKey,
      openAfterRun: !!payload.openAfterRun,
      startedAt: started,
      timer: null
    };
    pendingRuns[requestId] = item;

    var checkNow = function () {
      if (!window.sdk || typeof window.sdk.command !== 'function') return;
      if (!pendingRuns[requestId]) return;
      var elapsed = Date.now() - started;
      if (elapsed > timeoutMs + 30000) {
        clearRunWatch(requestId);
        sendDocBuilderError(requestId, 'timeout', 'No completion response from DocumentBuilder bridge.');
        return;
      }
      try {
        var map = {};
        map[checkKey] = outPath;
        window.sdk.command('files:check', JSON.stringify(map));
      } catch (e) {
        // ignore
      }
    };

    item.timer = setInterval(checkNow, 1000);
    checkNow();
  }

  function broadcastToReportsUi(message) {
    var payload = JSON.stringify(message || {});
    try {
      var iframe = document.getElementById('reports-iframe');
      if (iframe && iframe.contentWindow) {
        iframe.contentWindow.postMessage(payload, '*');
      }
    } catch (e) {
      // ignore
    }
    try {
      if (window.frames && window.frames.length) {
        for (var i = 0; i < window.frames.length; i += 1) {
          try {
            window.frames[i].postMessage(payload, '*');
          } catch (e2) {
            // ignore
          }
        }
      }
    } catch (e3) {
      // ignore
    }
  }

  function sendDocBuilderError(requestId, code, details) {
    broadcastToReportsUi({
      event: 'reportsDocBuilderResult',
      source: 'reports-bridge',
      data: {
        ok: false,
        requestId: requestId || '',
        error: code || 'bridge_error',
        details: details || ''
      }
    });
  }

  function bindNativeDocBuilderEvents() {
    try {
      if (window.__reportsDocBuilderBound) return true;
      if (!window.sdk || typeof window.sdk.on !== 'function') return false;
      window.sdk.on('on_native_message', function (cmd, param) {
        var normalizedCmd = String(cmd || '').trim();
        if (normalizedCmd === 'files:checked') {
          handleFilesChecked(param);
          return;
        }
        if (normalizedCmd === 'docbuilder:openResult') {
          var openParsed = parseMessage(param);
          if (!openParsed || typeof openParsed !== 'object') {
            openParsed = { ok: false, error: 'invalid_native_response', raw: param };
          }
          broadcastToReportsUi({
            event: 'reportsDocBuilderOpenResult',
            source: 'reports-bridge',
            data: openParsed
          });
          return;
        }
        if (normalizedCmd !== 'docbuilder:result' && normalizedCmd !== 'docbuilder:probeResult') return;
        var parsed = parseMessage(param);
        if (!parsed || typeof parsed !== 'object') {
          var fallbackRequestId = extractRequestId(param);
          parsed = { ok: false, requestId: fallbackRequestId || undefined, error: 'invalid_native_response', raw: param };
        }
        if (parsed.requestId) {
          clearRunWatch(String(parsed.requestId));
        }
        broadcastToReportsUi({
          event: (normalizedCmd === 'docbuilder:probeResult') ? 'reportsDocBuilderProbeResult' : 'reportsDocBuilderResult',
          source: 'reports-bridge',
          data: parsed
        });
      });
      window.__reportsDocBuilderBound = true;
      return true;
    } catch (e) {
      return false;
    }
  }

  (function ensureNativeDocBuilderBinding() {
    if (bindNativeDocBuilderEvents()) return;
    var retries = 0;
    var timer = setInterval(function () {
      retries += 1;
      if (bindNativeDocBuilderEvents() || retries > 240) {
        clearInterval(timer);
      }
    }, 500);
  })();

  function runDocBuilder(data) {
    bindNativeDocBuilderEvents();
    var payload = (data && data.payload) ? data.payload : {};
    if (!payload || typeof payload !== 'object') {
      payload = {};
    }
    if (!payload.requestId && data && data.requestId) {
      payload.requestId = data.requestId;
    }
    if (!payload.requestId) {
      payload.requestId = 'docbuilder-' + Date.now();
    }

    if (!window.sdk || typeof window.sdk.command !== 'function') {
      sendDocBuilderError(payload.requestId, 'sdk_unavailable', 'Desktop API is unavailable.');
      return;
    }

    try {
      log('docbuilder:run', { requestId: payload.requestId, script: payload.script || '' });
      startRunWatch(payload);
      window.sdk.command('docbuilder:run', JSON.stringify(payload));
    } catch (e) {
      clearRunWatch(payload.requestId);
      sendDocBuilderError(payload.requestId, 'docbuilder_run_error', String(e || 'unknown'));
    }
  }

  function probeDocBuilder(data) {
    bindNativeDocBuilderEvents();
    var payload = (data && data.payload) ? data.payload : {};
    if (!payload || typeof payload !== 'object') {
      payload = {};
    }
    if (!payload.requestId && data && data.requestId) {
      payload.requestId = data.requestId;
    }
    if (!payload.requestId) {
      payload.requestId = 'docbuilder-probe-' + Date.now();
    }

    if (!window.sdk || typeof window.sdk.command !== 'function') {
      broadcastToReportsUi({
        event: 'reportsDocBuilderProbeResult',
        source: 'reports-bridge',
        data: {
          ok: false,
          requestId: payload.requestId,
          error: 'sdk_unavailable'
        }
      });
      return;
    }

    try {
      log('docbuilder:probe', { requestId: payload.requestId });
      window.sdk.command('docbuilder:probe', JSON.stringify(payload));
    } catch (e) {
      broadcastToReportsUi({
        event: 'reportsDocBuilderProbeResult',
        source: 'reports-bridge',
        data: {
          ok: false,
          requestId: payload.requestId,
          error: 'docbuilder_probe_error',
          details: String(e || 'unknown')
        }
      });
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

  function openGeneratedFile(data) {
    if (!data || !data.path) return;
    if (!window.sdk || typeof window.sdk.command !== 'function') return;
    var typeId = Number(data.typeId || 0);
    if (!Number.isFinite(typeId) || typeId <= 0) {
      var ext = String(data.path || '').split('.').pop().toLowerCase();
      try {
        if (window.utils && typeof window.utils.fileExtensionToFileFormat === 'function') {
          typeId = Number(window.utils.fileExtensionToFileFormat(ext) || 0);
        }
      } catch (e) {
        // ignore
      }
      if (!typeId) {
        if (ext === 'xlsx') typeId = 257;
        else if (ext === 'xls') typeId = 258;
        else if (ext === 'xlsm') typeId = 261;
      }
    }
    try {
      log('reportsOpenFile', { path: data.path, typeId: typeId || 0 });
      openViaDocBuilder(data.path, typeId || 257);
    } catch (e) {
      openViaDocBuilder(data.path, typeId || 257);
    }
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
    if (!msg || msg.source !== 'reports-ui') return;
    if (msg.event === 'reportsDocBuilderRun') {
      runDocBuilder(msg.data);
      return;
    }
    if (msg.event === 'reportsDocBuilderProbe') {
      probeDocBuilder(msg.data);
      return;
    }
    if (msg.event === 'reportsOpenFile') {
      openGeneratedFile(msg.data);
      return;
    }
    if (msg.event !== 'reportsRun') return;
    runJob(msg.data);
  }, false);
})();
