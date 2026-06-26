/**
 * 模拟进度条工具
 * 用法:
 *   var pg = createProgress('loading');  // 传入 loading div 的 id
 *   pg.start();                          // 开始模拟进度
 *   // 当数据加载完成后:
 *   pg.finish();                         // 跳到100%并隐藏
 */
function createProgress(loadingId) {
    var el = document.getElementById(loadingId);
    var label = el.querySelector('.loading-label');
    var textEl = el.querySelector('.progress-text');
    var barEl = el.querySelector('.progress-bar-inner');
    var current = 0;
    var timer = null;

    function update(val, stage) {
        current = val;
        if (textEl) textEl.innerText = Math.round(current) + '%/100%';
        if (barEl) barEl.style.width = Math.round(current) + '%';
        if (stage && label) label.innerText = stage;
    }

    function start() {
        current = 0;
        update(0);
        timer = setInterval(function() {
            if (current < 30) {
                current += Math.random() * 5 + 2;
            } else if (current < 60) {
                current += Math.random() * 3 + 1;
            } else if (current < 85) {
                current += Math.random() * 1.5 + 0.3;
            } else if (current < 92) {
                current += Math.random() * 0.3 + 0.1;
            }
            if (current > 92) current = 92;
            update(current);
        }, 300);
    }

    function finish(callback) {
        if (timer) { clearInterval(timer); timer = null; }
        update(100);
        setTimeout(function() {
            el.style.display = 'none';
            if (callback) callback();
        }, 400);
    }

    return { start: start, finish: finish, update: update };
}

/**
 * SSE 真实进度请求工具
 * 用法:
 *   var pg = createProgress('loading');
 *   sseRequest('/api/keywords', pg, function(result) {
 *       // result 是后端返回的最终数据
 *   }, function(err) {
 *       // 错误处理
 *   });
 */
function sseRequest(url, pg, onSuccess, onError) {
    var es = new EventSource(url);
    es.onmessage = function(event) {
        try {
            var data = JSON.parse(event.data);
            if (data.done) {
                es.close();
                pg.finish(function() {
                    if (data.result && data.result.error) {
                        if (onError) onError(new Error(data.result.error));
                    } else {
                        if (onSuccess) onSuccess(data.result);
                    }
                });
            } else if (data.progress !== undefined) {
                pg.update(data.progress, data.stage);
            }
        } catch(e) {
            // ignore parse errors
        }
    };
    es.onerror = function() {
        es.close();
        pg.finish(function() {
            if (onError) onError(new Error('连接中断'));
        });
    };
}
