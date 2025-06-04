/**
 * ファイル名: content_script.js
 * 説明: Webページに「Preview」ボタンを追加し、HTML プレビューを次の HTML が見つかるまで保持し、
 *       PptxGenJS コードのダウンロード機能を併せて提供する拡張機能。
 */

(() => {
  const PANEL_ID    = 'custom-preview-panel';
  const STYLE_ID    = 'custom-preview-style';
  const TOGGLE_ID   = 'custom-preview-toggle';
  const ACTIVE_CLS  = 'custom-preview-active';
  const PANEL_W     = '60%';
  const BOX_ID      = 'pptx-sandbox-frame';

  let scrapedCode    = '';    // 最新の PptxGenJS コード
  let lastHtml       = '';    // 直近に表示した HTML
  let lastPptx       = '';    // 直近に検出した PptxGenJS
  let observer       = null;  // MutationObserver
  let renderTimer    = null;  // デバウンス用タイマー
  let downloadIframe = null;  // ダウンロード用 iframe

  /** ① CSSをページに挿入 */
  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const css = `
      #${PANEL_ID} {
        position: fixed;
        top: 0;
        right: 0;
        width: ${PANEL_W};
        height: 100vh;
        background: #f5f5f5;
        box-shadow: -2px 0 10px rgba(0,0,0,.2);
        z-index: 9999;
        display: flex;
        flex-direction: column;
      }
      body.${ACTIVE_CLS} {
        margin-right: ${PANEL_W};
        overflow: hidden;
      }
      #${TOGGLE_ID} {
        position: fixed;
        top: 50%;
        right: 0;
        transform: translateY(-50%);
        background: #0078d4;
        color: #fff;
        border: none;
        border-top-left-radius: 5px;
        border-bottom-left-radius: 5px;
        padding: 8px;
        cursor: pointer;
        z-index: 9998;
      }
      #${TOGGLE_ID}.open {
        right: ${PANEL_W};
      }
      #${PANEL_ID} .preview-header {
        padding: 10px;
        border-bottom: 1px solid #ccc;
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #eee;
      }
      #${PANEL_ID} .preview-header button {
        margin-left: 8px;
        padding: 4px 8px;
        font-size: 14px;
        cursor: pointer;
      }
      #${PANEL_ID} .preview-header button.active {
        background: #4CAF50;
        color: #fff;
        font-weight: bold;
      }
      #${PANEL_ID} #preview-content {
        flex: 1;
        position: relative;
        background: #fff;
        display: flex;
        justify-content: center;
        align-items: center;
      }
    `;
    const styleEl = Object.assign(document.createElement('style'), {
      id: STYLE_ID,
      textContent: css
    });
    document.head.appendChild(styleEl);
  }

  /** ② 「Preview」ボタンをページに設置 */
  function injectToggle() {
    if (document.getElementById(TOGGLE_ID)) return;
    injectStyles();
    const button = Object.assign(document.createElement('button'), {
      id: TOGGLE_ID,
      textContent: 'Preview'
    });
    button.onclick = togglePanel;
    document.body.appendChild(button);
  }

  function togglePanel() {
    const isOpen = !!document.getElementById(PANEL_ID);
    document.getElementById(TOGGLE_ID).classList.toggle('open', !isOpen);
    isOpen ? closePanel() : openPanel();
  }

  /** ③ パネルを閉じる */
  function closePanel() {
    document.getElementById(PANEL_ID)?.remove();
    document.body.classList.remove(ACTIVE_CLS);
    if (observer) {
      observer.disconnect();
      observer = null;
    }
    clearTimeout(renderTimer);
    // HTML プレビューは次の HTML が来るまで保持したままにするので、lastHtml は消しません
    lastPptx = '';
    scrapedCode = '';
  }

  /** ④ パネルを開く */
  function openPanel() {
    document.body.classList.add(ACTIVE_CLS);

    const panel = document.createElement('div');
    panel.id = PANEL_ID;
    panel.innerHTML = `
      <div class="preview-header">
        <strong>Preview</strong>
        <div>
          <button id="download-btn" disabled>Download PPTX</button>
          <button id="close-btn">×</button>
        </div>
      </div>
      <div id="preview-content"></div>
    `;
    document.body.appendChild(panel);

    panel.querySelector('#close-btn').onclick    = togglePanel;
    panel.querySelector('#download-btn').onclick = downloadPptx;

    renderPreview();       // 初回描画
    startObserveChanges(); // コード変化監視開始
  }

  /** ⑤ ページ内のコードブロックを抽出 (HTML と PPTX を個別取得) */
  function extractBlocks() {
    const nodes = [...document.querySelectorAll('code.language-html, code.language-js, pre')];
    if (!nodes.length) return { error: 'コードブロックが見当たりません' };

    let htmlText = '', pptxText = '';
    nodes.forEach(el => {
      const txt = (el.innerText || el.textContent).trim();
      if (/<[^>]+>/.test(txt)) {
        htmlText = txt;
      }
      if (/(new\s+PptxGenJS|\.addSlide|\bPptxGenJS\b|\.slide\(|\.addText\(|\.addImage\(|\.addShape\()/.test(txt)) {
        pptxText = txt;
      }
    });

    if (!htmlText && !pptxText) {
      return { error: '対象コードがありません' };
    }
    return { html: htmlText, pptx: pptxText };
  }

  /** ⑥ プレビューとボタン状態を更新 */
  function renderPreview() {
    const wrap = document.getElementById(PANEL_ID);
    if (!wrap) return;

    const view   = wrap.querySelector('#preview-content');
    const dlBtn  = wrap.querySelector('#download-btn');
    const { html, pptx, error } = extractBlocks();

    // —— HTML と PPTX が両方ない場合：何もしない（プレビュー保持） —— 
    if (error) {
      dlBtn.disabled = true;
      dlBtn.classList.remove('active');
      return;
    }

    // —— HTML 部分が新しくあればプレビューを差し替え —— 
    if (html && html !== lastHtml) {
      lastHtml = html;
      view.innerHTML = '';
      const w = wrap.clientWidth;
      const h = w * 9 / 16;
      Object.assign(view.style, { width: `${w}px`, height: `${h}px` });

      const frame = document.createElement('iframe');
      frame.srcdoc = html;
      Object.assign(frame.style, {
        position: 'absolute',
        top: 0,
        left: 0,

        width: '1280px',
        height: '1000px',
        transformOrigin: 'top left',
        transform: `scale(${w / 1280})`
      });
      view.appendChild(frame);
    }

    // —— PptxGenJS 部分が新しく検出されたらダウンロードボタンを有効化 —— 
    if (pptx && pptx !== lastPptx) {
      lastPptx    = pptx;
      scrapedCode = pptx;
      dlBtn.disabled = false;
      dlBtn.classList.add('active');
    }

    // —— PPTX コードが見当たらないときはボタンを無効化 —— 
    if (!pptx) {
      scrapedCode = '';
      dlBtn.disabled = true;
      dlBtn.classList.remove('active');
    }
  }

  /** ⑦ PPTX ダウンロード処理 (iframe を再利用) */
  function downloadPptx() {
    if (!scrapedCode) return;
    const btn = document.getElementById('download-btn');
    btn.disabled = true;
    btn.classList.remove('active');

    if (!downloadIframe) {
      downloadIframe = document.createElement('iframe');
      downloadIframe.id = BOX_ID;
      downloadIframe.style.display = 'none';
      downloadIframe.setAttribute('sandbox', 'allow-scripts');
      downloadIframe.src = chrome.runtime.getURL('sandbox/pptx-sandbox.html');
      document.body.appendChild(downloadIframe);
      downloadIframe.onload = () => {
        postToIframe(scrapedCode);
      };
    } else {
      postToIframe(scrapedCode);
    }

    function postToIframe(code) {
      window.removeEventListener('message', handleDownload);
      window.addEventListener('message', handleDownload);
      downloadIframe.contentWindow.postMessage(
        { action: 'generate', code: code, fileName: 'presentation.pptx' },
        '*'
      );
    }

    function handleDownload(e) {
      const m = e.data || {};
      if (m.action === 'done') {
        saveDataURL(m.dataURL, m.fileName);
        cleanup();
      }
      if (m.action === 'error') {
        alert('PPTX 生成失敗: ' + m.message);
        cleanup();
      }
    }

    function cleanup() {
      window.removeEventListener('message', handleDownload);
      btn.disabled = false;
      btn.classList.add('active');
    }

    function saveDataURL(url, name) {
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }
  }

  /** ⑧ コード変化を監視 (デバウンス付き) */
  function startObserveChanges() {
    observer = new MutationObserver(() => {
      clearTimeout(renderTimer);
      renderTimer = setTimeout(renderPreview, 300);
    });
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true
    });
  }

  /** ⑨ 起動時にトグルボタン設置 */
  document.readyState === 'loading'
    ? window.addEventListener('DOMContentLoaded', injectToggle)
    : injectToggle();

})();
