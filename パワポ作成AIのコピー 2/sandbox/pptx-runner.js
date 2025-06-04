/**
 * ファイル名: sandbox/pptx-runner.js
 * 説明: sandbox 内で受け取ったコードを実行し、.pptx を生成して返すスクリプト。
 * “compression: 'STORE'” を用いて ZIP 圧縮をスキップ。
 */

window.addEventListener('message', async (e) => {
  const { action, code, fileName = 'presentation.pptx', options = {} } = e.data || {};
  if (action !== 'generate') return;

  try {
    // 1) インスタンス生成
    const pptx = new PptxGenJS();

    // 2) カスタムスライドサイズ (16:9)
    pptx.defineLayout({ name: 'CUSTOM_LAYOUT', width: 13.33, height: 7.5 });
    pptx.layout = 'CUSTOM_LAYOUT';

    // 3) スライドマスター定義
    pptx.defineSlideMaster({ title: 'MASTER_SLIDE' });
    // 新規追加: よく使われる変数を事前定義
    let slide = pptx.addSlide();  // デフォルトのスライドを作成
    // 4) ユーザーコード実行
    /* eslint-disable no-eval */
    eval(code);
    /* eslint-enable no-eval */

    // 5) Blob 生成 (compression: 'STORE' で圧縮なし)
    const blob = await pptx.write('blob', { compression: options.compression || 'STORE' });

    // 6) Blob → DataURL
    const fr = new FileReader();
    fr.onload = () => {
      window.parent.postMessage(
        { action: 'done', dataURL: fr.result, fileName },
        '*'
      );
    };
    fr.readAsDataURL(blob);
  } catch (err) {
    window.parent.postMessage({ action: 'error', message: err.message }, '*');
  }
});
