/**
 * ファイル名: injectPptx.js
 * 説明: 拡張機能がページに埋め込んだPptxGenJSコードを実行して、
 *       PowerPointファイルを自動的に生成・ダウンロードするスクリプト。
 */

(() => {
    // ページコンテキスト上に set されているはず
    const code = window.__SCRAPED_PPTX_CODE__;
    if (!code) return;
    try {
      const pptx = new PptxGenJS();
      // scrapedPptxCode を実行
      eval(code);
      // 保存
      pptx.writeFile({ fileName: 'presentation.pptx' });
    } catch (e) {
      console.error('PPTX generation failed:', e);
      alert('PPTX 生成中にエラーが発生しました');
    } finally {
      // 後始末
      delete window.__SCRAPED_PPTX_CODE__;
    }
  })();
  