const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

async function convertSvgToPng() {
  console.log('🚀 ロボット軍曹アイコンのPNG変換を開始します...');

  const browser = await puppeteer.launch();
  const page = await browser.newPage();

  const iconsPath = 'src/edge-extension/assets/icons';
  const svgFiles = [
    { file: 'icon128.svg', size: 128 },
    { file: 'icon48.svg', size: 48 },
    { file: 'icon32.svg', size: 32 },
    { file: 'icon16.svg', size: 16 }
  ];

  for (const svgFile of svgFiles) {
    const svgPath = path.join(iconsPath, svgFile.file);
    const pngFile = svgFile.file.replace('.svg', '.png');
    const pngPath = path.join(iconsPath, pngFile);

    if (fs.existsSync(svgPath)) {
      try {
        console.log(`🔄 変換中: ${svgFile.file} -> ${pngFile}`);

        // SVGファイル読み込み
        const svgContent = fs.readFileSync(svgPath, 'utf8');

        // HTMLページ作成
        const html = `
                <!DOCTYPE html>
                <html>
                <head>
                    <style>
                        body { margin: 0; padding: 0; background: transparent; }
                        svg { display: block; }
                    </style>
                </head>
                <body>
                    ${svgContent}
                </body>
                </html>
                `;

        await page.setContent(html);
        await page.setViewport({
          width: svgFile.size,
          height: svgFile.size,
          deviceScaleFactor: 1
        });

        // PNG として スクリーンショット
        await page.screenshot({
          path: pngPath,
          omitBackground: true,
          clip: {
            x: 0,
            y: 0,
            width: svgFile.size,
            height: svgFile.size
          }
        });

        console.log(`✅ ${pngFile} の変換が完了しました`);
      } catch (error) {
        console.error(`❌ ${pngFile} の変換に失敗しました:`, error.message);
      }
    } else {
      console.warn(`⚠️  SVGファイルが見つかりません: ${svgPath}`);
    }
  }

  await browser.close();
  console.log('🎉 PNG変換が完了しました！');
  console.log(`📁 生成されたPNGファイルを確認してください: ${iconsPath}`);
}

convertSvgToPng().catch(console.error);
