/*
 * PTA AI業務支援ツール - ビルド設定（リファクタリング版）
 * Copyright (c) 2024 PTA Development Team
 */

const fs = require('fs');
const path = require('path');
const { build } = require('esbuild');

/**
 * ビルド設定クラス
 */
class PTABuildManager {
    constructor() {
        this.isProduction = process.argv.includes('--prod');
        this.isDevelopment = process.argv.includes('--dev') || !this.isProduction;
        this.isWatch = process.argv.includes('--watch');

        this.buildDir = path.join(__dirname, 'dist');
        this.srcDir = path.join(__dirname);

        console.log(`🏗️  PTA Edge拡張機能ビルド開始`);
        console.log(`📦 モード: ${this.isProduction ? '本番' : '開発'}`);
        console.log(`👀 ウォッチ: ${this.isWatch ? '有効' : '無効'}`);
    }

    /**
     * ビルドを実行
     */
    async runBuild() {
        try {
            // ビルドディレクトリの準備
            await this.prepareBuildDirectory();

            // 静的ファイルのコピー
            await this.copyStaticFiles();

            // JavaScriptファイルのビルド
            await this.buildJavaScript();

            // マニフェストファイルの処理
            await this.processManifest();

            // CSS ファイルの処理
            await this.processCSSFiles();

            // 本番用の最適化
            if (this.isProduction) {
                await this.optimizeForProduction();
            }

            console.log('✅ ビルドが完了しました');
            console.log(`📁 出力先: ${this.buildDir}`);

        } catch (error) {
            console.error('❌ ビルドエラー:', error);
            process.exit(1);
        }
    }

    /**
     * ビルドディレクトリの準備
     */
    async prepareBuildDirectory() {
        if (fs.existsSync(this.buildDir)) {
            fs.rmSync(this.buildDir, { recursive: true, force: true });
        }
        fs.mkdirSync(this.buildDir, { recursive: true });

        console.log('📁 ビルドディレクトリを準備しました');
    }

    /**
     * 静的ファイルのコピー
     */
    async copyStaticFiles() {
        const staticDirs = [
            'assets',
            'vendor',
            'infrastructure',
            'popup',
            'options',
            'offscreen',
        ];

        for (const dir of staticDirs) {
            const srcPath = path.join(this.srcDir, dir);
            const destPath = path.join(this.buildDir, dir);

            if (fs.existsSync(srcPath)) {
                this.copyDirectory(srcPath, destPath);
                console.log(`📋 ${dir} をコピーしました`);
            }
        }
    }

    /**
     * JavaScriptファイルのビルド
     */
    async buildJavaScript() {
        const entryPoints = [
            // 新しいコアモジュール
            'core/constants.js',
            'core/logger.js',
            'core/error-handler.js',
            'core/event-manager.js',
            'core/settings-manager.js',
            'core/api-client.js',

            // メインスクリプト
            'background/background-refactored.js',
            'content/content.js',
        ];

        const buildOptions = {
            entryPoints: entryPoints.map(entry => path.join(this.srcDir, entry)),
            bundle: true,
            outdir: this.buildDir,
            format: 'esm',
            target: 'es2020',
            sourcemap: this.isDevelopment,
            minify: this.isProduction,
            keepNames: this.isDevelopment,
            define: {
                'process.env.NODE_ENV': this.isProduction ? '"production"' : '"development"',
                'process.env.BUILD_TIME': `"${new Date().toISOString()}"`,
            },
            banner: {
                js: `/*
 * PTA AI業務支援ツール - ビルド版
 * ビルド日時: ${new Date().toISOString()}
 * モード: ${this.isProduction ? '本番' : '開発'}
 */`,
            },
            loader: {
                '.json': 'json',
            },
        };

        if (this.isWatch) {
            buildOptions.watch = {
                onRebuild(error, result) {
                    if (error) {
                        console.error('⚠️  リビルドエラー:', error);
                    } else {
                        console.log('🔄 リビルド完了:', new Date().toLocaleTimeString());
                    }
                },
            };
        }

        const result = await build(buildOptions);

        if (result.errors.length > 0) {
            console.error('❌ JavaScript ビルドエラー:', result.errors);
            throw new Error('JavaScript ビルドに失敗しました');
        }

        console.log('✅ JavaScript ファイルをビルドしました');
    }

    /**
     * マニフェストファイルの処理
     */
    async processManifest() {
        const manifestPath = path.join(this.srcDir, 'manifest.json');
        const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));

        // 本番用の調整
        if (this.isProduction) {
            // バックグラウンドスクリプトを新しい版に切り替え
            manifest.background.service_worker = 'background/background-refactored.js';

            // デバッグ関連の権限を削除
            if (manifest.permissions.includes('debugger')) {
                manifest.permissions = manifest.permissions.filter(p => p !== 'debugger');
            }
        } else {
            // 開発用の設定
            manifest.name += ' (開発版)';
            manifest.version += '-dev';
        }

        // ビルド情報を追加
        manifest.build_info = {
            build_time: new Date().toISOString(),
            mode: this.isProduction ? 'production' : 'development',
            version: require('./package.json').version,
        };

        fs.writeFileSync(
            path.join(this.buildDir, 'manifest.json'),
            JSON.stringify(manifest, null, 2)
        );

        console.log('✅ マニフェストファイルを処理しました');
    }

    /**
     * CSSファイルの処理
     */
    async processCSSFiles() {
        const cssFiles = [
            'content/content.css',
            'popup/popup.css',
            'options/options.css',
        ];

        for (const cssFile of cssFiles) {
            const srcPath = path.join(this.srcDir, cssFile);
            const destPath = path.join(this.buildDir, cssFile);

            if (fs.existsSync(srcPath)) {
                let cssContent = fs.readFileSync(srcPath, 'utf8');

                // 本番用の最適化
                if (this.isProduction) {
                    cssContent = this.minifyCSS(cssContent);
                }

                // ディレクトリを作成
                fs.mkdirSync(path.dirname(destPath), { recursive: true });
                fs.writeFileSync(destPath, cssContent);
            }
        }

        console.log('✅ CSS ファイルを処理しました');
    }

    /**
     * 本番用最適化
     */
    async optimizeForProduction() {
        console.log('⚡ 本番用最適化を実行中...');

        // 不要なファイルの削除
        const unnecessaryFiles = [
            'tests',
            'docs',
            '.map files',
        ];

        // ファイルサイズの報告
        await this.reportBuildSize();

        console.log('✅ 本番用最適化が完了しました');
    }

    /**
     * ビルドサイズを報告
     */
    async reportBuildSize() {
        const files = this.getAllFiles(this.buildDir);
        let totalSize = 0;

        console.log('\n📊 ビルドサイズレポート:');
        console.log('================================');

        for (const file of files) {
            const stats = fs.statSync(file);
            const relativePath = path.relative(this.buildDir, file);
            const size = (stats.size / 1024).toFixed(1);

            totalSize += stats.size;

            if (stats.size > 1024) { // 1KB以上のファイルのみ表示
                console.log(`  ${relativePath}: ${size} KB`);
            }
        }

        console.log('================================');
        console.log(`  合計サイズ: ${(totalSize / 1024).toFixed(1)} KB`);
        console.log(`  ファイル数: ${files.length}`);
    }

    /**
     * ディレクトリをコピー
     */
    copyDirectory(src, dest) {
        if (!fs.existsSync(dest)) {
            fs.mkdirSync(dest, { recursive: true });
        }

        const entries = fs.readdirSync(src, { withFileTypes: true });

        for (const entry of entries) {
            const srcPath = path.join(src, entry.name);
            const destPath = path.join(dest, entry.name);

            if (entry.isDirectory()) {
                this.copyDirectory(srcPath, destPath);
            } else {
                fs.copyFileSync(srcPath, destPath);
            }
        }
    }

    /**
     * すべてのファイルを取得
     */
    getAllFiles(dir, files = []) {
        const entries = fs.readdirSync(dir, { withFileTypes: true });

        for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);

            if (entry.isDirectory()) {
                this.getAllFiles(fullPath, files);
            } else {
                files.push(fullPath);
            }
        }

        return files;
    }

    /**
     * CSS最小化（簡易版）
     */
    minifyCSS(css) {
        return css
            .replace(/\/\*[\s\S]*?\*\//g, '') // コメント削除
            .replace(/\s+/g, ' ') // 空白の正規化
            .replace(/;\s*}/g, '}') // 最後のセミコロン削除
            .replace(/\s*{\s*/g, '{') // 開き括弧の空白削除
            .replace(/}\s*/g, '}') // 閉じ括弧の空白削除
            .replace(/:\s*/g, ':') // コロンの空白削除
            .replace(/;\s*/g, ';') // セミコロンの空白削除
            .trim();
    }
}

/**
 * ZIP パッケージ作成
 */
async function createZipPackage() {
    const archiver = require('archiver');
    const output = fs.createWriteStream('pta-edge-extension.zip');
    const archive = archiver('zip', { zlib: { level: 9 } });

    return new Promise((resolve, reject) => {
        output.on('close', () => {
            console.log(`📦 ZIP パッケージを作成しました: ${archive.pointer()} bytes`);
            resolve();
        });

        archive.on('error', reject);
        archive.pipe(output);
        archive.directory('dist/', false);
        archive.finalize();
    });
}

/**
 * メイン実行
 */
async function main() {
    const buildManager = new PTABuildManager();

    try {
        await buildManager.runBuild();

        // ZIP作成オプション
        if (process.argv.includes('--zip')) {
            await createZipPackage();
        }

        if (!buildManager.isWatch) {
            console.log('\n🎉 ビルドプロセスが正常に完了しました！');
        } else {
            console.log('\n👀 ファイル変更を監視中... (Ctrl+C で終了)');
        }

    } catch (error) {
        console.error('\n❌ ビルドプロセスでエラーが発生しました:', error);
        process.exit(1);
    }
}

// CLI オプションの表示
if (process.argv.includes('--help')) {
    console.log(`
PTA Edge拡張機能 ビルドツール

使用方法:
  node build.js [オプション]

オプション:
  --dev        開発モードでビルド (デフォルト)
  --prod       本番モードでビルド (最適化有効)
  --watch      ファイル変更を監視してリビルド
  --zip        ビルド後にZIPパッケージを作成
  --help       このヘルプを表示

例:
  node build.js --prod --zip    # 本番ビルド + ZIP作成
  node build.js --dev --watch   # 開発モード + ウォッチ
  `);
    process.exit(0);
}

// メイン実行
if (require.main === module) {
    main();
}

module.exports = { PTABuildManager };
