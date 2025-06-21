# SVGからPNGへの変換スクリプト
# このスクリプトはInkscapeを使用してSVGファイルをPNGに変換します
# 事前にInkscapeをインストールしてください: https://inkscape.org/

param(
  [string]$IconsPath = "c:\Users\DD20711\Source\Repos\PTA\src\edge-extension\assets\icons"
)

# Inkscapeのパスを設定
$InkscapePath = "C:\Program Files\Inkscape\bin\inkscape.exe"

# Inkscapeがインストールされているかチェック
if (-not (Test-Path $InkscapePath)) {
  Write-Host "Inkscapeが見つかりません。以下からダウンロード・インストールしてください:" -ForegroundColor Red
  Write-Host "https://inkscape.org/release/inkscape-1.3/" -ForegroundColor Yellow
  Write-Host "インストール後、このスクリプトを再実行してください。" -ForegroundColor Yellow
  exit 1
}

Write-Host "ロボット軍曹アイコンのPNG変換を開始します..." -ForegroundColor Green

# SVGファイルのリスト
$svgFiles = @(
  @{File = "icon128.svg"; Size = 128 },
  @{File = "icon48.svg"; Size = 48 }, 
  @{File = "icon32.svg"; Size = 32 },
  @{File = "icon16.svg"; Size = 16 }
)

foreach ($svgFile in $svgFiles) {
  $svgPath = Join-Path $IconsPath $svgFile.File
  $pngFile = $svgFile.File -replace '\.svg$', '.png'
  $pngPath = Join-Path $IconsPath $pngFile
    
  if (Test-Path $svgPath) {
    Write-Host "変換中: $($svgFile.File) -> $pngFile" -ForegroundColor Cyan
        
    # Inkscapeを使用してSVGをPNGに変換
    $arguments = @(
      "--export-type=png",
      "--export-filename=`"$pngPath`"",
      "--export-width=$($svgFile.Size)",
      "--export-height=$($svgFile.Size)",
      "`"$svgPath`""
    )
        
    try {
      $process = Start-Process -FilePath $InkscapePath -ArgumentList $arguments -Wait -NoNewWindow -PassThru
      if ($process.ExitCode -eq 0) {
        Write-Host "✓ $pngFile の変換が完了しました" -ForegroundColor Green
      }
      else {
        Write-Host "✗ $pngFile の変換に失敗しました (終了コード: $($process.ExitCode))" -ForegroundColor Red
      }
    }
    catch {
      Write-Host "✗ $pngFile の変換中にエラーが発生しました: $($_.Exception.Message)" -ForegroundColor Red
    }
  }
  else {
    Write-Host "⚠ SVGファイルが見つかりません: $svgPath" -ForegroundColor Yellow
  }
}

Write-Host "`nPNG変換が完了しました！" -ForegroundColor Green
Write-Host "生成されたPNGファイルを確認してください: $IconsPath" -ForegroundColor Cyan

# オプション：ファイルエクスプローラーでフォルダを開く
$openFolder = Read-Host "`nアイコンフォルダを開きますか？ (y/n)"
if ($openFolder -eq 'y' -or $openFolder -eq 'Y') {
  Start-Process explorer.exe -ArgumentList $IconsPath
}
