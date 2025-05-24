---
applyTo: "**/*.ps1,**/*.psm1"
---
# Microsoft 365 (PowerShell) コーディング基準

[一般コーディングガイドライン](./.github/instructions/general-coding.instructions.md)をすべてのコードに適用してください。

## PowerShell ガイドライン
- Microsoft.Graph モジュールや ExchangeOnlineManagement モジュールを適切に活用する
- コマンドレット名は標準の動詞-名詞の形式に従う
- パラメータは適切に型付けし、必須パラメータには [Parameter(Mandatory=$true)] を使用する
- エラー処理には try/catch ブロックと $ErrorActionPreference を適切に組み合わせて使用する

## スクリプト構造
- ヘルプコメントを含める（.SYNOPSIS, .DESCRIPTION, .PARAMETER, .EXAMPLE など）
- スクリプトの先頭で必要なモジュールを Import-Module で読み込む
- 大きなスクリプトは機能ごとに関数に分割する
- 再利用可能なコードはモジュール (.psm1) として実装する

## セキュリティ対策
- 認証情報は安全に管理し、平文でスクリプトに含めない
- 可能な限り最小権限の原則に従う
- API 呼び出しの前に適切な権限があることを確認する
- スクリプト実行のログを記録し、監査可能にする
