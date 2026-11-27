# Outlook CLI の開発方針

## コマンドと出力

- 操作は `mail list`、`mail show` のようにコマンドとサブコマンドで表し、絞り込みや対象指定には `--unread`、`--entry-id` などのオプションを使う。
- 標準出力には機械処理用の JSON を出す。一覧と監視イベントは1件1行の JSONL、明細と更新結果は1個の JSON オブジェクトにし、各オブジェクトに `schemaVersion` を含める。実行時の案内とエラーは標準エラーに出す。途中で失敗した場合は出力済みの行を保持し、非ゼロの終了コードを返す。
- JSON／JSONL は共通の出力処理を使う。既定は BOM なしの UTF-8 で、日本語や絵文字をそのまま出す。引用符、バックスラッシュ、改行など、JSON の構文に必要なエスケープは維持する。共通オプション `--ascii` を指定した場合は、非 ASCII 文字を Unicode エスケープにした従来の ASCII 出力に切り替える。どちらも同じ項目と値を返す。

## Outlook と検証

- 実行対象は Windows の .NET Framework 4.8.1 とクラシック Outlook。ターゲットは `Directory.Build.props`、ビルド用 SDK は `global.json` で指定する。
- COM 呼び出しは `OutlookCli.OutlookCom` に置く。CLI は `OutlookCli.Core` のインターフェースを介して呼び出し、コマンドのテストではスタブに差し替える。
- Outlook には COM 登録を使って接続する。取得した COM 参照を解放し、CLI が起動した Outlook の終了を最大20秒待つ。時間内に終了しなければエラーを返す。すでに起動していた Outlook は閉じない。COM 接続や後始末を変更したときは実機でも確認する。
- メールの一覧・明細・監視は既読状態を変更しない。監視は Outlook の受信イベントを使い、Ctrl+C で購読を解除する。既読・未読の更新は対象の Entry ID に限定し、指定状態に達していれば `changed: false` を返す。

## ビルドと配布

- ビルド、テスト、開発中の実行には `mise run build`、`mise run test`、`mise run run` を使う。リリースは `mise run release` で `vX.Y.Z` タグをプッシュし、GitHub Actions から GitHub Releases にインストーラーを発行する。
- 配布には署名なしのユーザー単位の Inno Setup インストーラーを使う。インストーラーは `Path` を変更せず、`%LOCALAPPDATA%\Microsoft\WinGet\Links` に shim を置き、アンインストール時に削除する。
- 自動更新は明示的な `update` で実行する。公開リポジトリの GitHub Releases を認証なしで参照し、正式版とインストール済みのバージョンを比較する。新版のインストーラーは GitHub が返す SHA-256 を照合してから、CLI と shim の終了後にサイレント実行する。CLI の通常起動時に更新確認や同意プロンプトを出さない。
