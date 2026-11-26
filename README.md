# outlook-cli

クラシック Outlook の既定の受信トレイにあるメールを取得し、既読・未読を変更できる Windows 用 CLI です。実行には、設定済みの Outlook プロファイルと .NET Framework 4.8.1 が必要です。

## インストールと更新

Windows x64 では、[GitHub Releases](https://github.com/R0007-RJ-SolDev/outlook-cli/releases) から `outlook-cli-<バージョン>-setup.exe` を取得して実行します。署名なしの Inno Setup インストーラーで、管理者権限を使わずにユーザー単位でインストールします。.NET Framework 4.8.1 は事前にインストールしてください。

既定のインストール先で PowerShell から動作を確認できます。

```powershell
& "$env:LOCALAPPDATA\Programs\outlook-cli\outlook-cli.exe" --version
```

インストーラーは `%LOCALAPPDATA%\Microsoft\WinGet\Links` に起動用の shim を作成します。この共通フォルダーがすでに `Path` に登録されていれば、`Path` の変更なしで `outlook-cli` と名前だけで起動できます。アンインストール時に shim も削除されます。更新時は GitHub Releases から新しいインストーラーを取得し、実行します。

## 使い方

```powershell
outlook-cli mail list --unread
```

`--unread` を省略すると、受信トレイ直下のすべてのメールが対象になります。受信日時の新しい順に取得し、既読状態は維持します。

Outlook が起動済みなら現在のセッションを使います。未起動なら Windows の COM 登録で起動し、取得後に COM 参照を解放して Outlook の終了を待ちます。20秒以内に終了しなければ診断を出して終了コード `1` を返します。

標準出力の JSON と JSONL は BOM なしの UTF-8 で、非 ASCII 文字を Unicode エスケープで表現します。PowerShell では `ConvertFrom-Json` で元の文字列を取り出せます。

一覧は JSONL で、各行に `schemaVersion`、`entryId`、`receivedAt`、`senderName`、`subject` が含まれます。0件の場合は出力0行で終了コード `0` です。失敗時は標準エラー出力に診断を出し、終了コード `1` を返します。途中で失敗すると、標準出力には取得済みの行が残ります。

```powershell
outlook-cli mail list --unread |
    ForEach-Object { $_ | ConvertFrom-Json } |
    Format-List entryId, receivedAt, senderName, subject
```

一覧の `entryId` を指定すると、メール1件の明細を JSON で取得できます。明細には一覧の項目に加え、Outlook の宛先表示名リスト `to`、`cc`、`bcc` と、プレーンテキストの `body` が含まれます。取得時に既読状態は変わりません。

```powershell
outlook-cli mail show --entry-id '<一覧の entryId>'
```

一覧の `entryId` を指定してメール1件の既読・未読を変更できます。

```powershell
outlook-cli mail mark-read --entry-id '<一覧の entryId>'
outlook-cli mail mark-unread --entry-id '<一覧の entryId>'
```

成功時は `schemaVersion`、`entryId`、`unread`、`changed` を含む JSON を1件返します。`unread` は `mark-read` で `false`、`mark-unread` で `true` です。状態が変わった場合は `changed: true`、すでに指定の状態だった場合は `changed: false` を返します。

```powershell
outlook-cli mail watch
```

クラシック Outlook の受信イベントを待機し、新着メールごとに `mail list` と同じ形式の JSONL を標準出力に1行書きます。購読の開始・終了は標準エラー出力に表示されます。待機中は Outlook と接続を保持し、Ctrl+C で購読を解除して終了します。メールの既読状態は維持します。

## 開発

リポジトリのルートで mise を使ってビルド、テスト、実行できます。ビルドには `global.json` に指定した .NET 10 SDK を使用し、実行ファイルは .NET Framework 4.8.1 を対象にします。

```powershell
mise run build
mise run test
mise run run -- mail list --unread
```

`run` は Release ビルド済みの `outlook-cli.exe` に `--` 以降の引数を渡します。CLI を変更したら `mise run build` を再実行してください。`run` の標準出力は、そのまま PowerShell のパイプに渡せます。

```powershell
mise run run -- mail list --unread |
    ForEach-Object { $_ | ConvertFrom-Json } |
    Format-List entryId, receivedAt, senderName, subject
```

`OutlookCli.Core` が `IMailQuery` とデータ型を定義し、`OutlookCli.OutlookCom` がクラシック Outlook に接続します。コマンドのテストでは `IMailQuery` をスタブに差し替えるため、Outlook のインストールは不要です。

## リリース

配布するコミットをチェックアウトし、次のいずれかを実行します。

```powershell
mise run release          # 最新の vX.Y.Z タグからパッチ番号を1増やす
mise run release -- 1.2.0 # バージョンを指定する
```

安定版のタグがない場合、バージョンを省略した初回リリースは `v0.1.0` です。バージョンの各要素は、Windows のファイルバージョンに合わせて `0` から `65534` までです。タスクは現在のコミットに `vX.Y.Z` タグを作成し、そのタグを `origin` にプッシュします。タグのプッシュで GitHub Actions が起動し、テストを実行したうえでインストーラーをビルドし、GitHub Releases に発行します。
