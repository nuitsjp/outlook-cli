# outlook-cli

クラシック Outlook でメールの取得、既読・未読の変更、送信ができる Windows 用 CLI です。実行には、設定済みの Outlook プロファイルと .NET 10 ランタイムが必要です。

## 使い方

`outlook-cli` または `outlook-cli mail` をサブコマンドなしで実行すると、その階層の `-h` と同じヘルプを表示します。

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

```powershell
outlook-cli mail send --to 'recipient@example.com' --subject '打ち合わせ' --body '明日の資料を共有します。'
```

`--to` に宛先1件、`--subject` に件名、`--body` にプレーンテキストの本文を指定します。クラシック Outlook セッションの既定のアカウントを使います。Outlook への送信依頼が成功すると、標準出力に `{"schemaVersion":1,"submitted":true}` を1行出力し、終了コード `0` を返します。エラーは標準エラー出力に表示し、終了コード `1` を返します。

## 開発

```powershell
dotnet build
dotnet run --project src/OutlookCli -- mail list --unread
dotnet test
```

`OutlookCli.Core` が `IMailQuery` とデータ型を定義し、`OutlookCli.OutlookCom` がクラシック Outlook に接続します。コマンドのテストでは `IMailQuery` をスタブに差し替えるため、Outlook のインストールは不要です。
