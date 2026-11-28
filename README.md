# outlook-cli

クラシック Outlook のメール、予定表、タスク、連絡先を操作する Windows 用 CLI です。実行には、設定済みの Outlook プロファイルと .NET 10 ランタイムが必要です。

## 使い方

ルートと各コマンドグループをサブコマンドなしで実行すると、その階層の `-h` と同じヘルプを表示します。

```powershell
outlook-cli mail list --unread
```

`--folder` を省略すると、既定の受信トレイ直下が対象になります。`--unread` を指定すると未読のメールに絞ります。受信日時の新しい順に取得し、既読状態は維持します。

Outlook が起動済みなら現在のセッションを使います。未起動なら Windows の COM 登録で起動し、取得後に COM 参照を解放して Outlook の終了を待ちます。20秒以内に終了しなければ標準エラー出力に診断を出します。通常の終了コードは `1` です。`mail send` で送信依頼が受付済みの場合は `0` を返します。COM 参照の解放に失敗した場合も、残りの参照の解放を試みます。操作と後処理の両方に失敗した場合は、標準エラー出力に操作の診断と後処理の診断を順に表示します。

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

`--to` に宛先1件、`--subject` に件名、`--body` にプレーンテキストの本文を指定します。クラシック Outlook セッションの既定のアカウントを使います。Outlook への送信依頼が成功すると、標準出力に `{"schemaVersion":1,"submitted":true}` を1行出力し、終了コード `0` を返します。送信依頼が失敗した場合は標準エラー出力に診断を表示し、終了コード `1` を返します。Outlook が送信依頼を受け付けた後で COM 参照の解放に失敗した場合は、`submitted: true` と終了コード `0` を返し、後処理の診断を標準エラー出力に表示します。

### フォルダー、添付ファイル、返信

`mail folder list` は既定の受信トレイと同じストアのメールフォルダーを JSONL で列挙します。各行の `folderId` を `mail list --folder` や `mail move --to-folder` に渡せます。`--subject` は件名の部分一致、`--since` は指定日時以降の受信メールで絞り込みます。日付だけ指定すると Outlook のローカル時刻の午前0時が基準になります。`--unread` と組み合わせて使えます。

```powershell
outlook-cli mail folder list
outlook-cli mail list --folder '<一覧の folderId>' --subject '請求書' --since '2026-11-27T00:00:00+09:00'
outlook-cli mail move --entry-id '<メールの entryId>' --to-folder '<移動先の folderId>'
```

移動時は移動後の `entryId`、`folderId`、`storeId`、`changed` を返します。

添付ファイルの一覧から `attachmentId` を選び、ファイル名まで含めた保存先を指定します。保存先に同名ファイルがある場合はエラーになります。

```powershell
outlook-cli mail attachment list --entry-id '<メールの entryId>'
outlook-cli mail attachment save --entry-id '<メールの entryId>' --attachment-id 1 --path '.\請求書.pdf'
```

`mail reply` は元メールへの返信を下書きに保存し、`draftEntryId` を返します。Outlook の下書きフォルダーで確認して送信できます。

```powershell
outlook-cli mail reply --entry-id '<メールの entryId>' --body '資料を受領しました。'
```

### 予定表

既定の予定表から、指定期間に重なる予定を開始時刻の昇順で取得します。期間の両端は `YYYY-MM-DD` で指定し、Outlook のローカルタイムゾーンで `--from` の午前0時から `--to` の午前0時の直前までを対象にします。下の例は11月27日から30日までです。

```powershell
outlook-cli calendar list --from 2026-11-27 --to 2026-12-01
outlook-cli calendar show --id '<一覧の id>'
```

一覧は1件1行の JSONL です。繰り返し予定は開催回ごとに表示されます。各行の `id` をそのまま `calendar show` に渡すと、その回の主催者・参加者・プレーンテキスト本文を含む JSON を取得できます。

閲覧権限のある共有予定表を参照するには、所有者を `--owner` で指定します。共有予定表から取得した `id` を表示するときも、同じ所有者を指定します。

```powershell
outlook-cli calendar list --owner 'owner@example.com' --from 2026-12-01 --to 2026-12-02
outlook-cli calendar show --owner 'owner@example.com' --id '<一覧の id>'
```

`calendar create` は、既定の予定表に予定ありの単発予定を登録します。日時は分までと UTC オフセットを指定します。返された `id` は `calendar show` で使えます。

```powershell
outlook-cli calendar create --subject '作業枠' --start '2026-12-01T13:00+09:00' --end '2026-12-01T14:00+09:00'
```

`calendar free-busy` は参加者ごとに空き時間情報を JSONL で返します。既定の予定表を持つアカウントの SMTP アドレスを指定した場合、その予定表にある予定から計算します。ほかの参加者は Outlook の空き時間照会で取得します。

1日から1か月までの日付範囲を指定し、`--to` の日は含みません。`slots` の先頭は開始日の午前0時で、1文字が `minutesPerSlot`（30分）に相当します。`0` は空き、`1` は仮の予定、`2` は予定あり、`3` は外出中を表します。本人の予定表から算出する場合、`5` は他の場所で作業中を表します。

```powershell
outlook-cli calendar free-busy --attendee 'a@example.com' --attendee 'b@example.com' --from 2026-12-01 --to 2026-12-02
```

### タスク

既定のタスクフォルダーの未完了タスクを JSONL で取得します。新しいタスクには任意で期限を指定できます。完了時は `changed` が `true` になり、すでに完了している場合は `false` になります。

```powershell
outlook-cli task list
outlook-cli task add --subject '資料を確認する' --due 2026-12-01
outlook-cli task complete --entry-id '<タスクの entryId>'
```

### 連絡先

既定の連絡先フォルダーから連絡先を取得します。`list` は氏名や会社名の一部で絞り込めます。`show` はメールアドレスと勤務先電話番号を返します。`add` は氏名とメールアドレスを指定して連絡先を追加します。

```powershell
outlook-cli contact list --name '山田' --company '会社名'
outlook-cli contact show --entry-id '<一覧の entryId>'
outlook-cli contact add --name '山田 太郎' --email 'taro@example.com'
```

`contact lookup` は、接続中の Exchange のアドレス帳で名前を解決し、SMTP アドレスを返します。

```powershell
outlook-cli contact lookup --name '山田 太郎'
```

## 開発

```powershell
dotnet build
dotnet run --project src/OutlookCli -- mail list --unread
dotnet test
```

`OutlookCli.Core` はメール・予定表・連絡先・タスクのインターフェースとデータ型を定義します。`OutlookCli` の起動時に、クラシック Outlook を操作する `OutlookCli.OutlookCom` の実装を渡します。コマンドのテストではすべての操作をテスト用の実装に差し替えます。COM の後処理、添付ファイルの保存、予定の ID と空き時間の計算も、Outlook を起動せずに検証できます。
