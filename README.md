# outlook-cli

クラシック Outlook のメール、予定表、タスク、連絡先を操作する Windows 用 CLI です。実行には、設定済みの Outlook プロファイルと .NET Framework 4.8.1 が必要です。

## インストールと更新

Windows x64 では、[GitHub Releases](https://github.com/nuitsjp/outlook-cli/releases) から `outlook-cli-<バージョン>-setup.exe` を取得して実行します。署名なしの Inno Setup インストーラーで、管理者権限を使わずにユーザー単位でインストールします。.NET Framework 4.8.1 は事前にインストールしてください。

既定のインストール先で PowerShell から動作を確認できます。

```powershell
& "$env:LOCALAPPDATA\Programs\outlook-cli\outlook-cli.exe" --version
```

インストーラーは `%LOCALAPPDATA%\Microsoft\WinGet\Links` に起動用の shim を作成します。この共通フォルダーがすでに `Path` に登録されていれば、`Path` の変更なしで `outlook-cli` と名前だけで起動できます。アンインストール時に shim も削除されます。更新時は GitHub Releases から新しいインストーラーを取得し、実行します。

インストーラーで導入した CLI は、`outlook-cli update` で公開リポジトリの GitHub Releases から最新版に更新できます。GitHub へのログインやトークンの設定は不要です。

```powershell
outlook-cli update
```

新版がある場合は、現在のバージョンと新版を標準エラー出力に表示して `[Y/n]` で確認します。`y` または Enter で更新を進め、`n` で見送ります。

`status` が `up-to-date` の場合は現在のバージョンを使い続けます。`declined` の場合は更新を見送り、インストーラーのダウンロードは行いません。`scheduled` の場合は、ダウンロードしたインストーラーの SHA-256 を照合済みで、CLI と shim の終了後に同じ場所にサイレントインストールします。インストーラーの終了結果は `logPath` のログで確認できます。通常のコマンド起動時には更新確認を行いません。

## 使い方

ルートと各コマンドグループをサブコマンドなしで実行すると、その階層の `-h` と同じヘルプを表示します。

```powershell
outlook-cli mail list --unread
```

`--folder` を省略すると、既定の受信トレイ直下が対象になります。`--unread` を指定すると未読のメールに絞ります。受信日時の新しい順に取得し、既読状態は維持します。

Outlook が起動済みなら現在のセッションを使います。未起動なら Windows の COM 登録で起動し、取得後に COM 参照を解放して Outlook の終了を待ちます。20秒以内に終了しなければ診断を出して終了コード `1` を返します。

標準出力の JSON と JSONL は BOM なしの UTF-8 です。既定では日本語や絵文字をそのまま出力し、文字列中の引用符や改行は JSON の構文に従ってエスケープします。CP932 などで外部コマンドの出力を読み取る環境では、共通オプション `--ascii` を指定すると、非 ASCII 文字を従来どおり Unicode エスケープで出力します。両方の形式で `ConvertFrom-Json` は同じ値を返します。

```powershell
outlook-cli mail list --unread --ascii
```

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

### メール送信

```powershell
outlook-cli mail send --to 'recipient@example.com' --subject '打ち合わせ' --body '明日の資料を共有します。'
```

`--to`、`--subject`、`--body` に宛先、件名、プレーンテキストの本文を指定します。クラシック Outlook の既定のアカウントから送信し、成功時は標準出力に `{"schemaVersion":1,"submitted":true}` を1行出力して終了コード `0` を返します。`--ascii` を指定した場合は既存の JSON 出力規則に従います。

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

一覧は1件1行の JSONL です。繰り返し予定は開催回ごとに表示されます。各行の `id` をそのまま `calendar show` に渡すと、その回の主催者・参加者・プレーンテキスト本文を含む JSON を取得できます。メールコマンドと同じく、`--ascii` で ASCII 出力に切り替えられます。

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

`OutlookCli.Core` が `IMailQuery`、`ICalendarQuery` などのインターフェースとデータ型を定義し、`OutlookCli.OutlookCom` がクラシック Outlook に接続します。コマンドのテストではインターフェースをスタブに差し替えるため、Outlook のインストールは不要です。

## リリース

配布するコミットをチェックアウトし、次のいずれかを実行します。

```powershell
mise run release          # 最新の vX.Y.Z タグからパッチ番号を1増やす
mise run release -- 1.2.0 # バージョンを指定する
```

安定版のタグがない場合、バージョンを省略した初回リリースは `v0.1.0` です。バージョンの各要素は、Windows のファイルバージョンに合わせて `0` から `65534` までです。タスクは現在のコミットに `vX.Y.Z` タグを作成し、そのタグを `origin` にプッシュします。タグのプッシュで GitHub Actions が起動し、テストを実行したうえでインストーラーをビルドし、GitHub Releases に発行します。
