# screen_capture

Windows で静止画のスクリーンキャプチャを取るための PowerShell ツールです。

`Alt+Shift+S` でキャプチャ UI を開き、全画面キャプチャまたは範囲指定キャプチャを行います。保存は自動で行われ、PNG 形式で `captures` フォルダに出力されます。

## できること

- `Alt+Shift+S` のグローバルショートカットでキャプチャ UI を開く
- 全画面キャプチャを自動保存する
- ドラッグで範囲選択した領域を自動保存する
- 画像を日時付きファイル名の PNG として保存する
- 将来の動画機能追加を見越して、起動処理と保存処理を分離した構成にしている

## ファイル

### `screen_capture.ps1`

本体スクリプトです。

主なモード:

- `Listen`: 常駐して `Alt+Shift+S` を待ち受ける
- `Capture`: 単発でキャプチャする

主な引数:

- `-Mode Listen`
- `-Mode Capture`
- `-CaptureMode Prompt`
- `-CaptureMode FullScreen`
- `-CaptureMode Region`
- `-OutputDirectory`
- `-LogPath`

## 使い方

常駐起動:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\screen_capture.ps1
```

起動後は `Alt+Shift+S` を押すと選択ダイアログが開きます。

- `全画面`: 仮想スクリーン全体を保存します
- `範囲指定`: ドラッグした領域を保存します
- `キャンセル`: 何も保存しません

単発で全画面キャプチャする例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\screen_capture.ps1 -Mode Capture -CaptureMode FullScreen
```

単発で範囲指定キャプチャする例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\screen_capture.ps1 -Mode Capture -CaptureMode Region
```

## 保存先

既定ではスクリプトと同じディレクトリ配下の `captures` フォルダへ保存します。

- 画像保存先: `screen_capture\captures`
- ログ: `screen_capture\screen_capture.log`

必要なら `-OutputDirectory` と `-LogPath` で変更できます。

## 操作

- 起動ショートカット: `Alt+Shift+S`
- `F`: 全画面
- `R`: 範囲指定
- `Esc`: キャンセル
- 左ドラッグで選択
- `Esc` でキャンセル

## 制約

- 範囲指定はオーバーレイを閉じたあとに取得するため、選択 UI 自体は画像に含まれません
- 最小選択サイズ未満の小さなドラッグはキャンセル扱いになります
- 管理者権限アプリや一部の保護された画面では、期待どおりに取得できない場合があります
- 常駐はこのスクリプトを開いた PowerShell セッションが生きている間だけ有効です

## 将来の動画拡張

現状は静止画専用です。

後から動画機能を追加するときは、以下を流用できます。

- ホットキー監視
- 保存先管理
- キャプチャ対象の選択フロー
- 将来の `VideoFullScreen` / `VideoRegion` 追加に向けた引数構成
