# explorer_copytab

Windows 11 の `Explorer` で、現在のタブと同じフォルダを新しいタブで開くための PowerShell スクリプトです。

AutoHotkey は使わず、`RegisterHotKey` と `SendInput` を `user32.dll` 経由で呼び出します。

## できること

- グローバルショートカットで `Explorer` の現在タブを複製する
- ログオン時に自動起動する

## ファイル

### `explorer_copytab.ps1`

常駐スクリプトです。

- 既定ショートカットは `Ctrl+Alt+Shift+D`
- 前面の `Explorer` ウィンドウが対象
- `Shell.Application` から現在のフォルダパスを取得する
- まず `Navigate2(..., navOpenInNewTab)` を試し、だめなら `Ctrl+T` と貼り付けによる入力注入へフォールバックする

例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1
```

前面の `Explorer` に対して 1 回だけ実行して終了する場合:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -RunOnce
```

`-RunOnce` は既定で 5 秒間、前面の `Explorer` を待ちます。必要なら待機時間を変えられます。

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -RunOnce -RunOnceWaitMs 10000
```

ホットキーを変える場合:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -HotKey Ctrl+Shift+Y
```

デバッグログを明示したい場合:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -LogPath "$env:TEMP\explorer_copytab.log"
```

## セットアップ

1. `1_setup_explorer_copytab.ps1`
2. `2_setup_explorer_copytab_taskscheduler.ps1`

`1_setup_explorer_copytab.ps1` は、実行用スクリプトを `C:\Apps\windows_explorer_copytab\explorer_copytab.ps1` にコピーします。

`2_setup_explorer_copytab_taskscheduler.ps1` は、ログオン時に非表示で起動する `ExplorerCopyTab` タスクを登録します。

## 制約

- 公開 API に「Explorer のタブを直接複製する」機能はないため、実装は「現在パスを取得して新しいタブで同じ場所を開く」方式です
- `SendInput` ベースなので、Explorer の UI 変更や重いマシンでは待機時間の調整が必要な場合があります
- 入力注入前にホットキーの離鍵待ちをしているため、ショートカットは押しっぱなしではなく一度しっかり離す必要があります
- 管理者権限で動作している Explorer に対しては、通常権限のスクリプトから入力注入できないことがあります
- 特殊フォルダでは `Document.Folder.Self.Path` が期待どおり取れない場合があります

## デバッグ

既定では `%TEMP%\explorer_copytab.log` にログを書きます。

- `Listener started` が無ければ起動失敗
- `HotKey received` が無ければショートカット未受信
- `Foreground ... class=...` が `CabinetWClass` でなければ前面が Explorer ではない
- `Matched Explorer` が無ければ COM から対象ウィンドウを見つけられていない
- `COM Navigate2 accepted ...` があれば入力注入を使わず COM 経由で新規タブを開いている

## 動作確認

1. `Explorer` を開く
2. 任意のフォルダへ移動する
3. `explorer_copytab.ps1` を実行する
4. `Ctrl+Alt+Shift+D` を押す
5. 同じパスの新しいタブが開くことを確認する
