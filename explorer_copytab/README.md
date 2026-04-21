# explorer_copytab

Windows 11 の `Explorer` で、アクティブなタブと同じ場所を新しいタブで開くための PowerShell スクリプトです。

`RegisterHotKey` と `SendInput` を `user32.dll` 経由で呼び出します。新しいタブへの移動自体は `Shell.Application` の `Navigate2()` を使います。

## できること

- グローバルショートカット `Ctrl+Alt+Shift+D` で現在の Explorer タブを複製する
- `RunOnce` モードで 1 回だけ複製して終了する
- ログオン時にタスク スケジューラから自動起動する

## ファイル

### `explorer_copytab.ps1`

常駐する本体スクリプトです。

処理の流れ:

1. 前面ウィンドウが `Explorer` か確認する
2. アクティブタブのアドレスバーを `Ctrl+L` / `Ctrl+A` / `Ctrl+C` で読み取り、複製元パスを取得する
3. `Ctrl+T` で新しいタブを開く
4. 同じ `HWND` に属する新しい Explorer タブを見つける
5. そのタブに対して `Navigate2()` で複製元パスへ移動する

主な引数:

- `-HotKey`
- `-NewTabDelayMs`
- `-AddressBarDelayMs`
- `-PostNewTabSettleDelayMs`
- `-PostEnterVerifyDelayMs`
- `-LogPath`
- `-RunOnce`
- `-RunOnceWaitMs`

起動例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1
```

ログを明示する例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -LogPath "$env:TEMP\explorer_copytab.log"
```

前面の Explorer に対して 1 回だけ実行して終了する例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -RunOnce
```

ホットキーを変える例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -HotKey Ctrl+Shift+Y
```

### `1_setup_explorer_copytab.ps1`

本体スクリプトを `C:\Apps\windows_explorer_copytab\explorer_copytab.ps1` に配置します。

配置時に以下を確認します。

- 配布元ファイルの存在
- コピー先ファイルの生成
- コピー元とコピー先の `SHA256` 一致

### `2_setup_explorer_copytab_taskscheduler.ps1`

ログオン時に `explorer_copytab.ps1` を非表示で起動する `ExplorerCopyTab` タスクを登録します。

登録内容:

- 実行ファイル: `powershell.exe`
- 引数: `-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Apps\windows_explorer_copytab\explorer_copytab.ps1" -LogPath "%TEMP%\explorer_copytab.log"`
- トリガー: ログオン時
- 多重起動: `IgnoreNew`

## セットアップ

1. `1_setup_explorer_copytab.ps1` を実行する
2. `2_setup_explorer_copytab_taskscheduler.ps1` を実行する
3. サインアウト/サインインするか、必要なら手動で `explorer_copytab.ps1` を起動する

## 制約

- 公開 API に「Explorer タブの複製」はないため、実装は「アクティブタブの場所を取得して新しいタブで同じ場所を開く」方式です
- `SendInput` を使うため、Explorer の UI 変更や高負荷時には待機時間調整が必要になる場合があります
- ホットキーは離鍵待ちをしてから注入するため、押しっぱなしだと反応が不安定になります
- 管理者権限で動いている Explorer に対しては、通常権限のスクリプトから入力注入できない場合があります
- 特殊フォルダやアドレスバーから通常のファイル システム パスが取得できない場所では、期待どおり複製できない場合があります
- `Shell.Application` の列挙順はアクティブタブ順とは一致しないため、デバッグログの `Matched Explorer` は必ずしも複製元タブそのものを示しません

## デバッグ

既定ログは `%TEMP%\explorer_copytab.log` です。

主な確認ポイント:

- `Listener started`: 常駐起動に成功
- `HotKey received`: ホットキー受信に成功
- `Captured active tab target from UI: ...`: 複製元パスの取得に成功
- `Explorer windows before Ctrl+T count=...` / `after Ctrl+T count=...`: 新しいタブ候補の列挙結果
- `Navigate2 invoked for path=...`: 新しいタブへの移動要求を実行
- `Failed to capture active tab target from UI...`: アクティブタブの取得に失敗し、COM のパスへフォールバック

## 動作確認

1. Explorer を開く
2. 複製したいタブを前面にする
3. `explorer_copytab.ps1` を起動する
4. `Ctrl+Alt+Shift+D` を押して離す
5. 同じ場所の新しいタブが開くことを確認する
