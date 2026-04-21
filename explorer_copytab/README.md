# explorer_copytab

Windows 11 の `Explorer` で、アクティブなタブと同じ場所を新しいタブで開くための PowerShell スクリプトです。

`RegisterHotKey` と `SendInput` を `user32.dll` 経由で呼び出します。新しいタブへの移動自体は `Shell.Application` の `Navigate2()` を使います。
アクティブタブの特定は `ShellTabWindowClass` 配下の UI Automation で選択中タブ名を取得し、`Shell.Application` の列挙結果と突き合わせます。新しいタブの生成は Explorer の `ShellTabWindowClass` に対する `WM_COMMAND 0xA21B` で行います。
選択中タブ名が取れない場合は、UI Automation のフォーカス要素名と各タブの `Document.FocusedItem` を突き合わせるフォールバックを使います。

## できること

- グローバルショートカット `Alt+Shift+Z` で現在の Explorer タブを複製する
- `RunOnce` モードで 1 回だけ複製して終了する
- ログオン時にタスク スケジューラから自動起動する

## ファイル

### `explorer_copytab.ps1`

常駐する本体スクリプトです。

処理の流れ:

1. 前面ウィンドウが `Explorer` か確認する
2. `ShellTabWindowClass` 配下の UI Automation から選択中タブ名を取得し、対応するタブのパスを特定する
3. `ShellTabWindowClass` に `WM_COMMAND 0xA21B` を送って新しいタブを開く
4. 同じ `HWND` に属する新しい Explorer タブを見つける
5. そのタブに対して `Navigate2()` で複製元パスへ移動する

主な引数:

- `-HotKey`
- `-PostHotKeyCooldownMs`
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
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -HotKey Alt+Shift+Z
```

ホットキー直後の合成入力待機を長めにする例:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\explorer_copytab.ps1 -PostHotKeyCooldownMs 900
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
- 引数: `-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "C:\Apps\windows_explorer_copytab\explorer_copytab.ps1" -PostHotKeyCooldownMs 900  -HotKey Alt+Shift+D -LogPath "%TEMP%\explorer_copytab.log"`
- トリガー: ログオン時
- 多重起動: `IgnoreNew`

## セットアップ

1. `1_setup_explorer_copytab.ps1` を実行する
2. `2_setup_explorer_copytab_taskscheduler.ps1` を実行する
3. サインアウト/サインインするか、必要なら手動で `explorer_copytab.ps1` を起動する

## 制約

- 公開 API に「Explorer タブの複製」はないため、実装は「アクティブタブの場所を取得して新しいタブで同じ場所を開く」方式です
- 新しいタブ生成には Explorer の undocumented な `WM_COMMAND 0xA21B` を使うため、将来の Explorer 更新で変わる可能性があります
- ホットキーは離鍵待ち後に `PostHotKeyCooldownMs` だけ待ってから注入するため、別アプリの二重 `Ctrl` 判定に巻き込まれにくくしています
- 離鍵待ちがタイムアウトした場合だけ保険として修飾キーの解放を送るため、押しっぱなしだと反応が不安定になります
- 管理者権限で動いている Explorer に対しては、通常権限のスクリプトから入力注入できない場合があります
- 特殊フォルダやアドレスバーから通常のファイル システム パスが取得できない場所では、期待どおり複製できない場合があります
- `Shell.Application` の列挙順はアクティブタブ順とは一致しないため、新規タブ候補の特定は列挙差分と空パスタブの有無に依存します

## デバッグ

既定ログは `%TEMP%\explorer_copytab.log` です。

主な確認ポイント:

- `Listener started`: 常駐起動に成功
- `HotKey received`: ホットキー受信に成功
- `Matched Explorer ... shellTabHwnd=... shellTabClass=ShellTabWindowClass`: アクティブタブと新規タブ送信先の取得に成功
- `Selected tab lookup succeeded. ... title=...`: UI Automation で選択中タブ名の取得に成功
- `Focused element lookup succeeded. ...`: UI Automation で現在フォーカス中の要素取得に成功
- `Explorer windows before new-tab command count=...` / `after new-tab command count=...`: 新しいタブ候補の列挙結果
- `Navigate2 invoked for path=...`: 新しいタブへの移動要求を実行
- `No Explorer navigation target found after new-tab command.`: 新しいタブ候補を特定できなかった

## 動作確認

1. Explorer を開く
2. 複製したいタブを前面にする
3. `explorer_copytab.ps1` を起動する
4. `Ctrl+Alt+Shift+D` を押して離す
5. 同じ場所の新しいタブが開くことを確認する
