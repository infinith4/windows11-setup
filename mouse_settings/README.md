# mouse_settings

このディレクトリには、Windows のマウス感度設定を自動化する PowerShell スクリプトがあります。

## できること

- Windows のマウス感度を `9` に設定するスクリプトを生成する
- ログオン時にその設定を自動適用するタスクをタスクスケジューラへ登録する
- 設定用スクリプトを非表示で実行し、毎回のサインイン後に同じ感度を再適用できるようにする

## スクリプトの内容

### `1_setup_mouse_sensitivity.ps1`

このスクリプトは、実際にマウス感度を変更するための PowerShell スクリプトを生成します。

- `user32.dll` の `SystemParametersInfo` を呼び出すコードを書き出す
- `C:\Apps\windows_mouse_settings` ディレクトリを作成する
- `C:\Apps\windows_mouse_settings\mouse_sensitivity.ps1` を生成する
- 生成されたスクリプト内でマウス感度を `9` に設定する

### `2_setup_mouse_taskscheduler.ps1`

このスクリプトは、上で生成した `mouse_sensitivity.ps1` をログオン時に自動実行する設定を追加します。

- `powershell.exe` で `C:\Apps\windows_mouse_settings\mouse_sensitivity.ps1` を実行するタスクを作成する
- 実行ウィンドウを非表示にする
- `ExecutionPolicy Bypass` で実行する
- ログオン時に起動するトリガーを設定する
- `MouseSensitivity` という名前でタスクスケジューラに登録する
- 最高権限で実行する設定にする

## 実行順

1. `1_setup_mouse_sensitivity.ps1`
2. `2_setup_mouse_taskscheduler.ps1`

先に 1 を実行して設定用スクリプトを作成し、その後に 2 を実行して自動実行を登録します。
