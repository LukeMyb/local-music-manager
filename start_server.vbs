Set objFSO = CreateObject("Scripting.FileSystemObject")
' VBSファイル自身が置かれているフォルダのパスを自動取得
currentDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

Set WshShell = CreateObject("WScript.Shell")

' 取得したパス(currentDir)を使ってコマンドを動的に組み立てる
command = "cmd /c cd /d """ & currentDir & """ && .venv\Scripts\python.exe main.py"

' 0は「ウィンドウを非表示にする」、Falseは「実行の完了を待たない（裏で動かし続ける）」
WshShell.Run command, 0, False