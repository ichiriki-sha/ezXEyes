Option Explicit

Private fso
Private wsh

'-------------------------------------------------------------------------------
' 関数名 : GetFileName
' 概要   : ファイルパスよりファイル名を取得する。
' 引数   : path - ファイルパス
' 戻り値 : ファイル名を返却する。
'-------------------------------------------------------------------------------
Private Function GetFileName(path)
	Dim pos
	Dim ret

	ret = path

	pos = InStrRev(path, "\")
	
	If pos > 0 Then

		ret = Mid(path, pos + 1)
	End If

	GetFileName = ret
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetBaseFileName
' 概要   : ファイルパスよりファイルベース名を取得する。
' 引数   : path - ファイルパス
' 戻り値 : ファイルベース名を返却する。
'-------------------------------------------------------------------------------
Private Function GetBaseFileName(path)
	Dim pos
	Dim fileName
	Dim ret

	ret = GetFileName(path)

	pos = InStrRev(ret, ".")
	
	If pos > 0 Then

		ret = Left(ret, pos - 1)
	End If

	GetBaseFileName = ret
End Function

'-------------------------------------------------------------------------------
' 関数名 : Main
' 概要   : メイン
' 引数   : なし
' 戻り値 : 正常終了の場合、0　異常の場合、1を返却する。
'-------------------------------------------------------------------------------
Private Function Main()
	On Error Resume Next

	Dim baseFile , scriptPath , curDir , cmd , ret
	Const cmd_sklton = "powershell -ExecutionPolicy Bypass -File ""${File}"""
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set wsh = CreateObject("WScript.Shell")

	ret = 0
	
	curDir = fso.GetParentFolderName(WScript.ScriptFullName)
	scriptPath = curDir & "\" & GetBaseFileName(WScript.ScriptFullName) & ".ps1"

	cmd = Replace(cmd_sklton, "${File}", scriptPath)
	wsh.Run cmd, 0

	If Err.Number <> 0 Then

		ret = 1
	End If

	Set fso = Nothing
	Set wsh = Nothing

	Main = ret
End Function

WScript.Quit(Main())