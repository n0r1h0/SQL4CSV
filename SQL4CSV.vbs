Set reg = CreateObject("VBScript.RegExp")
reg.Pattern = ".*system32.*"

If reg.Test(LCase(WScript.FullName)) Then
  ' Shellオブジェクトを作成する
  Dim sh
  Set sh = WScript.CreateObject("WScript.Shell")
  Dim arg
  arg = ""
  For Each a In WScript.Arguments
    arg = arg & " """ & a & """"
  next
  sh.Run "cmd /C C:\Windows\SysWow64\cscript.exe //Nologo """ & WScript.ScriptFullName & """ "& _
  arg & _
  " & echo. & set /p=何かキーを押して終了してください<NUL & pause >NUL & echo.", 1, False

  'オブジェクトを開放する
  Set sh = Nothing
  WScript.Quit(0)
End If

Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")

Dim	cn
Dim	rs
Dim outputPath
Dim inputDir
Dim inputFile
Dim inputPath

If WScript.Arguments.Count <> 1 Then
    WScript.Echo ""
    WScript.Echo "エラー!!!"
    WScript.echo "csvファイルを指定してください。"
    WScript.Quit -1
End If

inputPath = Trim(WScript.Arguments(0))

If Not fs.FileExists(inputPath) Then
  WScript.Echo ""
  WScript.Echo "エラー!!!"
  WScript.echo "ファイルが見つかりません。"
  WScript.Quit -1
End If

inputDir = Left(inputPath, InStrRev(inputPath, "\"))
inputFile = Mid(inputPath, Len(inputDir) + 1, Len(inputPath))

Set cn = CreateObject("ADODB.Connection")
cn.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};DBQ=" & inputDir & ";ReadOnly=1"

'On Error Resume Next
Set rs = cn.Execute("SELECT top 1 * " & _ 
"FROM [" & inputFile & "] ")
On Error Goto 0

If cn.Errors.Count > 0 Then
  WScript.Echo ""
  WScript.Echo "エラー!!!"
  WScript.echo "csvファイルが読み取れませんでした。"
  WScript.echo "指定のファイルが正しいcsvフォーマットとなっているか確認してください。"
  WScript.Quit -1
End If

'レコードの読み込み、編集
rs.MoveFirst
Dim colnumnames

Dim Cr
Cr = Chr(13)
'ヘッダを取得
For Each f In rs.Fields
  colnumnames = colnumnames & f.Name & Cr
Next

Dim where
where = InputBox("検索条件を入力して下さい。(WHERE 以降)" & Cr & "【列名】" & Cr & colnumnames)

If where = "" Then
  WScript.Echo ""
  WScript.echo "検索条件が指定されませんでした。"
  WScript.echo "処理を終了します。"
  WScript.Quit -1
End If

WScript.Echo inputFile & "から 条件'" & where & "' を検索中。。。"

Set rs = cn.Execute("SELECT * " & _ 
"FROM [" & inputFile & "] " & _
"WHERE " & where & " " )

'レコードの読み込み、編集
rs.MoveFirst

Dim outputBuffer
outputBuffer = ""

'出力ファイルの展開
suffix = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
outputPath = inputDir & Left(inputFile, InStrRev(inputFile, ".") - 1) & "_" & suffix & ".csv"
Set fw = fs.OpenTextFile(outputPath, 2, True)

WScript.Echo "ヘッダを書き込み中。。。"
For Each f In rs.Fields
  outputBuffer = outputBuffer & f.Name & ","
Next
outputBuffer = Left(outputBuffer, Len(outputBuffer)-1)
fw.WriteLine outputBuffer

WScript.Echo "データ読み込み開始。。。"
Dim i
i = 1

Do Until rs.EOF
  outputBuffer = ""
  '項目の出力
  For Each f In rs.Fields
    outputBuffer = outputBuffer & f.Value & ","
  Next
  '最後の","を削除
  outputBuffer = Left(outputBuffer, Len(outputBuffer)-1)
  fw.WriteLine outputBuffer

  i = i + 1
  rs.MoveNext
Loop

rs.Close
Set rs = Nothing

WScript.Echo "完了" 

Set fs = Nothing

cn.Close
Set cn = Nothing

WScript.Quit(0)