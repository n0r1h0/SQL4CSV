Set reg = CreateObject("VBScript.RegExp")
reg.Pattern = ".*system32.*"

If reg.Test(LCase(WScript.FullName)) Then
  ' Shell�I�u�W�F�N�g���쐬����
  Dim sh
  Set sh = WScript.CreateObject("WScript.Shell")
  Dim arg
  arg = ""
  For Each a In WScript.Arguments
    arg = arg & " """ & a & """"
  next
  sh.Run "cmd /C C:\Windows\SysWow64\cscript.exe //Nologo """ & WScript.ScriptFullName & """ "& _
  arg & _
  " & echo. & set /p=�����L�[�������ďI�����Ă�������<NUL & pause >NUL & echo.", 1, False

  '�I�u�W�F�N�g���J������
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
    WScript.Echo "�G���[!!!"
    WScript.echo "csv�t�@�C�����w�肵�Ă��������B"
    WScript.Quit -1
End If

inputPath = Trim(WScript.Arguments(0))

If Not fs.FileExists(inputPath) Then
  WScript.Echo ""
  WScript.Echo "�G���[!!!"
  WScript.echo "�t�@�C����������܂���B"
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
  WScript.Echo "�G���[!!!"
  WScript.echo "csv�t�@�C�����ǂݎ��܂���ł����B"
  WScript.echo "�w��̃t�@�C����������csv�t�H�[�}�b�g�ƂȂ��Ă��邩�m�F���Ă��������B"
  WScript.Quit -1
End If

'���R�[�h�̓ǂݍ��݁A�ҏW
rs.MoveFirst
Dim colnumnames

Dim Cr
Cr = Chr(13)
'�w�b�_���擾
For Each f In rs.Fields
  colnumnames = colnumnames & f.Name & Cr
Next

Dim where
where = InputBox("������������͂��ĉ������B(WHERE �ȍ~)" & Cr & "�y�񖼁z" & Cr & colnumnames)

If where = "" Then
  WScript.Echo ""
  WScript.echo "�����������w�肳��܂���ł����B"
  WScript.echo "�������I�����܂��B"
  WScript.Quit -1
End If

WScript.Echo inputFile & "���� ����'" & where & "' ���������B�B�B"

Set rs = cn.Execute("SELECT * " & _ 
"FROM [" & inputFile & "] " & _
"WHERE " & where & " " )

'���R�[�h�̓ǂݍ��݁A�ҏW
rs.MoveFirst

Dim outputBuffer
outputBuffer = ""

'�o�̓t�@�C���̓W�J
suffix = Replace(Replace(Replace(Now(), "/", ""), ":", ""), " ", "")
outputPath = inputDir & Left(inputFile, InStrRev(inputFile, ".") - 1) & "_" & suffix & ".csv"
Set fw = fs.OpenTextFile(outputPath, 2, True)

WScript.Echo "�w�b�_���������ݒ��B�B�B"
For Each f In rs.Fields
  outputBuffer = outputBuffer & f.Name & ","
Next
outputBuffer = Left(outputBuffer, Len(outputBuffer)-1)
fw.WriteLine outputBuffer

WScript.Echo "�f�[�^�ǂݍ��݊J�n�B�B�B"
Dim i
i = 1

Do Until rs.EOF
  outputBuffer = ""
  '���ڂ̏o��
  For Each f In rs.Fields
    outputBuffer = outputBuffer & f.Value & ","
  Next
  '�Ō��","���폜
  outputBuffer = Left(outputBuffer, Len(outputBuffer)-1)
  fw.WriteLine outputBuffer

  i = i + 1
  rs.MoveNext
Loop

rs.Close
Set rs = Nothing

WScript.Echo "����" 

Set fs = Nothing

cn.Close
Set cn = Nothing

WScript.Quit(0)