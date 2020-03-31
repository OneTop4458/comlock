Attribute VB_Name = "Moudle_WriteLog"
'---------------------------------------------------------------------------------------
' Module    : Moudle_WriteLog
' Author    : http://www.lakshmikanth.com/write-log-using-vb6/
' Date      : 2019-05-03
' Purpose   : 로그 모듈
'---------------------------------------------------------------------------------------

Option Explicit '변수 미선언 방지 선언

Public Sub WriteLog(strMessage As String)

Dim hFile As Long
Dim sFolder As String
Dim sFile As String
Dim sLog As String

On Error GoTo errHandler

hFile = FreeFile()
sFolder = App.Path & "\Logs"
sFile = sFolder & "\ComLock_Setting.Log"

sLog = Format(Now(), "yyyy-mm-dd HH:nn:ss") & "." & Right(Format(Timer(), "0.000"), 3)
sLog = sLog & " : " & strMessage

Open sFile For Append As #hFile
    Print #hFile, sLog
Close hFile

Exit Sub

errHandler:

If Err.Number = 76 Then 'If folder not exists, create it
    MkDir sFolder
    Resume
End If
MsgBox Err.Number & ":" & Err.Description

End Sub

