VERSION 5.00
Begin VB.Form Frm_First 
   BorderStyle     =   1  '단일 고정
   Caption         =   "환영합니다."
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   Icon            =   "Frm_First.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_First.frx":94E7
   ScaleHeight     =   6825
   ScaleWidth      =   5670
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "Frm_First"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '변수 미선언 방지 선언

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

AlwaysTop Frm_First, True '폼 최상위
WriteLog ("AlwaysTop Err_First -> True")

WriteLog (" _____                 _                _    ") '로그 write
WriteLog ("/  __ \               | |              | |   ")
WriteLog ("| /  \/ ___  _ __ ___ | |     ___   ___| | __")
WriteLog ("| |    / _ \| '_ ` _ \| |    / _ \ / __| |/ /")
WriteLog ("| \__/\ (_) | | | | | | |___| (_) | (__|   < ")
WriteLog (" \____/\___/|_| |_| |_\_____/\___/ \___|_|\_\")
WriteLog ("                                             ")
WriteLog ("ComLock_Setting was first executed.")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_First"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_First")

End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

AlwaysTop Frm_First, False '폼 최상위 해제
WriteLog ("AlwaysTop Err_First -> False")

If MsgBox("창을 닫으시겠습니까?" & vbCrLf & "" & vbCrLf & "본 창은 최초 실행 1회만 표시됩니다" _
    & vbCrLf & "창을 닫으시기전 초기 ID / PW 를 숙지하신후 닫아주세요", vbCritical + vbYesNo, "창을 닫으시겠습니까?") = vbYes Then
    Unload Me
    WriteLog ("Frm_Main Called")
    Frm_Main.Show
Else
    Cancel = True
End If

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_First"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_First")
End Sub


