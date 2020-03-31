VERSION 5.00
Begin VB.Form Err_File 
   BorderStyle     =   1  '단일 고정
   Caption         =   "Error! | Code:404 |"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11085
   Icon            =   "Err_File.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Err_File.frx":94E7
   ScaleHeight     =   2700
   ScaleWidth      =   11085
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "Err_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '변수 미선언 방지 선언

Private Sub Form_Load()
AlwaysTop Err_File, True '폼 최상위
WriteLog ("AlwaysTop Err_File -> True")
   On Error GoTo Form_Load_Error

WriteLog ("[Error] Code 404")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Err_File"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Err_File")
End Sub

Private Sub Form_Unload(Cancel As Integer)
AlwaysTop Err_File, False '폼 최상위 해제
WriteLog ("AlwaysTop Err_File -> False")
   On Error GoTo Form_Unload_Error
   
WriteLog ("[Success] The ComLock_Setting has successfully terminated.")
MsgBox "프로그램을 종료합니다!", vbInformation, "EXIT"
End

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Err_File"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Err_File")
End Sub

