VERSION 5.00
Begin VB.Form Err_Permission 
   BorderStyle     =   1  '���� ����
   Caption         =   "Error! | Code:400 |"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11100
   Icon            =   "Err_Permission.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Err_Permission.frx":94E7
   ScaleHeight     =   2790
   ScaleWidth      =   11100
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "Err_Permission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '���� �̼��� ���� ����

Private Sub Form_Load()
AlwaysTop Err_Permission, True '�� �ֻ���
WriteLog ("AlwaysTop Err_Permission -> True")
   On Error GoTo Form_Load_Error

WriteLog ("[Error] Code 400")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Err_Permission"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Err_Permission")
End Sub

Private Sub Form_Unload(Cancel As Integer)
AlwaysTop Err_Permission, False '�� �ֻ��� ����
WriteLog ("AlwaysTop Err_Permission -> False")
   On Error GoTo Form_Unload_Error

WriteLog ("[Success] The ComLock has successfully terminated.")
MsgBox "���α׷��� �����մϴ�!", vbInformation, "EXIT"
End

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Err_Permission"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Err_Permission")
End Sub

