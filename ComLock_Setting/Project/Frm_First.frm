VERSION 5.00
Begin VB.Form Frm_First 
   BorderStyle     =   1  '���� ����
   Caption         =   "ȯ���մϴ�."
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
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "Frm_First"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '���� �̼��� ���� ����

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

AlwaysTop Frm_First, True '�� �ֻ���
WriteLog ("AlwaysTop Err_First -> True")

WriteLog (" _____                 _                _    ") '�α� write
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

AlwaysTop Frm_First, False '�� �ֻ��� ����
WriteLog ("AlwaysTop Err_First -> False")

If MsgBox("â�� �����ðڽ��ϱ�?" & vbCrLf & "" & vbCrLf & "�� â�� ���� ���� 1ȸ�� ǥ�õ˴ϴ�" _
    & vbCrLf & "â�� �����ñ��� �ʱ� ID / PW �� �����Ͻ��� �ݾ��ּ���", vbCritical + vbYesNo, "â�� �����ðڽ��ϱ�?") = vbYes Then
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


