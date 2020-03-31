VERSION 5.00
Begin VB.Form Frm_Process_Setting 
   BorderStyle     =   1  '���� ����
   Caption         =   "���� ��ϵ� ���� ���α׷� ���"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5820
   Icon            =   "Frm_Process_Setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Process_Setting.frx":94E7
   ScaleHeight     =   3735
   ScaleWidth      =   5820
   StartUpPosition =   1  '������ ���
   Begin VB.Timer Refresh_Timer 
      Interval        =   10000
      Left            =   5160
      Top             =   360
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���� ���μ��� �߰��ϱ�"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.ListBox List 
      Height          =   2040
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton2 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���� ���μ��� �ʱ�ȭ"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
End
Attribute VB_Name = "Frm_Process_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '���� �̼��� ���� ����

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

WriteLog ("Refresh Process List.")
List.Clear
List.AddItem ("-------------------- �⺻ ���� ���μ��� ��� -------------------")
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDE", Environ$("AppData") & "\System.ini") ':: �⺻ 1 explorer
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDI", Environ$("AppData") & "\System.ini") ':: �⺻ 2 cmd
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDM", Environ$("AppData") & "\System.ini") ':: �⺻ 3 Taskmgr
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDQ", Environ$("AppData") & "\System.ini") ':: �⺻ 4 perfmon
List.AddItem ("-------------------- ���� ���� ���μ��� ��� -------------------")
List.AddItem ReadINI("a2lsbA", "a2lsbDE", Environ$("AppData") & "\System.ini") ':: ���� 1
List.AddItem ReadINI("a2lsbA", "a2lsbDI", Environ$("AppData") & "\System.ini") ':: ���� 2
List.AddItem ReadINI("a2lsbA", "a2lsbDM", Environ$("AppData") & "\System.ini") ':: ���� 3
List.AddItem ReadINI("a2lsbA", "a2lsbDQ", Environ$("AppData") & "\System.ini") ':: ���� 4
List.AddItem ReadINI("a2lsbA", "a2lsbDU", Environ$("AppData") & "\System.ini") ':: ���� 5

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Process_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of From Frm_Process_Setting")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Setting Called")
Frm_Setting.Show

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Process_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of From Frm_Process_Setting")
End Sub

Private Sub Refresh_Timer_Timer()

   On Error GoTo Refresh_Timer_Refresh_Timer_Error
   
WriteLog ("Refresh Process List.")
List.Clear
List.AddItem ("-------------------- �⺻ ���� ���μ��� ��� -------------------")
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDE", Environ$("AppData") & "\System.ini") ':: �⺻ 1 explorer
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDI", Environ$("AppData") & "\System.ini") ':: �⺻ 2 cmd
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDM", Environ$("AppData") & "\System.ini") ':: �⺻ 3 Taskmgr
List.AddItem ReadINI("a2lsbA", "ZGVmYXVsdDQ", Environ$("AppData") & "\System.ini") ':: �⺻ 4 perfmon
List.AddItem ("-------------------- ���� ���� ���μ��� ��� -------------------")
List.AddItem ReadINI("a2lsbA", "a2lsbDE", Environ$("AppData") & "\System.ini") ':: ���� 1
List.AddItem ReadINI("a2lsbA", "a2lsbDI", Environ$("AppData") & "\System.ini") ':: ���� 2
List.AddItem ReadINI("a2lsbA", "a2lsbDM", Environ$("AppData") & "\System.ini") ':: ���� 3
List.AddItem ReadINI("a2lsbA", "a2lsbDQ", Environ$("AppData") & "\System.ini") ':: ���� 4
List.AddItem ReadINI("a2lsbA", "a2lsbDU", Environ$("AppData") & "\System.ini") ':: ���� 5

   On Error GoTo 0
   Exit Sub

Refresh_Timer_Refresh_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Refresh_Timer_Refresh_Timer of Form Frm_Process_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Refresh_Timer_Refresh_Timer of From Frm_Process_Setting")
End Sub

Private Sub UserControl_CandyButton1_Click()

   On Error GoTo UserControl_CandyButton1_Click_Error

Frm_Process_Setting.Hide
WriteLog ("Frm_PS Called")
Frm_PS.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Process_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Process_Setting")
End Sub

Private Sub UserControl_CandyButton2_Click() '��ϵ� ���μ��� ����Ʈ �ϰ� ���� (������ ���� ���Ǹ� ���� ����)
   On Error GoTo UserControl_CandyButton2_Click_Error

WriteLog ("Process_D Called")
If MsgBox("���ܵ� ���μ����� �ʱ�ȭ �մϱ�?" & vbCrLf & "" & vbCrLf & "�⺻ ���μ��� ���� ���ܵ�ϵ�" _
    & vbCrLf & "��� ���μ����� �������ϴ�.", vbCritical + vbYesNo, "���μ����� �ʱ�ȭ�մϱ�?") = vbYes Then
    Call WriteINI("a2lsbA", "a2lsbDE", "", Environ$("AppData") & "\System.ini") ':: ���� 1
    Call WriteINI("a2lsbA", "a2lsbDI", "", Environ$("AppData") & "\System.ini") ':: ���� 2
    Call WriteINI("a2lsbA", "a2lsbDM", "", Environ$("AppData") & "\System.ini") ':: ���� 3
    Call WriteINI("a2lsbA", "a2lsbDQ", "", Environ$("AppData") & "\System.ini") ':: ���� 4
    Call WriteINI("a2lsbA", "a2lsbDU", "", Environ$("AppData") & "\System.ini") ':: ���� 5
    Call WriteINI("a2lsbA", "PS_List", "1", Environ$("AppData") & "\System.ini") ':: ��ϰ����� ���α׷��� �ʱ�ȭ
    WriteLog ("[Warning] Initializes registered processes.")
    MsgBox "�⺻ ���μ��� �����ϰ�" & vbCrLf & "��� ���μ����� �ʱ�ȭ �߽��ϴ�!" _
    & vbCrLf & "" & vbCrLf & "��� �ݿ����� ��(10��) �ҿ�...", vbInformation, "����!"
Else
End If

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_Process_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_Process_Setting")
End Sub


