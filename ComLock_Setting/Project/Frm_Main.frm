VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  '���� ����
   Caption         =   "ComLock_Setting"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Main.frx":94E7
   ScaleHeight     =   3015
   ScaleWidth      =   5775
   StartUpPosition =   1  '������ ���
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton4 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "���α׷� ����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton3 
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ID / PW ����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton2 
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ComLock Ŭ���̾�Ʈ ����"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
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
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '���� �̼��� ���� ����
Private Declare Function IsUserAnAdmin Lib "Shell32" () As Long '������ ���� ���� �˻� �Լ� ȣ��

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

If IsUserAnAdmin = 1 Then '������ ���� ���� Ȯ��
    WriteLog ("[Success] Check Administrator has run")
    On Error GoTo FileErr '���� �߻��� FileErr �̵�
        If Dir(Environ$("AppData") & "\System.ini") = vbNullString Then '���α׷� ���࿡ �ʿ��� INI������
            WriteLog ("[Failed] Check System.ini")
            GoTo FileErr
        Else
            WriteLog ("[Success] Check System.ini")
                If ReadINI("Y2tm", "Rmlyc3Q", Environ$("AppData") & "\System.ini") = "VHJ1ZQ" Then '��ǰ ���� �����Ͻ�
                    SaveSetting "System", "root", "SUQ=", "8c7af77e178c5e6b8ede8217fc6859d5" '�ʱ� ID �ο�
                    SaveSetting "System", "root", "UFc=", "5a690d842935c51f26f473e025c1b97a" '�ʱ� PW �ο�
                    MsgBox "��ǰ ���� ������ �����Ǿ����ϴ�.", vbInformation, "�ȳ�!"
                    Frm_Main.Hide
                    WriteLog ("Frm_First Called")
                    Frm_First.Show
                    Call WriteINI("Y2tm", "Rmlyc3Q", "RmFsc2U", Environ$("AppData") & "\System.ini") '���� ���ప ����
                    WriteLog ("[Success] Initial ID / PW granted.")
                    WriteLog ("[Success] ComLock_Setting has run.")
                Else
                    WriteLog ("[Success] ComLock_Setting has run.")
                End If
        End If
Else
    WriteLog ("[Failed] Check Administrator has run")
    Frm_Main.Hide
    WriteLog ("Err_Permission Called")
    Err_Permission.Show
End If

   On Error GoTo 0
   Exit Sub
   
FileErr:

    Frm_Main.Hide
    WriteLog ("Err_File Called")
    Err_File.Show

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Main")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

WriteLog ("[Success] The ComLock_Setting has successfully terminated.")

MsgBox "���α׷��� �����մϴ�!", vbInformation, "EXIT"
End

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Main")
End Sub

Private Sub UserControl_CandyButton1_Click() 'ID/PW ���� Ŭ����
   On Error GoTo UserControl_CandyButton1_Click_Error

Call WriteINI("R29Ubw", "RnJtQ2hhbmdl", "True", Environ$("AppData") & "\System.ini") 'INI �� Frm_Login ���� ���� �̵����� ���
Call WriteINI("R29Ubw", "RnJtU2V0dGluZw", "False", Environ$("AppData") & "\System.ini")
Frm_Main.Hide
WriteLog ("Frm_Login Called")
Frm_Login.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Main")
End Sub

Private Sub UserControl_CandyButton2_Click() 'ComLock Ŭ���̾�Ʈ ���� Ŭ����
   On Error GoTo UserControl_CandyButton2_Click_Error

Call WriteINI("R29Ubw", "RnJtU2V0dGluZw", "True", Environ$("AppData") & "\System.ini")
Call WriteINI("R29Ubw", "RnJtQ2hhbmdl", "False", Environ$("AppData") & "\System.ini")
Frm_Main.Hide
WriteLog ("Frm_Login Called")
Frm_Login.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_Main")
End Sub

Private Sub UserControl_CandyButton3_Click() '���� ��ư Ŭ����

   On Error GoTo UserControl_CandyButton3_Click_Error

WriteLog ("ComLock Help Called")

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton3_Click of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton3_Click of Form Frm_Main")
End Sub

Private Sub UserControl_CandyButton4_Click() '���α׷� ���� Ŭ����
   On Error GoTo UserControl_CandyButton4_Click_Error

WriteLog ("[Success] The ComLock_Setting has successfully terminated.")

MsgBox "���α׷��� �����մϴ�!", vbInformation, "EXIT"
End

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton4_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton4_Click of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton4_Click of Form Frm_Main")
End Sub

