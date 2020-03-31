VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  '���� ����
   Caption         =   "���� ��ǻ�Ͱ� ����ֽ��ϴ�..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7140
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Main.frx":8146
   ScaleHeight     =   2775
   ScaleWidth      =   7140
   StartUpPosition =   1  '������ ���
   Begin VB.Timer Timer_Restore 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6720
      Top             =   360
   End
   Begin VB.Timer Timer_Failed 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6720
      Top             =   0
   End
   Begin VB.Timer Timer_Success 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   0
   End
   Begin VB.Timer Timer_Block 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   0
   End
   Begin VB.Timer Timer_Tray 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   360
   End
   Begin VB.Timer Timer_Restore_Block 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5760
      Top             =   360
   End
   Begin VB.TextBox PW 
      Height          =   390
      IMEMode         =   3  '��� ����
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2115
      Width           =   3850
   End
   Begin VB.TextBox ID 
      Height          =   390
      IMEMode         =   3  '��� ����
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1613
      Width           =   3850
   End
   Begin ComLock.UserControl_CandyButton UserControl_CandyButton 
      Height          =   1215
      Left            =   5760
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�α���"
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
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function IsUserAnAdmin Lib "Shell32" () As Long '������ ���� ���� �˻� �Լ� ȣ��
Dim md5Test As MD5 'MD5 ��ȣȭ ��� ����
Dim Time As Integer '���� �ð� ��� ����
Dim Block As Integer '�α��� �õ� Ƚ�� ����
Dim Desktop As Boolean '���â �ּ�ȭ

Private Sub Form_Load()

   On Error GoTo Form_Load_Error
   Set md5Test = New MD5 'md5 ���� ����
   Time = ReadINI("TG9naW5fU2V0dGluZw", "VGltZQ", Environ$("AppData") & "\System.ini") 'INI ���� ���� �ð� �ҷ���
   Block = ReadINI("TG9naW5fU2V0dGluZw", "bnVtYmVyIG9mIHRpbWVz", Environ$("AppData") & "\System.ini") 'INI ���� �õ� Ƚ�� �ҷ���
   Desktop = ReadINI("RW5hYmxlZA", "TWluaW1peg", Environ$("AppData") & "\System.ini") 'INI ���� ���â �ּ�ȭ �� �ҷ���
   AlwaysTop Frm_Main, True '�� �ֻ���
   ProtectProcess 'ũ��Ƽ�� ���μ��� ���
   'HideMyProcess '���μ��� ���� ����
   
If IsUserAnAdmin = 1 Then '������ ���� ���� Ȯ��
    WriteLog ("[Success] Check Administrator has run")
    On Error GoTo FileErr
        If Dir(Environ$("AppData") & "\System.ini") = vbNullString Then '���α׷� ���࿡ �ʿ��� INI������
            WriteLog ("[Failed] Check System.ini")
            GoTo FileErr
        Else '���� �����
            WriteLog ("[Success] Check System.ini")
            WriteLog ("[Success] ComLock has run.")
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
Cancel = 1
End Sub

Private Sub Timer_Block_Timer() '�α��� ���ܽ� Ÿ�̸�
ID.Enabled = False
PW.Enabled = False
ID.Visible = False
PW.Visible = False
Label1.Visible = True
Label1.Enabled = True
Label1.Caption = "�α����� ���ܵǾ����ϴ�!" & vbCrLf & "���߿� �ٽýõ� �Ͻʽÿ�."
UserControl_CandyButton.Enabled = False
Timer_Restore_Block.Enabled = True
End Sub

Private Sub Timer_Restore_Block_Timer() '�α��� ���� ���� Ÿ�̸� (INI ���� �ð� ����)
'Timer ���͹� �ִ밪 65535

Time = Time - 1 'Ÿ�̸Ӱ� �ѹ� �������� 1�о� ����

If Time = 0 Then 'Time �� 0 �� ������ �и�ŭ �ٵ���
    Timer_Block.Enabled = False
    ID.Enabled = True
    PW.Enabled = True
    ID.Visible = True
    PW.Visible = True
    Label1.Visible = False
    Label1.Enabled = False
    UserControl_CandyButton.Enabled = True
    Block = ReadINI("TG9naW5fU2V0dGluZw", "bnVtYmVyIG9mIHRpbWVz", Environ$("AppData") & "\System.ini") 'block �� �ʱ�ȭ
    Time = ReadINI("TG9naW5fU2V0dGluZw", "VGltZQ", Environ$("AppData") & "\System.ini") 'time �� �ʱ�ȭ
    Timer_Restore_Block.Enabled = False
End If
End Sub

Private Sub Timer_Success_Timer() '�α��� ������ Ÿ�̸�
ID.Enabled = False
PW.Enabled = False
ID.Visible = False
PW.Visible = False
Label1.Visible = True
Label1.Enabled = True
Label1.Caption = "����� Ʈ���� ���� ��ȯ�մϴ�.."
UserControl_CandyButton.Enabled = False
Timer_Tray.Enabled = True
End Sub

Private Sub Timer_Failed_Timer() '�α��� Ʋ���� Ÿ�̸�
ID.Enabled = False
PW.Enabled = False
ID.Visible = False
PW.Visible = False
Label1.Visible = True
Label1.Enabled = True
Label1.Caption = "ID/PW �� Ʋ���ϴ�."
UserControl_CandyButton.Enabled = False
Timer_Restore.Enabled = True
End Sub

Private Sub Timer_Restore_Timer() '�α��� ���� Ÿ�̸� (2��)
Timer_Failed.Enabled = False
ID.Enabled = True
PW.Enabled = True
ID.Visible = True
PW.Visible = True
Label1.Visible = False
Label1.Enabled = False
UserControl_CandyButton.Enabled = True
Timer_Restore.Enabled = False
End Sub

Private Sub UserControl_CandyButton_Click()

WriteLog ("[Warning] Login block " & Block & " times left.")
ID.Text = LCase(md5Test.DigestStrToHexStr(ID.Text)) '�ؽ�Ʈ�� ��ȣȭ
PW.Text = LCase(md5Test.DigestStrToHexStr(PW.Text))
If GetSetting("System", "root", "SUQ=") = ID.Text And GetSetting("System", "root", "UFc=") = PW.Text Then
    Timer_Success.Enabled = True
    WriteLog ("[Success] Timer_Success -> True")
    ID.Text = vbNullString
    PW.Text = vbNullString
    WriteLog ("[Success] Manager Authentication Successful!.")
Else
    If Block = 0 Then '�α��� Ƚ���� 0�̸�
        Timer_Block.Enabled = True
        WriteLog ("[Success] Timer_Block -> True")
        ID.Text = vbNullString
        PW.Text = vbNullString
        WriteLog ("[Failed] Manager Authentication Failed")
    Else
        Timer_Failed.Enabled = True
        WriteLog ("[Success] Timer_Failed -> True")
        ID.Text = vbNullString
        PW.Text = vbNullString
        WriteLog ("[Failed] Manager Authentication Failed")
        Block = Block - 1 '�α��� �õ��ø��� �α��� Ƚ�� ����
    End If
End If
End Sub

Private Sub PW_KeyPress(KeyAscii As Integer)
   On Error GoTo PW_KeyPress_Error

If KeyAscii = 13 Then
    UserControl_CandyButton_Click
End If

   On Error GoTo 0
   Exit Sub

PW_KeyPress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PW_KeyPress of Form Frm_Login"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure PW_KeyPress of Form Frm_Login")
End Sub
