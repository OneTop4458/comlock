VERSION 5.00
Begin VB.Form Frm_Login 
   BorderStyle     =   1  '단일 고정
   Caption         =   "관리자 확인"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   Icon            =   "Frm_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Login.frx":94E7
   ScaleHeight     =   2790
   ScaleWidth      =   7080
   StartUpPosition =   1  '소유자 가운데
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton 
      Height          =   855
      Left            =   5760
      TabIndex        =   2
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "확인"
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
   Begin VB.TextBox PW 
      Height          =   390
      IMEMode         =   3  '사용 못함
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   4335
   End
   Begin VB.TextBox ID 
      Height          =   390
      IMEMode         =   3  '사용 못함
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Width           =   4335
   End
End
Attribute VB_Name = "Frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '변수 미선언 방지 선언
Dim md5Test As MD5 'MD5 암호화 모듈 선언

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

Set md5Test = New MD5 'md5 변수 선언

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Login"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Login")
End Sub

Private Sub UserControl_CandyButton_Click()
   On Error GoTo UserControl_CandyButton_Click_Error

ID.Text = LCase(md5Test.DigestStrToHexStr(ID.Text)) '텍스트값 암호화
PW.Text = LCase(md5Test.DigestStrToHexStr(PW.Text))
If GetSetting("System", "root", "SUQ=") = ID.Text And GetSetting("System", "root", "UFc=") = PW.Text Then
    MsgBox "관리자 인증 성공!", vbDefaultButton1, "성공!"
        If ReadINI("R29Ubw", "RnJtQ2hhbmdl", Environ$("AppData") & "\System.ini") = True _
            And ReadINI("R29Ubw", "RnJtU2V0dGluZw", Environ$("AppData") & "\System.ini") = False Then
            Frm_Login.Hide
            ID.Text = vbNullString
            PW.Text = vbNullString
            WriteLog ("Frm_Change Called")
            Frm_Change.Show
        Else
            Frm_Login.Hide
            ID.Text = vbNullString
            PW.Text = vbNullString
            WriteLog ("Frm_Setting Called")
            Frm_Setting.Show
        End If
    WriteLog ("[Success] Manager Authentication Successful!.")
Else
    MsgBox "관리자 인증 실패!", vbCritical, "ERROR!"
    ID.Text = vbNullString
    PW.Text = vbNullString
    WriteLog ("[Failed] Manager Authentication Failed")
End If

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Login"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Login")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Main Called")
Frm_Main.Show
Unload Me

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Login"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Login")
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

