VERSION 5.00
Begin VB.Form Frm_Change 
   BorderStyle     =   1  '단일 고정
   Caption         =   "ID/PW 설정"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7080
   Icon            =   "Frm_Change.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Change.frx":94E7
   ScaleHeight     =   2730
   ScaleWidth      =   7080
   StartUpPosition =   1  '소유자 가운데
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton 
      Height          =   855
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
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
      Caption         =   "변경"
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
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox ID 
      Height          =   375
      IMEMode         =   3  '사용 못함
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1560
      Width           =   4695
   End
End
Attribute VB_Name = "Frm_Change"
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Change"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Change")
End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Main Called")
Frm_Main.Show

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Change"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Change")
End Sub

Private Sub UserControl_CandyButton_Click()
   On Error GoTo UserControl_CandyButton_Click_Error

WriteLog ("ID/PW Change Called")
If MsgBox("입력하신 ID / PW 는 다음과 같습니다." & vbCrLf & "ID / PW 를 변경하시겠습니까?" _
    & vbCrLf & "-----------------------------" _
    & vbCrLf & "ID = " & ID.Text & vbCrLf & "PW = " & PW.Text & vbCrLf _
    & "-----------------------------", vbQuestion + vbYesNo, " 확인!") = vbYes Then
    MsgBox "정상적으로 변경되었습니다 !", vbInformation, "성공!"
    ID.Text = LCase(md5Test.DigestStrToHexStr(ID.Text))
    PW.Text = LCase(md5Test.DigestStrToHexStr(PW.Text))
    SaveSetting "System", "root", "SUQ=", ID.Text
    SaveSetting "System", "root", "UFc=", PW.Text
    ID.Text = vbNullString
    PW.Text = vbNullString
    WriteLog ("[Warning] ID/PW Changed")
    Frm_Change.Hide
    WriteLog ("Frm_Main Called")
    Frm_Main.Show
Else
    MsgBox "ID/PW 변경이 취소되었습니다", vbInformation
    ID.Text = vbNullString
    PW.Text = vbNullString
    WriteLog ("ID/PW Cancel Change")
End If

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Change"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Change")
End Sub

Private Sub PW_KeyPress(KeyAscii As Integer)
   On Error GoTo PW_KeyPress_Error

If KeyAscii = 13 Then
    UserControl_CandyButton_Click
End If

   On Error GoTo 0
   Exit Sub

PW_KeyPress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PW_KeyPress of Form Frm_Change"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure PW_KeyPress of Form Frm_Change")
End Sub
