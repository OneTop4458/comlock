VERSION 5.00
Begin VB.Form Frm_Login_Setting_I 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "로그인 시도 가능 횟수 설정"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   Icon            =   "Frm_Login_Setting_I.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Login_Setting_I.frx":94E7
   ScaleHeight     =   2850
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Text            =   "단위 (분)"
      Top             =   2200
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Text            =   "단위 (회)"
      Top             =   1460
      Width           =   2535
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton 
      Height          =   1335
      Left            =   5880
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "설정"
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
Attribute VB_Name = "Frm_Login_Setting_I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Setting Called")
Frm_Setting.Show

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Login_Setting_I"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Login_Setting_I")
End Sub


Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub UserControl_CandyButton_Click()
   On Error GoTo UserControl_CandyButton_Click_Error

If MsgBox("입력하신 값은 다음과 같습니다." & vbCrLf & "변경하시겠습니까?" _
    & vbCrLf & "-----------------------------" _
    & vbCrLf & "시도 횟수 (회) = " & Text1.Text _
    & vbCrLf & "차단 시간 (분) = " & Text2.Text _
    & vbCrLf _
    & "-----------------------------", vbQuestion + vbYesNo, " 확인!") = vbYes Then
    MsgBox "정상적으로 변경되었습니다 !", vbInformation, "성공!"
    Call WriteINI("TG9naW5fU2V0dGluZw", "bnVtYmVyIG9mIHRpbWVz", Text1.Text, Environ$("AppData") & "\System.ini") '시도 가능 횟수 저장
    Call WriteINI("TG9naW5fU2V0dGluZw", "VGltZQ", Text2.Text, Environ$("AppData") & "\System.ini") '차단 시간 저장
    WriteLog ("[Warning] Login_Setting Changed")
    Frm_Login_Setting_I.Hide
    WriteLog ("Frm_Setting Called")
    Frm_Setting.Show
Else
    MsgBox "변경이 취소되었습니다", vbInformation
    WriteLog ("Login_Setting Cancel Change")
End If

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Login_Setting_I"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Login_Setting_I")
End Sub
