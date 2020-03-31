VERSION 5.00
Begin VB.Form Frm_Login_Setting 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "로그인 시도 횟수 설정"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5820
   Icon            =   "Frm_Login_Setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Login_Setting.frx":94E7
   ScaleHeight     =   3165
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "설정 변경"
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
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Frm_Login_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

WriteLog ("Refresh TG9naW5fU2V0dGluZw.")
Label1.Caption = ReadINI("TG9naW5fU2V0dGluZw", "bnVtYmVyIG9mIHRpbWVz", Environ$("AppData") & "\System.ini") '시도 가능 횟수 읽어옴
Label2.Caption = ReadINI("TG9naW5fU2V0dGluZw", "VGltZQ", Environ$("AppData") & "\System.ini") '차단 시간 읽어옴

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Login_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Login_Setting")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Setting Called")
Frm_Setting.Show

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Login_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Login_Setting")
End Sub

Private Sub UserControl_CandyButton_Click()
   On Error GoTo UserControl_CandyButton_Click_Error

Frm_Login_Setting.Hide
WriteLog ("Login_Setting_I Called")
Frm_Login_Setting_I.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Login_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton_Click of Form Frm_Login_Setting")
End Sub
