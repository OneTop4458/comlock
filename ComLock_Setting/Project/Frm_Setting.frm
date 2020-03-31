VERSION 5.00
Begin VB.Form Frm_Setting 
   BorderStyle     =   1  '단일 고정
   Caption         =   "ComLock 클라이언트 설정"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5670
   Icon            =   "Frm_Setting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_Setting.frx":94E7
   ScaleHeight     =   4005
   ScaleWidth      =   5670
   StartUpPosition =   1  '소유자 가운데
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton6 
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "잠금시 모든창 최소화 해제"
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
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "로그인 시도 횟수 설정"
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "잠금시 차단 프로그램 관리"
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
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "윈도우 시작시 실행 설정"
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
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton4 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1920
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "윈도우 시작시 실행 해제"
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
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton5 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "잠금시 모든창 최소화 설정"
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
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton7 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "UAC 설정"
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
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton8 
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ComLock_Setting 로그 확인"
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
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton9 
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ComLock 로그 확인"
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
Attribute VB_Name = "Frm_Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Result As Integer '시작프로그램 등록 성공 여부 확인용

Private Sub Form_Load()

   On Error GoTo Form_Load_Error
   
WriteLog ("Refresh RW5hYmxlZA.")
If ReadINI("RW5hYmxlZA", "V2luZG93c19TdGFydHVw", Environ$("AppData") & "\System.ini") = True Then '윈도우 시작시 실행 설정되어있을시
    UserControl_CandyButton3.Enabled = False
    UserControl_CandyButton4.Enabled = True
Else
    UserControl_CandyButton3.Enabled = True
    UserControl_CandyButton4.Enabled = False
End If

If ReadINI("RW5hYmxlZA", "TWluaW1peg", Environ$("AppData") & "\System.ini") = True Then '잠금시 모든창 최소화 설정되어있을시
    UserControl_CandyButton5.Enabled = False
    UserControl_CandyButton6.Enabled = True
Else
    UserControl_CandyButton5.Enabled = True
    UserControl_CandyButton6.Enabled = False
End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_Setting")
End Sub

Private Sub Form_Unload(Cancel As Integer)

   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Main Called")
Frm_Main.Show

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Setting")
End Sub


Private Sub UserControl_CandyButton1_Click() '잠금시 차단프로그램 관리 클릭시

   On Error GoTo UserControl_CandyButton1_Click_Error

Frm_Setting.Hide
WriteLog ("Frm_Process_Setting Called")
Frm_Process_Setting.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton2_Click() '로그인 시도 설정 클릭시

   On Error GoTo UserControl_CandyButton2_Click_Error

Frm_Setting.Hide
WriteLog ("Login_Setting Called")
Frm_Login_Setting.Show

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton3_Click() '윈도우 시작시 실행 설정 클릭시
   On Error GoTo UserControl_CandyButton3_Click_Error

WriteLog ("Windows_Startup_I Called")
UserControl_CandyButton3.Enabled = False
UserControl_CandyButton4.Enabled = True
Call WriteINI("RW5hYmxlZA", "V2luZG93c19TdGFydHVw", "True", Environ$("AppData") & "\System.ini")
WriteLog ("V2luZG93c19TdGFydHVw -> True")
If InStr(RegGetSectionValueName("SOFTWARE\Microsoft\Windows\CurrentVersion\Run"), "ComLock") = 0 Then '레지스트리에 시작프로그램으로 등록
    Result = IIf(SHRegWriteString("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ComLock", App.Path & "\" & "ComLock.exe"), 1, 0)
    Call MsgBox("ComLock 이 윈도우 시작시 실행됩니다!", vbInformation, App.Title)
End If
If Result = 0 Then
    Frm_Setting.Hide
    WriteLog ("Err_File Called")
    Err_File.Show
End If

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton3_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton3_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton4_Click() '윈도우 시작시 실행 해제 클릭시
   On Error GoTo UserControl_CandyButton4_Click_Error

WriteLog ("Windows_Startup_D Called")
UserControl_CandyButton3.Enabled = True
UserControl_CandyButton4.Enabled = False
Call WriteINI("RW5hYmxlZA", "V2luZG93c19TdGFydHVw", "False", Environ$("AppData") & "\System.ini")
WriteLog ("V2luZG93c19TdGFydHVw -> False")
If InStr(RegGetSectionValueName("SOFTWARE\Microsoft\Windows\CurrentVersion\Run"), "ComLock") Then '레지스트리에 시작프로그램으로 등록 삭제
    Result = IIf(SHRegDelValue("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "ComLock"), 1, 0)
    Call MsgBox("ComLock 이 윈도우 시작시 실행안함으로 설정되었습니다.!", vbInformation, App.Title)
End If
If Result = 0 Then
    Frm_Setting.Hide
    WriteLog ("Err_File Called")
    Err_File.Show
End If


   On Error GoTo 0
   Exit Sub

UserControl_CandyButton4_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton4_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton4_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton5_Click() '잠금시 모든창 최소화 설정 클릭시
   On Error GoTo UserControl_CandyButton5_Click_Error

WriteLog ("Minimize_I Called")
UserControl_CandyButton5.Enabled = False
UserControl_CandyButton6.Enabled = True
Call WriteINI("RW5hYmxlZA", "TWluaW1peg", "True", Environ$("AppData") & "\System.ini")
WriteLog ("TWluaW1peg -> True")
Call MsgBox("설정되었습니다.!", vbInformation, App.Title)

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton5_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton5_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton5_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton6_Click() '잠금시 모든창 최소화 해제 클릭시
   On Error GoTo UserControl_CandyButton6_Click_Error

WriteLog ("Minimize_D Called")
UserControl_CandyButton5.Enabled = True
UserControl_CandyButton6.Enabled = False
Call WriteINI("RW5hYmxlZA", "TWluaW1peg", "False", Environ$("AppData") & "\System.ini")
WriteLog ("TWluaW1peg -> False")
Call MsgBox("해지되었습니다.!", vbInformation, App.Title)

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton6_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton6_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton6_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton7_Click() 'UAC 설정 클릭시
   On Error GoTo UserControl_CandyButton7_Click_Error

WriteLog ("UAC_Setting_Help Called")
Shell "explorer.exe https://prolite.tistory.com/1238"

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton7_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton7_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton7_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton8_Click() 'ComLock_Setting 로그 확인 클릭시
Dim sLog As String
   On Error GoTo UserControl_CandyButton8_Click_Error

WriteLog ("ComLock_Setting_Log Called")
Shell "notepad.exe " & App.Path & "\Logs" & "\ComLock_Setting.Log"

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton8_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton8_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton8_Click of Form Frm_Setting")
End Sub

Private Sub UserControl_CandyButton9_Click() 'ComLock 로그 확인 클릭시
   On Error GoTo UserControl_CandyButton9_Click_Error

WriteLog ("ComLock_Log Called")
Shell "notepad.exe " & App.Path & "\Logs" & "\ComLock.Log"

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton9_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton9_Click of Form Frm_Setting"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton9_Click of Form Frm_Setting")
End Sub

