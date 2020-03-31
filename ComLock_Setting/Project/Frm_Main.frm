VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  '단일 고정
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
   StartUpPosition =   1  '소유자 가운데
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
      Caption         =   "프로그램 종료"
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
      Caption         =   "도움말"
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
      Caption         =   "ID / PW 설정"
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
      Caption         =   "ComLock 클라이언트 설정"
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
Option Explicit '변수 미선언 방지 선언
Private Declare Function IsUserAnAdmin Lib "Shell32" () As Long '관리자 권한 실행 검사 함수 호출

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

If IsUserAnAdmin = 1 Then '관리자 실행 여부 확인
    WriteLog ("[Success] Check Administrator has run")
    On Error GoTo FileErr '에러 발생시 FileErr 이동
        If Dir(Environ$("AppData") & "\System.ini") = vbNullString Then '프로그램 실행에 필요한 INI없을시
            WriteLog ("[Failed] Check System.ini")
            GoTo FileErr
        Else
            WriteLog ("[Success] Check System.ini")
                If ReadINI("Y2tm", "Rmlyc3Q", Environ$("AppData") & "\System.ini") = "VHJ1ZQ" Then '제품 최초 실행일시
                    SaveSetting "System", "root", "SUQ=", "8c7af77e178c5e6b8ede8217fc6859d5" '초기 ID 부여
                    SaveSetting "System", "root", "UFc=", "5a690d842935c51f26f473e025c1b97a" '초기 PW 부여
                    MsgBox "제품 최초 실행이 감지되었습니다.", vbInformation, "안내!"
                    Frm_Main.Hide
                    WriteLog ("Frm_First Called")
                    Frm_First.Show
                    Call WriteINI("Y2tm", "Rmlyc3Q", "RmFsc2U", Environ$("AppData") & "\System.ini") '최초 실행값 변경
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

MsgBox "프로그램을 종료합니다!", vbInformation, "EXIT"
End

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_Main")
End Sub

Private Sub UserControl_CandyButton1_Click() 'ID/PW 설정 클릭시
   On Error GoTo UserControl_CandyButton1_Click_Error

Call WriteINI("R29Ubw", "RnJtQ2hhbmdl", "True", Environ$("AppData") & "\System.ini") 'INI 에 Frm_Login 에서 어디로 이동할지 기록
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

Private Sub UserControl_CandyButton2_Click() 'ComLock 클라이언트 설정 클릭시
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

Private Sub UserControl_CandyButton3_Click() '도움말 버튼 클릭시

   On Error GoTo UserControl_CandyButton3_Click_Error

WriteLog ("ComLock Help Called")

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton3_Click of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton3_Click of Form Frm_Main")
End Sub

Private Sub UserControl_CandyButton4_Click() '프로그램 종료 클릭시
   On Error GoTo UserControl_CandyButton4_Click_Error

WriteLog ("[Success] The ComLock_Setting has successfully terminated.")

MsgBox "프로그램을 종료합니다!", vbInformation, "EXIT"
End

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton4_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton4_Click of Form Frm_Main"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton4_Click of Form Frm_Main")
End Sub

