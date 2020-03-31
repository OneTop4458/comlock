VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  '단일 고정
   Caption         =   "현재 컴퓨터가 잠겨있습니다..."
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
   StartUpPosition =   1  '소유자 가운데
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
      IMEMode         =   3  '사용 못함
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2115
      Width           =   3850
   End
   Begin VB.TextBox ID 
      Height          =   390
      IMEMode         =   3  '사용 못함
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
      Caption         =   "로그인"
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
      Alignment       =   2  '가운데 맞춤
      BackColor       =   &H8000000E&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "굴림"
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
Private Declare Function IsUserAnAdmin Lib "Shell32" () As Long '관리자 권한 실행 검사 함수 호출
Dim md5Test As MD5 'MD5 암호화 모듈 선언
Dim Time As Integer '차단 시간 계산 변수
Dim Block As Integer '로그인 시도 횟수 변수
Dim Desktop As Boolean '모든창 최소화

Private Sub Form_Load()

   On Error GoTo Form_Load_Error
   Set md5Test = New MD5 'md5 변수 선언
   Time = ReadINI("TG9naW5fU2V0dGluZw", "VGltZQ", Environ$("AppData") & "\System.ini") 'INI 에서 차단 시간 불러옴
   Block = ReadINI("TG9naW5fU2V0dGluZw", "bnVtYmVyIG9mIHRpbWVz", Environ$("AppData") & "\System.ini") 'INI 에서 시도 횟수 불러옴
   Desktop = ReadINI("RW5hYmxlZA", "TWluaW1peg", Environ$("AppData") & "\System.ini") 'INI 에서 모든창 최소화 값 불러옴
   AlwaysTop Frm_Main, True '폼 최상위
   ProtectProcess '크리티컬 프로세스 등록
   'HideMyProcess '프로세스 정보 숨김
   
If IsUserAnAdmin = 1 Then '관리자 실행 여부 확인
    WriteLog ("[Success] Check Administrator has run")
    On Error GoTo FileErr
        If Dir(Environ$("AppData") & "\System.ini") = vbNullString Then '프로그램 실행에 필요한 INI없을시
            WriteLog ("[Failed] Check System.ini")
            GoTo FileErr
        Else '정상 실행시
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

Private Sub Timer_Block_Timer() '로그인 차단시 타이머
ID.Enabled = False
PW.Enabled = False
ID.Visible = False
PW.Visible = False
Label1.Visible = True
Label1.Enabled = True
Label1.Caption = "로그인이 차단되었습니다!" & vbCrLf & "나중에 다시시도 하십시오."
UserControl_CandyButton.Enabled = False
Timer_Restore_Block.Enabled = True
End Sub

Private Sub Timer_Restore_Block_Timer() '로그인 차단 복구 타이머 (INI 설정 시간 받음)
'Timer 인터벌 최대값 65535

Time = Time - 1 '타이머가 한번 돌때마다 1분씩 제거

If Time = 0 Then 'Time 이 0 즉 설정된 분만큼 다돌면
    Timer_Block.Enabled = False
    ID.Enabled = True
    PW.Enabled = True
    ID.Visible = True
    PW.Visible = True
    Label1.Visible = False
    Label1.Enabled = False
    UserControl_CandyButton.Enabled = True
    Block = ReadINI("TG9naW5fU2V0dGluZw", "bnVtYmVyIG9mIHRpbWVz", Environ$("AppData") & "\System.ini") 'block 값 초기화
    Time = ReadINI("TG9naW5fU2V0dGluZw", "VGltZQ", Environ$("AppData") & "\System.ini") 'time 값 초기화
    Timer_Restore_Block.Enabled = False
End If
End Sub

Private Sub Timer_Success_Timer() '로그인 성공시 타이머
ID.Enabled = False
PW.Enabled = False
ID.Visible = False
PW.Visible = False
Label1.Visible = True
Label1.Enabled = True
Label1.Caption = "잠시후 트레이 모드로 전환합니다.."
UserControl_CandyButton.Enabled = False
Timer_Tray.Enabled = True
End Sub

Private Sub Timer_Failed_Timer() '로그인 틀릴시 타이머
ID.Enabled = False
PW.Enabled = False
ID.Visible = False
PW.Visible = False
Label1.Visible = True
Label1.Enabled = True
Label1.Caption = "ID/PW 가 틀립니다."
UserControl_CandyButton.Enabled = False
Timer_Restore.Enabled = True
End Sub

Private Sub Timer_Restore_Timer() '로그인 복구 타이머 (2초)
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
ID.Text = LCase(md5Test.DigestStrToHexStr(ID.Text)) '텍스트값 암호화
PW.Text = LCase(md5Test.DigestStrToHexStr(PW.Text))
If GetSetting("System", "root", "SUQ=") = ID.Text And GetSetting("System", "root", "UFc=") = PW.Text Then
    Timer_Success.Enabled = True
    WriteLog ("[Success] Timer_Success -> True")
    ID.Text = vbNullString
    PW.Text = vbNullString
    WriteLog ("[Success] Manager Authentication Successful!.")
Else
    If Block = 0 Then '로그인 횟수가 0이면
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
        Block = Block - 1 '로그인 시도시마다 로그인 횟수 줄음
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
