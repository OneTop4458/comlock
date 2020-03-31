VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Frm_PS 
   BorderStyle     =   1  '단일 고정
   Caption         =   "프로세스 관리"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   Icon            =   "Frm_PS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frm_PS.frx":94E7
   ScaleHeight     =   6705
   ScaleWidth      =   5880
   StartUpPosition =   1  '소유자 가운데
   Begin MSComctlLib.ListView lvProcess 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComLock_Setting.UserControl_CandyButton UserControl_CandyButton1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   6120
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "프로세스 목록 새고고침"
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
      Top             =   6120
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "선택 프로세스 차단"
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
Attribute VB_Name = "Frm_PS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '변수 미선언 방지 선언
'API  상수선언 (경우에따라서 GetCommandLine , GetModuleFIleName 도 필요)
Dim PName As String, PID As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWDEFAULT = 10

Private Sub Form_Load()
Dim Process
Dim lv As ListItem
   On Error GoTo Form_Load_Error
   
With lvProcess.ColumnHeaders

    .Add , , "프로세스", 3900
    .Add , , "프로세스 ID", 1500
    
End With

lvProcess.ListItems.Clear

For Each Process In GetObject("winmgmts:"). _
    ExecQuery("select * from Win32_Process")
    
    Set lv = lvProcess.ListItems.Add(, , Process.Name)
    lv.SubItems(1) = Process.ProcessID
   
Next
WriteLog ("Get Computer Processes List")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_PS"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Frm_PS")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

WriteLog ("Frm_Setting Called")
Frm_Setting.Show

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_PS"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form Frm_PS")
End Sub

Private Sub lvProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
   On Error GoTo lvProcess_ItemClick_Error

PName = Item.Text
PID = Item.SubItems(1)
WriteLog ("Process" & Item.Text & " Click")

   On Error GoTo 0
   Exit Sub

lvProcess_ItemClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lvProcess_ItemClick of Form Frm_PS"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure lvProcess_ItemClick of Form Frm_PS")
End Sub

Private Sub UserControl_CandyButton1_Click()
Dim Process
Dim lv As ListItem
   On Error GoTo UserControl_CandyButton1_Click_Error

MsgBox "현재 실시간 컴퓨터의 프로세스 목록을 새고고침했습니다!", vbInformation, "성공!"

With lvProcess.ColumnHeaders

    .Add , , "프로세스", 3900
    .Add , , "프로세스 ID", 1500
    
End With

lvProcess.ListItems.Clear

For Each Process In GetObject("winmgmts:"). _
    ExecQuery("select * from Win32_Process")
    
    Set lv = lvProcess.ListItems.Add(, , Process.Name)
    lv.SubItems(1) = Process.ProcessID
   
Next
WriteLog ("Get Computer Processes List")

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_PS"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton1_Click of Form Frm_PS")
End Sub

Private Sub UserControl_CandyButton2_Click()
Dim PS_List As String '프로세스 등록가능수 확인 변수
Dim List As String 'INI 리스트 수 정의 변수
   On Error GoTo UserControl_CandyButton2_Click_Error

PS_List = ReadINI("a2lsbA", "PS_List", Environ$("AppData") & "\System.ini") 'INI 에서 PS_List 값 불러옴

If PS_List = 1 Then
    List = "a2lsbDE"
ElseIf PS_List = 2 Then
    List = "a2lsbDI"
ElseIf PS_List = 3 Then
    List = "a2lsbDM"
ElseIf PS_List = 4 Then
    List = "a2lsbDQ"
ElseIf PS_List = 5 Then
    List = "a2lsbDU"
Else
    List = "a2lsbDE"
End If

If PS_List <= 5 Then
    Call WriteINI("a2lsbA", List, PName, Environ$("AppData") & "\System.ini")
    WriteLog ("[Warning] Process" & PName & " registered")
    MsgBox "선택한 프로세스 " & PName & " 이 정상적으로 등록되었습니다", vbDefaultButton1, "등록완료"
    PS_List = PS_List + 1
    Call WriteINI("a2lsbA", "PS_List", PS_List, Environ$("AppData") & "\System.ini")
Else
    Call MsgBox("등록 가능한 프로세스를 초과했습니다" _
                & vbCrLf & "차단 프로세스 초기화후 다시시도하세요!" _
                , vbCritical, "실패!")
    WriteLog ("[Failed] Process" & PName & " registered Failed")
    
End If

   On Error GoTo 0
   Exit Sub

UserControl_CandyButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_PS"
    WriteLog ("[Error] " & Err.Number & " (" & Err.Description & ") in procedure UserControl_CandyButton2_Click of Form Frm_PS")

End Sub
