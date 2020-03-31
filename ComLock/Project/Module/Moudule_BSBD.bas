Attribute VB_Name = "Module_BSBD"
'---------------------------------------------------------------------------------------
' Module    : Module_BSBD
' Author    : http://cafe.daum.net/_c21_/bbs_search_read?grpid=1EgTQ&fldid=4Yrp&datanum=67
' Date      : 2019-05-12
' Purpose   :
'---------------------------------------------------------------------------------------

'�� ����� ���α׷� ���� ����� ��罺ũ�� �߻� ��� �Դϴ�
'�� ����� ������ ���α׷��� ������ �߿����α׷����� �����Ͽ�
'���������� ��η� ���α׷� ����� ��罺ũ���� �߻���ŵ�ϴ�
'�ε� ���ҽ��� �ǿ��ϴ����� ������ ��� ��Ź�帳�ϴ�.
'��� ����� ������Ʈ �ε�� ProtectProcess �Է�
'������Ʈ ��ε�� RestoreProcess �Է� ���ֽø� �ǰ�
'���������� ��η� �����ϴ� ������ RestoreProcess �� ���ְ� �����ߴ³� ���η� �Ǵܵ˴ϴ�

Option Explicit

' ### Ư�� Ȱ��ȭ �ڵ�
Private Declare Function RtlAdjustPrivilege Lib "ntdll" ( _
    ByVal Privilege As Long, _
    ByVal bEnablePrivilege As Long, _
    ByVal IsThreadPrivilege As Long, _
    ByRef PreviousValue As Long _
) As Long

' ### �Ӱ� ���μ��� ����
Private Declare Function RtlSetProcessIsCritical Lib "ntdll" ( _
    ByVal NewValue As Long, _
    ByRef OldValue As Long, _
    ByVal IsWinlogon As Long _
) As Long

' ### ��ó�� ���� �ڵ鷯 ����
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32.dll" ( _
    ByVal lpTopLevelExceptionFilter As Long _
) As Long

' ### ������ ���� API
Private Declare Function CreateThread Lib "kernel32.dll" ( _
    ByRef lpThreadAttributes As Any, _
    ByVal dwStackSize As Long, _
    ByVal lpStartAddress As Long, _
    ByRef lpParameter As Any, _
    ByVal dwCreationFlags As Long, _
    ByRef lpThreadId As Long _
) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" ( _
    ByVal hHandle As Long, _
    ByVal dwMilliseconds As Long _
) As Long
Private Declare Function GetExitCodeThread Lib "kernel32.dll" ( _
    ByVal hThread As Long, _
    ByRef lpExitCode As Long _
) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long _
) As Long
Private Const INFINITE& = &HFFFFFFFF
Private Const WAIT_OBJECT_0& = 0&
Private Const SeDebugPrivilege& = 20&
Private OldSEH As Long, OldValue As Long, Protected As Boolean

Public Function ProtectProcess() As Boolean
    On Error GoTo Failed

    ' ### �̹� ��ȣ�Ǿ� �ִٸ� Ż��
    If Protected Then ProtectProcess = True: Exit Function

    ' ### IDE�� ��� Ż��
    If App.LogMode = 0& Then Exit Function

    ' ### ����� Ư���� ��´�.
    If RtlAdjustPrivilege(SeDebugPrivilege, 1&, 0&, 0&) >= 0& Then
        ' ### ��ó�� ���� �ڵ鷯 ����
        ' ### (API ���ٰ� ���� ���� ������ �𸣹Ƿ�...)
        OldSEH = SetUnhandledExceptionFilter(AddressOf SafeSEH)

        ' ### �Ӱ� ���μ��� ����
        If RtlSetProcessIsCritical(1&, OldValue, 0&) >= 0& Then
            Protected = True: ProtectProcess = True
        End If
    End If

Failed:
End Function

Public Sub RestoreProcess()
    On Error GoTo Failed
    ' ### �����·� ����
    SetUnhandledExceptionFilter OldSEH
    RtlSetProcessIsCritical OldValue, 0&, 0&
    Protected = False
Failed:
End Sub

Private Function SafeSEH(ByVal pvExceptPointer As Long) As Long
    ' ### �����·� ����
    RestoreProcess

    ' ### �������� ó���� ���� ���� ���� ó���� �Լ��� ȣ�����ش�.
    If OldSEH Then
        Dim hThread As Long, retVal As Long
        hThread = CreateThread(ByVal 0&, 0&, OldSEH, ByVal pvExceptPointer, 0&, 0&)
        If WaitForSingleObject(hThread, INFINITE) = WAIT_OBJECT_0 Then
            If GetExitCodeThread(hThread, retVal) Then
                SafeSEH = retVal
            End If
        End If
        CloseHandle hThread
    End If
End Function

