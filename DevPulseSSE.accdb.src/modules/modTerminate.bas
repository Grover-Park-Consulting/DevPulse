Attribute VB_Name = "modTerminate"
Option Compare Database
Option Explicit

' TerminateAccess pattern by Philipp Steifel — private communication.
' Adapted for DevPulse by GPC with confirmation dialog added.

Private m_CloseWithoutTerminate As Boolean

Private Const PROCESS_TERMINATE As Long = &H1

Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr

Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Boolean, _
    ByVal dwProcessId As Long) As LongPtr

Private Declare PtrSafe Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Boolean

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Boolean

' Emergency hard shutdown — kills the Access process immediately via Windows API.
' Use only when Application.Quit hangs or the database is unresponsive.
' No events fire, no save prompts run, unsaved local data is lost.
' SQL Server transactions are rolled back automatically.
Public Sub TerminateAccess()

    If m_CloseWithoutTerminate Then Exit Sub

    If MsgBox("This will force-terminate DevPulse immediately." & vbCrLf & vbCrLf & _
        "Any unsaved data will be lost. Continue?", _
        vbCritical + vbYesNo, "Force Terminate DevPulse") = vbNo Then Exit Sub

Dim processId As Long
Dim processHandle As LongPtr

    processId = GetCurrentProcessId
    processHandle = OpenProcess(PROCESS_TERMINATE, False, processId)

    TerminateProcess processHandle, 0

    CloseHandle processHandle

End Sub

' Call from a normal close path to suppress TerminateAccess if it is somehow reached.
Private Sub cmdCloseWithoutTerminate(ByRef frm As Form)

    m_CloseWithoutTerminate = True
    DoCmd.Close acForm, frm.Name, acSavePrompt

End Sub
