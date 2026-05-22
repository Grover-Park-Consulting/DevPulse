Attribute VB_Name = "modAppSetup"
Option Compare Database
Option Explicit

' ============================================================
' DevPulse — Application Setup (Access Backend Edition)
'
' This version links to a local Access backend (DevPulse_BE.accdb)
' rather than a SQL Server database. No ODBC configuration required.
'
' The backend file must be in the same folder as this front end.
' ============================================================

Private Const cBEFileName As String = "DevPulse_BE.accdb"


Private Function BEPath() As String
    BEPath = CurrentProject.Path & "\" & cBEFileName
End Function


Private Sub RelinkBETables()

    Dim db          As DAO.Database
    Dim tdf         As DAO.TableDef
    Dim rst         As DAO.Recordset
    Dim bFunctional As Boolean
    Dim strBEPath   As String
    Dim strConn     As String

    strBEPath = BEPath()

    If Len(Dir(strBEPath)) = 0 Then
        MsgBox "Cannot find the DevPulse data file:" & vbCrLf & vbCrLf & _
               strBEPath & vbCrLf & vbCrLf & _
               "Ensure DevPulse_BE.accdb is in the same folder as DevPulse.accdb.", _
               vbCritical, "DevPulse"
        Application.Quit
    End If

    strConn = ";DATABASE=" & strBEPath

    Set db = CurrentDb

    On Error Resume Next
    Set rst = db.OpenRecordset("SELECT TOP 1 * FROM tblAppConfig", dbOpenSnapshot)
    bFunctional = (Err.Number = 0)
    If bFunctional Then rst.Close
    Set rst = Nothing
    On Error GoTo 0

    If bFunctional Then
        Set db = Nothing
        Exit Sub
    End If

    Set rst = db.OpenRecordset("SELECT TableName FROM USysLinkedTables ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rst.EOF
        On Error Resume Next
        db.TableDefs.Delete CStr(rst!TableName)
        On Error GoTo 0
        Set tdf = db.CreateTableDef(CStr(rst!TableName))
        tdf.Connect = strConn
        tdf.SourceTableName = CStr(rst!TableName)
        db.TableDefs.Append tdf
        Set tdf = Nothing
        rst.MoveNext
    Loop
    rst.Close
    db.TableDefs.Refresh
    Set db = Nothing

End Sub


Public Function InitApp() As Boolean
100       On Error GoTo errHandler

110       Call RelinkBETables
120       Call InitApplication
130       InitApp = True

Cleanup:
140       Exit Function

errHandler:
150       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
160       Resume Cleanup
170       Resume
End Function


Public Sub CloseLookupForm(sLocalTable As String)
    ' Access BE edition: lookup tables are directly linked — no cache to refresh.
    ' This stub is retained for call-site compatibility.
End Sub


Public Function StartupApp()
100       On Error GoTo errHandler

110       Call InitApp
120       DoCmd.OpenForm "frmMain"

Cleanup:
130       Exit Function

errHandler:
140       MsgBox "DevPulse failed to start." & vbCrLf & vbCrLf & _
              "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
              "Ensure DevPulse_BE.accdb is in the same folder as this file.", _
              vbCritical, "DevPulse"
150       Resume Cleanup
160       Resume
End Function

Private Sub ForceRelink()
    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim rst     As DAO.Recordset
    Dim strConn As String

    strConn = ";DATABASE=" & BEPath()
    Set db = CurrentDb

    Set rst = db.OpenRecordset("SELECT TableName FROM USysLinkedTables ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rst.EOF
        On Error Resume Next
        db.TableDefs.Delete CStr(rst!TableName)
        On Error GoTo 0
        Set tdf = db.CreateTableDef(CStr(rst!TableName))
        tdf.Connect = strConn
        tdf.SourceTableName = CStr(rst!TableName)
        db.TableDefs.Append tdf
        Set tdf = Nothing
        rst.MoveNext
    Loop
    rst.Close
    db.TableDefs.Refresh
    Set db = Nothing
End Sub
