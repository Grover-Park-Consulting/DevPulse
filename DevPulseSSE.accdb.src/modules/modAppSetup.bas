Attribute VB_Name = "modAppSetup"
Option Compare Database
Option Explicit

' ============================================================
' DevPulse -- Application Setup and Initialization
' ============================================================

Private Function BuildConnStr() As String
    Dim strUID As String
    Dim strPWD As String
    strUID = Nz(DLookup("UID", "USysConnect"), "")
    strPWD = Nz(DLookup("PWD", "USysConnect"), "")
    BuildConnStr = StrConnectionString(UID:=strUID, PWD:=strPWD)
End Function

Private Function IsTableLinked(ByVal strName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If tdf.Name = strName Then
            IsTableLinked = True
            Exit Function
        End If
    Next tdf
End Function

Private Sub LinkSQLTables(ByVal strConn As String)

    Dim db          As DAO.Database
    Dim tdf         As DAO.TableDef
    Dim rst         As DAO.Recordset
    Dim dicExpected As Object
    Dim lngExpected As Long
    Dim lngLinked   As Long
    Dim bFunctional As Boolean

    Set db = CurrentDb

    On Error Resume Next
    Set rst = db.OpenRecordset("SELECT TOP 1 * FROM tblAppConfig", dbOpenSnapshot)
    bFunctional = (Err.Number = 0)
    If bFunctional Then rst.Close
    Set rst = Nothing
    On Error GoTo 0

    If bFunctional Then
        Set dicExpected = CreateObject("Scripting.Dictionary")
        Set rst = db.OpenRecordset("SELECT TableName FROM USysLinkedTables", dbOpenSnapshot)
        Do While Not rst.EOF
            dicExpected(CStr(rst!TableName)) = 0
            rst.MoveNext
        Loop
        rst.Close
        lngExpected = dicExpected.Count
        For Each tdf In db.TableDefs
            If dicExpected.Exists(tdf.Name) Then lngLinked = lngLinked + 1
        Next tdf
        Set dicExpected = Nothing
        If lngLinked = lngExpected Then
            Set db = Nothing
            Exit Sub
        End If
    End If

    Set rst = db.OpenRecordset("SELECT TableName, SourceName FROM USysLinkedTables ORDER BY SortOrder", dbOpenSnapshot)
    Do While Not rst.EOF
        On Error Resume Next
        db.TableDefs.Delete CStr(rst!TableName)
        On Error GoTo 0
        Set tdf = db.CreateTableDef(CStr(rst!TableName))
        tdf.Connect = strConn
        tdf.SourceTableName = Nz(rst!SourceName, "dbo." & CStr(rst!TableName))
        tdf.Attributes = dbAttachSavePWD
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

      Dim strConn As String

          ' First-run detection: no linked tables means connection not yet configured
110       If Not IsTableLinked("tblAppConfig") Then
120           DoCmd.OpenForm "frmSetupConnection", , , , , acDialog
130           If Not IsTableLinked("tblAppConfig") Then
140               Application.Quit
150               Exit Function
160           End If
170       End If

180       strConn = BuildConnStr()
190       Call LinkSQLTables(strConn)
200       Call InitApplication
210       Call CreateLookupCache("DevPulseLookups")
220       InitApp = True

Cleanup:

230       Exit Function

errHandler:

240       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
250       Resume Cleanup
260       Resume

End Function


Public Sub CloseLookupForm(sLocalTable As String)
100       On Error GoTo errHandler

110       CurrentDb.Execute "DELETE FROM " & sLocalTable
120       Call CreateLookupCache("DevPulseLookups")

Cleanup:
130       Exit Sub

errHandler:
140       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
150       Resume Cleanup
160       Resume
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
              "Check your SQL Server connection and try again.", _
              vbCritical, "DevPulse"
150       Resume Cleanup
160       Resume
End Function
