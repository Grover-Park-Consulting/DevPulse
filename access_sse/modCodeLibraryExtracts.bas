Attribute VB_Name = "modCodeLibraryExtracts"
Option Compare Database
Option Explicit

' ============================================================
' modCodeLibraryExtracts (SQL Server Edition)
'
' Functions extracted from GPC CodeArchive for standalone
' distribution. In production GPC applications these are called
' via the CodeArchive external library reference.
' Source: D:\CodeLibrary\CodeArchive-GPG201964B_Prod.accdb
' Full library available to licensed GPC clients.
'
' Extracted functions (common to all editions):
'   AppErr enum           - Error number constants
'   GlblErrMsg            - Standard error handler
'   WriteErrorLog         - Log errors to local file
'   WhoAmI / HostName     - Environment identification
'   SetTempVar            - Set a TempVar safely
'   TempVarsLong/String/Boolean/Date/Double - TempVar readers
'   InitApplication       - Load USysAppString into TempVars
'   AppString             - Resolve AppString table name
'   New_AppConstants      - AppConstants class factory
'
' Additional functions (SQL Server Edition only):
'   CreateScratchaccdb    - Create local scratch database for lookup cache
'   CreateLookupCache     - Build local lookup table cache (reads tblLookupCacheConfig)
'   MaintainLookupCache   - Cache engine; called by CreateLookupCache only
' ============================================================


' ---- Module-level variables (used by GlblErrMsg) ----

Public strMBTitle As String
Public lngMBBtn As Long

' ---- String constants ----

Public Const strDoubleLine As Variant = vbNewLine & vbNewLine
Public Const strSingleLine As Variant = vbNewLine


' ============================================================
' AppErr Enum — common Access/VBA error number constants
' ============================================================

Public Enum AppErr
    AccessDenied = 5
    actioncancelled = 2501
    BadInputMask = 2279
    CancelledbyUser = 3059
    CantDelete = 2046
    CantFindTable = 3078
    CantDeleteFile = 70
    CantGotoRecord = 2105
    CantOpenFile = 2220
    CantPerformOperation = 3032
    CantSetProp = 3267
    ControlNotFound = 2465
    DBExists = 3204
    DupeIxorPK = 3022
    FileExists = 58
    FileNotFound = 53
    FrmMustbeActive = 2475
    FormNotFound = 2102
    IMExNotFound = 3625
    InvalidFldVal = 2113
    InvalidFormRef = 2452
    InvalidOperation = 3219
    ItemnotFound = 3265
    MissingMcrTlbr = 2485
    MissingObject = 2450
    MissingOptr = 3075
    NoConnection = 3151
    NoControl = 2474
    NoCurrentRecord = 3021
    NoObject = 2467
    NoParent = 2452
    NoRemoteTable = vbObjectError + 2000
    NotSupportedonObjectType = 3251
    NumberInvalid = 2200
    ODBCCalledFailed = 3146
    ObjNotSet = 91
    PermissionDenied = 70
    PropNotFound = 3270
    QueryExists = 3012
    QdefExists = 3012
    RecordIsLocked = 3260
    SrchKeyNotFound = 3709
    TableExists = 3010
    TableLnkExists = 3012
    Tablelocked = 3211
    TooFewParam = 3061
    TypeMismatch = 13
    Unknown = 3316
    Unsupported = 438
    UserCancel = vbObjectError + 1000
    ViolateRIAdd = 3201
    ViolateRIDelete = 3200
    ViolateRIFrm = 3314
End Enum


' ============================================================
' Error handling
' ============================================================

Public Function GlblErrMsg( _
    Optional ByVal iLn As Long, _
    Optional ByVal sFrm As String, _
    Optional ByVal sCtl As String, _
    Optional ByVal bLog As Boolean = False) As Boolean

Dim appc As Appconstants
Dim strErrMessage As String
Dim strSupport As String
Dim errNum As Long
Dim strErrDesc As String
Dim strLineNum As String

    If Err.Number = AppErr.actioncancelled Then Exit Function
    errNum = Err.Number
    strErrDesc = IIf(Len(Err.Description & "") = 0, "Unidentified Error", Err.Description)
    strLineNum = IIf(iLn = 0, "Not provided", CStr(iLn))
    sCtl = IIf(Len(sCtl & "") = 0, "Unidentified Procedure", sCtl)
    sFrm = IIf(Len(sFrm & "") = 0, "Environment", sFrm)

    Set appc = New_AppConstants
    GlblErrMsg = False
    strSupport = Replace(appc.SupportEmail, ",", vbNewLine)
    strMBTitle = appc.MBErr
    lngMBBtn = appc.MBYNBtn

    strErrMessage = "Please report this error to Your Support Person: " & appc.SupportPerson & "." & strDoubleLine & _
        "The Error Number was: """ & errNum & """." & strSingleLine & _
        "The Error Description was: """ & strErrDesc & """." & strDoubleLine & _
        "The error occurred at Line Number: " & strLineNum & strSingleLine & _
        "In procedure: """ & sCtl & """." & strSingleLine & _
        "In module: """ & sFrm & """." & strDoubleLine & _
        "Please submit this error report " & _
        "and a brief description of what you were doing when the error occurred." & strDoubleLine

    If bLog Then Call WriteErrorLog(appc.AppName, errNum, strErrDesc, strLineNum, sCtl, sFrm)

    On Error Resume Next

    GlblErrMsg = True

    ' NOTE: DoCmd.SendObject requires Classic Outlook.
    ' Microsoft plans to retire Classic Outlook ~2026-2028.
    ' When that happens this send will silently fail. Replace
    ' with Outlook Automation or a mailto: shell call at that time.
    If MsgBox(Prompt:=strErrMessage & vbCrLf & vbCrLf & "Send report?", _
              buttons:=lngMBBtn, Title:=strMBTitle & " Send Report?") = vbYes Then
        DoCmd.SendObject ObjectType:=acSendNoObject, To:=strSupport, _
            Subject:="Error Occurred in " & appc.AppName, _
            messagetext:=strErrMessage, templatefile:=False
    Else
        If Nz(appc.booDebug) Then Stop
    End If
End Function


Private Sub WriteErrorLog( _
    ByVal strAppName As String, _
    ByVal lngErrNum As Long, _
    ByVal strErrDesc As String, _
    ByVal strLineNum As String, _
    ByVal strProc As String, _
    ByVal strModule As String)

Dim strLogDir As String
Dim strLogPath As String
Dim intFile As Integer

    On Error Resume Next

    strLogDir = CurrentProject.Path & "\Logs"
    If Len(Dir(strLogDir, vbDirectory)) = 0 Then MkDir strLogDir

    strLogPath = strLogDir & "\ErrorLog_" & Format(Date, "YYYYMM") & ".txt"

    intFile = FreeFile
    Open strLogPath For Append As #intFile
    Print #intFile, String(80, "=")
    Print #intFile, Format(Now, "YYYY-MM-DD HH:MM:SS") & _
        "  |  App: " & strAppName & _
        "  |  User: " & WhoAmI() & _
        "  |  Host: " & HostName()
    Print #intFile, "Error #  " & lngErrNum & "  :  " & strErrDesc
    Print #intFile, "Line: " & strLineNum & _
        "    Procedure: " & strProc & _
        "    Module: " & strModule
    Print #intFile, String(80, "=")
    Print #intFile, ""
    Close #intFile

End Sub


' ============================================================
' Environment identification
' ============================================================

Public Function HostName() As String
    HostName = Environ$("computername")
End Function

Public Function WhoAmI() As String
    WhoAmI = CurrentProject.Path & "\" & CurrentProject.Name
End Function


' ============================================================
' TempVar helpers
' ============================================================

Public Function TempVarsBoolean(ByVal VarName As String) As Boolean
    TempVarsBoolean = Nz(TempVars(VarName), False)
End Function

Public Function TempVarsDouble(ByVal VarName As String) As Double
    TempVarsDouble = Nz(TempVars(VarName), 0)
End Function

Public Function TempVarsDate(ByVal VarName As String) As Date
    If IsDate(TempVars(VarName)) Then
        TempVarsDate = DateValue(Nz(TempVars(VarName), #1/1/2099#))
    Else
        TempVarsDate = #1/1/2099#
    End If
End Function

Public Function TempVarsLong(ByVal VarName As String) As Long
    TempVarsLong = CLng(Nz(TempVars(VarName), 0))
End Function

Public Function TempVarsString(ByVal VarName As String) As String
    TempVarsString = CStr(Nz(TempVars(VarName), vbNullString))
End Function

Public Sub SetTempVar(ByVal TempVarName As String, ByVal TempVarValue As Variant)
100       On Error GoTo errHandler

110       TempVars.Add Name:=TempVarName, Value:=TempVarValue

Cleanup:
120       On Error Resume Next
130       Exit Sub

errHandler:
140       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
150       Resume Cleanup
160       Resume
End Sub


' ============================================================
' Application initialisation
' ============================================================

Public Function New_AppConstants() As Appconstants
    Set New_AppConstants = New Appconstants
End Function

Public Function AppString() As String
100       On Error GoTo errHandler

110       If DCount("*", "tblAppString") >= 0 Then
120           AppString = "tblAppString"
130           TempVars.Add Name:="AppString", Value:="tblAppString"
140       End If

Cleanup:
150       On Error Resume Next
160       Exit Function

errHandler:
170       If Err = AppErr.CantFindTable Then
180           AppString = "USysAppString"
190           TempVars.Add Name:="AppString", Value:="USysAppString"
200           Resume Cleanup
210       End If
220       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
230       Resume Cleanup
240       Resume
End Function

Public Function InitApplication() As Boolean
100       On Error GoTo errHandler

      Dim db      As DAO.Database
      Dim rst     As DAO.Recordset
      Dim strVal  As String

110       Set db = CurrentDb
120       Set rst = db.OpenRecordset(AppString(), dbOpenSnapshot)
130       Do While Not rst.EOF
140           strVal = Nz(rst!AppStringValue, "")
150           TempVars.Add Name:=rst!AppStringName, Value:= _
                  IIf(strVal = "Date", Date, _
                  IIf(strVal = "Year", Year(Date), _
                  IIf(strVal = "Month", Month(Date), _
                  IIf(strVal = "Day", Day(Date), _
                  strVal))))
160           rst.MoveNext
170       Loop
180       InitApplication = True

Cleanup:
190       On Error Resume Next
200       rst.Close
210       Set rst = Nothing
220       Set db = Nothing
230       Exit Function

errHandler:
240       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
250       Resume Cleanup
260       Resume
End Function


' ============================================================
' Lookup cache (SQL Server Edition)
' ============================================================

Public Function CreateScratchaccdb(ByVal strScratchDatabaseName As String, ByVal strScratchTable As String) As String

100       On Error GoTo errHandler

      Dim path      As String
      Dim dbScratch As DAO.Database

NewAccdb:
110       path = CurrentProject.Path & "\" & strScratchDatabaseName & ".accdb"
120       Set dbScratch = CreateDatabase(path, dbLangGeneral)
130       dbScratch.Close
140       Set dbScratch = Nothing

TransferTable:
150       DoCmd.TransferDatabase transferType:=acExport, DatabaseType:="Microsoft Access", _
              DatabaseName:=path, ObjectType:=acTable, Source:=strScratchTable, _
              Destination:=strScratchTable, StructureOnly:=False
160       CreateScratchaccdb = path

Cleanup:
170       On Error Resume Next
180       If Not dbScratch Is Nothing Then dbScratch.Close: Set dbScratch = Nothing
190       Exit Function

errHandler:
200       If Err = 3204 Then    ' Scratch file already exists — skip creation, go straight to copy
210           Resume TransferTable
220       Else
230           Call GlblErrMsg(iLn:=Erl, _
                  sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
                  sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
                  bLog:=True)
240       End If
250       Resume Cleanup
260       Resume
End Function


Public Function CreateLookupCache(ByVal strScratchDB As String) As Boolean

          ' Reads tblLookupCacheConfig from the calling front-end database to build
          ' the list of SQL Server lookup tables to cache locally, then hands off to
          ' MaintainLookupCache (also in this module) to do the work.
          ' To add or remove a cached table, edit tblLookupCacheConfig in the front
          ' end — no VBA changes needed anywhere.

100       On Error GoTo errHandler

      Dim db          As DAO.Database
      Dim rs          As DAO.Recordset
      Dim lngCount    As Long
      Dim i           As Long
      Dim aTmpl()     As String
      Dim aODBC()     As String
      Dim aFields()   As String

110       Set db = CurrentDb
120       Set rs = db.OpenRecordset( _
              "SELECT TemplateName, ODBCTableName, FieldList " & _
              "FROM tblLookupCacheConfig ORDER BY SortOrder", dbOpenSnapshot)

130       rs.MoveLast
140       lngCount = rs.RecordCount
150       rs.MoveFirst

160       ReDim aTmpl(lngCount - 1)
170       ReDim aODBC(lngCount - 1)
180       ReDim aFields(lngCount - 1)

190       Do While Not rs.EOF
200           aTmpl(i) = rs!TemplateName
210           aODBC(i) = rs!ODBCTableName
220           aFields(i) = rs!FieldList
230           i = i + 1
240           rs.MoveNext
250       Loop
260       rs.Close: Set rs = Nothing

270       CreateLookupCache = MaintainLookupCache(strScratchDB, aTmpl, aODBC, aFields)

Cleanup:
280       On Error Resume Next
290       If Not rs Is Nothing Then rs.Close: Set rs = Nothing
300       Set db = Nothing
310       Exit Function

errHandler:
320       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
330       Resume Cleanup
340       Resume
End Function


Public Function MaintainLookupCache( _
    ByVal strScratchDB As String, _
    aTmpl() As String, _
    aODBC() As String, _
    aFields() As String) As Boolean

                ' Engine for local lookup table caching.
                ' Supports both ODBC (SQL Server) and non-ODBC (Access BE) source tables.
                ' Called by CreateLookupCache (also in this module), which supplies the
                ' table list by reading tblLookupCacheConfig in the calling front end.
                ' Do not call this function directly — use CreateLookupCache instead.

100         On Error GoTo errHandler

            Dim strScratchPath  As String
            Dim strTempPath     As String
            Dim db              As DAO.Database
            Dim tdf             As DAO.TableDef
            Dim qdf             As DAO.QueryDef
            Dim rs              As DAO.Recordset
            Dim i               As Long
            Dim j               As Long
            Dim k               As Long
            Dim lngUB           As Long
            Dim alngSvr()       As Long
            Dim strConn         As String
            Dim strSQL          As String
            Dim bLinkExists     As Boolean

110         lngUB = UBound(aODBC)
      '     ReDim alngSvr(lngUB)    ' used only by the commented-out passthrough section below

120         strScratchPath = CurrentProject.Path & "\" & strScratchDB & ".accdb"
130         strTempPath = CurrentProject.Path & "\" & strScratchDB & "_cmp.accdb"
140         Set db = CurrentDb

150         If Dir(strScratchPath) = "" Then GoTo FullCreate

            ' -- Stale link check: if any lcl_ table points to a different path, force a full rebuild.
            ' This handles the case where the accdb has been moved or copied to a new machine or folder.
160         For i = 0 To lngUB
170             For j = 0 To db.TableDefs.Count - 1
180                 If LCase(db.TableDefs(j).Name) = LCase("lcl_" & aODBC(i)) Then
190                     If InStr(1, db.TableDefs(j).Connect, strScratchPath, vbTextCompare) = 0 Then
200                         GoTo FullCreate
210                     End If
220                     Exit For
230                 End If
240             Next j
250         Next i

            ' -- Compact & repair before any comparison --------------------------
260         If Dir(strTempPath) <> "" Then Kill strTempPath
270         DBEngine.CompactDatabase strScratchPath, strTempPath
280         Kill strScratchPath
290         Name strTempPath As strScratchPath

            ' -- Row count comparison uses DCount per table (see note below).
            ' A more efficient alternative would fetch all counts in a single SQL Server
            ' round trip using a UNION ALL passthrough query — one network call instead of n.
            ' That approach is faster on a remote SQL Server, but it only works for ODBC
            ' sources; it fails with Access BE tables because passthrough queries require
            ' an ODBC connection string. In the interest of keeping DevPulse portable across
            ' both backends, the simpler per-table DCount approach is used here instead.
            ' Anyone wishing to restore the optimised path can uncomment the block below
            ' and replace DCount("*", aODBC(i)) in the ElseIf comparison with alngSvr(i).
            '
      '         strConn = db.TableDefs(aODBC(0)).Connect
      '         strSQL = ""
      '         For k = 0 To lngUB
      '             If k > 0 Then strSQL = strSQL & " UNION ALL "
      '             strSQL = strSQL & "SELECT '" & aODBC(k) & "' AS tbl, COUNT(*) AS n FROM " & db.TableDefs(aODBC(k)).SourceTableName
      '         Next k
      '         Set qdf = db.CreateQueryDef("")
      '         qdf.Connect = strConn
      '         qdf.ReturnsRecords = True
      '         qdf.sql = strSQL
      '         Set rs = qdf.OpenRecordset()
      '         Do While Not rs.EOF
      '             For i = 0 To lngUB
      '                 If rs!tbl = aODBC(i) Then alngSvr(i) = rs!n: Exit For
      '             Next i
      '             rs.MoveNext
      '         Loop
      '         rs.Close: Set rs = Nothing
      '         Set qdf = Nothing

            ' -- Per-table: create link if missing, refresh only if count changed -
300         For i = 0 To lngUB
310             bLinkExists = False
320             For j = 0 To db.TableDefs.Count - 1
330                 If db.TableDefs(j).Name = "lcl_" & aODBC(i) Then
340                     bLinkExists = True
350                     Exit For
360                 End If
370             Next j

380             If Not bLinkExists Then
390                 CreateScratchaccdb strScratchDB, aTmpl(i)
400                 Set tdf = db.CreateTableDef("lcl_" & aODBC(i))
410                 tdf.Connect = ";DATABASE=" & strScratchPath
420                 tdf.SourceTableName = aTmpl(i)
430                 db.TableDefs.Append tdf
440                 Set tdf = Nothing
450                 db.TableDefs.Refresh
460                 db.Execute "INSERT INTO [lcl_" & aODBC(i) & "] (" & aFields(i) & ") " & _
                        "SELECT " & aFields(i) & " FROM [" & aODBC(i) & "]", dbFailOnError + dbSeeChanges
470             ElseIf DCount("*", aODBC(i)) <> DCount("*", "lcl_" & aODBC(i)) Then
480                 db.Execute "DELETE * FROM [lcl_" & aODBC(i) & "]", dbFailOnError
490                 db.Execute "INSERT INTO [lcl_" & aODBC(i) & "] (" & aFields(i) & ") " & _
                        "SELECT " & aFields(i) & " FROM [" & aODBC(i) & "]", dbFailOnError + dbSeeChanges
500             End If
                ' Else: counts match — no action needed for this table
510         Next i
520         GoTo Done

FullCreate:
            ' Export each template (local table) to new scratch accdb — local-to-local, no ODBC
530         For i = 0 To lngUB
540             CreateScratchaccdb strScratchDB, aTmpl(i)
550         Next i

            ' Remove any orphaned lcl_ links from a prior session
560         For i = 0 To lngUB
570             On Error Resume Next
580             db.TableDefs.Delete "lcl_" & aODBC(i)
590         Next i
600         db.TableDefs.Refresh
610         On Error GoTo errHandler

            ' Create lcl_ links pointing into the new scratch accdb
620         For i = 0 To lngUB
630             Set tdf = db.CreateTableDef("lcl_" & aODBC(i))
640             tdf.Connect = ";DATABASE=" & strScratchPath
650             tdf.SourceTableName = aTmpl(i)
660             db.TableDefs.Append tdf
670             Set tdf = Nothing
680         Next i
690         db.TableDefs.Refresh

            ' Populate all lcl_ tables from source tables
700         For i = 0 To lngUB
710             db.Execute "INSERT INTO [lcl_" & aODBC(i) & "] (" & aFields(i) & ") " & _
                    "SELECT " & aFields(i) & " FROM [" & aODBC(i) & "]", dbFailOnError + dbSeeChanges
720         Next i

Done:
730         db.TableDefs.Refresh
740         MaintainLookupCache = True

Cleanup:
750         On Error Resume Next
760         If Not rs Is Nothing Then rs.Close: Set rs = Nothing
770         Set qdf = Nothing
780         Set tdf = Nothing
790         Set db = Nothing
800         Exit Function

errHandler:
810         Call GlblErrMsg(iLn:=Erl, _
                sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
                sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
                bLog:=True)
820         Resume Cleanup
830         Resume
End Function
