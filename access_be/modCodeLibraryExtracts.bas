Attribute VB_Name = "modCodeLibraryExtracts"
Option Compare Database
Option Explicit

' ============================================================
' modCodeLibraryExtracts
'
' Functions extracted from GPC CodeArchive for standalone
' distribution. In production GPC applications these are called
' via the CodeArchive external library reference.
' Source: D:\CodeLibrary\CodeArchive-GPG201964B_Prod.accdb
' Full library available to licensed GPC clients.
'
' Extracted functions:
'   AppErr enum           - Error number constants
'   GlblErrMsg            - Standard error handler
'   WriteErrorLog         - Log errors to local file
'   WhoAmI / HostName     - Environment identification
'   SetTempVar            - Set a TempVar safely
'   TempVarsLong/String/Boolean/Date/Double - TempVar readers
'   InitApplication       - Load USysAppString into TempVars
'   AppString             - Resolve AppString table name
'   New_AppConstants      - AppConstants class factory
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
