Attribute VB_Name = "modSQLConnect"
Option Compare Database
Option Explicit

' ============================================================
' DevPulse -- SQL Server Connection Management
'
' Originally written by Doug Steele, MVP  AccessMVPHelp@gmail.com
' Adapted by George Hepworth, MVP for DevPulse standalone distribution:
'   - StrConnectionString: added TrustServerCertificate and Authentication
'   - InitConnect: consolidated with StrConnectionString; validates before writing
'   - FixConnections: drives from USysLinkedTables (safer on error than TableDefs loop)
'   - GenerateIndexSQL: unchanged
' ============================================================


Public Function StrConnectionString( _
    ByVal UID As String, _
    ByVal PWD As String) As String

Dim strDrvr As String
Dim strSvr  As String
Dim strDB   As String
Dim strApp  As String

    strDrvr = Nz(DLookup("Drvr", "USysConnect"), "ODBC Driver 18 for SQL Server")
    strSvr = Nz(DLookup("ServerAddress", "USysConnect"), "")
    strDB = Nz(DLookup("DatabaseName", "USysConnect"), "DevPulse")
    strApp = Nz(DLookup("AppName", "USysConnect"), "DevPulse")

    StrConnectionString = _
        "ODBC;DRIVER={" & strDrvr & "};" & _
        "Encrypt=yes;TrustServerCertificate=yes;" & _
        "Trusted_Connection=no;Authentication=SqlPassword;" & _
        "Server=" & strSvr & ";" & _
        "Database=" & strDB & ";" & _
        "PWD=" & PWD & ";" & _
        "UID=" & UID & ";" & _
        "APP=" & strApp & ";"

End Function


Public Function InitConnect( _
    ByVal ServerAddress As String, _
    ByVal DatabaseName As String, _
    ByVal UserName As String, _
    ByVal Password As String, _
    ByVal FixConnectionsFlag As Boolean) As Boolean

100       On Error GoTo errHandler

      Dim db      As DAO.Database
      Dim qdf     As DAO.QueryDef
      Dim rst     As DAO.Recordset
      Dim strConn As String

110       InitConnect = False

          ' Build a test connection string directly from the supplied params
          ' (not from USysConnect, which is not yet written)
120       strConn = "ODBC;DRIVER={" & _
              Nz(DLookup("Drvr", "USysConnect"), "ODBC Driver 18 for SQL Server") & "};" & _
              "Encrypt=yes;TrustServerCertificate=yes;" & _
              "Trusted_Connection=no;Authentication=SqlPassword;" & _
              "Server=" & ServerAddress & ";" & _
              "Database=" & DatabaseName & ";" & _
              "UID=" & UserName & ";" & _
              "PWD=" & Password & ";"

          ' Validate credentials with a lightweight passthrough query
130       Set db = DBEngine(0)(0)
140       Set qdf = db.CreateQueryDef("")
150       With qdf
160           .Connect = strConn
170           .SQL = "SELECT @@OPTIONS;"
180           Set rst = .OpenRecordset(dbOpenSnapshot, dbSQLPassThrough)
190       End With

          ' Credentials valid -- write to USysConnect for future opens
200       CurrentDb.Execute _
              "UPDATE USysConnect SET ServerAddress='" & ServerAddress & "', " & _
              "DatabaseName='" & DatabaseName & "', " & _
              "UID='" & UserName & "', PWD='" & Password & "'", dbFailOnError

210       InitConnect = True

220       If FixConnectionsFlag Then
230           FixConnections UID:=UserName, PWD:=Password
240       End If

Cleanup:

250       On Error Resume Next
260       If Not rst Is Nothing Then rst.Close: Set rst = Nothing
270       If Not qdf Is Nothing Then qdf.Close: Set qdf = Nothing
280       Set db = Nothing
290       Exit Function

errHandler:

          Select Case Err.Number
              Case 18456, 3146, 3151
                  ' Login failure or ODBC connection error -- return False silently;
                  ' frmSetupConnection shows the user-facing message
300               Resume Cleanup
              Case Else
310               Call GlblErrMsg(iLn:=Erl, _
                      sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
                      sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
                      bLog:=True)
320               Resume Cleanup
          End Select
330       Resume

End Function


Public Function FixConnections( _
    ByVal UID As String, _
    ByVal PWD As String) As Boolean

          ' Originally written by Doug Steele, MVP  AccessMVPHelp@gmail.com
          ' Adapted by George Hepworth, MVP: drives from USysLinkedTables rather than
          ' iterating all TableDefs, so no table information is lost if an error occurs
          ' after tables have already been deleted from the TableDefs collection.

100       On Error GoTo errHandler

      Dim db      As DAO.Database
      Dim tdf     As DAO.TableDef
      Dim rst     As DAO.Recordset
      Dim prp     As DAO.Property
      Dim strConn As String
      Dim strName As String
      Dim strSrc  As String
      Dim strIdx  As String
      Dim strDesc As String

110       FixConnections = False
120       strConn = StrConnectionString(UID:=UID, PWD:=PWD)
130       Set db = DBEngine.Workspaces(0).Databases(0)
140       Set rst = db.OpenRecordset( _
              "SELECT TableName, SourceName FROM USysLinkedTables ORDER BY SortOrder", dbOpenSnapshot)

150       Do While Not rst.EOF
160           strName = CStr(rst!TableName)
170           strSrc = Nz(rst!SourceName, "dbo." & strName)
180           strIdx = GenerateIndexSQL(strName)
190           strDesc = ""

200           On Error Resume Next
210           strDesc = CStr(db.TableDefs(strName).Properties("Description"))
220           db.TableDefs.Delete strName
230           On Error GoTo errHandler

240           Set tdf = db.CreateTableDef(strName)
250           tdf.Connect = strConn
260           tdf.SourceTableName = strSrc
270           tdf.Attributes = DAO.dbAttachSavePWD
280           db.TableDefs.Append tdf

290           If Len(strDesc) > 0 Then
300               Set prp = tdf.CreateProperty("Description", dbText, strDesc)
310               tdf.Properties.Append prp
320               Set prp = Nothing
330           End If

340           If Len(strIdx) > 0 Then
350               db.Execute Query:=strIdx, Options:=dbFailOnError
360           End If

370           Set tdf = Nothing
380           rst.MoveNext
390       Loop

400       db.TableDefs.Refresh
410       FixConnections = True

Cleanup:

420       On Error Resume Next
430       If Not rst Is Nothing Then rst.Close: Set rst = Nothing
440       Set tdf = Nothing
450       Set db = Nothing
460       Exit Function

errHandler:

470       Select Case Err.Number
              Case 3270
480               Resume Next
490           Case 3291
500               MsgBox "Problem creating index: " & strIdx, vbOKOnly + vbCritical, "DevPulse Setup"
510               Resume Cleanup
520           Case 18456
530               MsgBox "Wrong User ID or Password.", vbOKOnly + vbCritical, "DevPulse Setup"
540               Resume Cleanup
550           Case Else
560               Call GlblErrMsg(iLn:=Erl, _
                      sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
                      sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
                      bLog:=True)
570               Resume Cleanup
580       End Select
590       Resume

End Function


Public Function GenerateIndexSQL(ByVal TableName As String) As String

    ' Originally written by Doug Steele, MVP  AccessMVPHelp@gmail.com

    On Error GoTo errHandler

    Dim db      As DAO.Database
    Dim tdf     As DAO.TableDef
    Dim idx     As DAO.Index
    Dim fld     As DAO.Field
    Dim strSQL  As String

100     Set db = CurrentDb()
110     Set tdf = db.TableDefs(TableName)

120     If tdf.Indexes.Count > 0 Then
130         On Error Resume Next
140         Set idx = tdf.Indexes("__uniqueindex")
150         If Err.Number = 0 Then
160             On Error GoTo errHandler
170             If idx.Fields.Count > 0 Then
180                 strSQL = "CREATE INDEX __UniqueIndex ON [" & TableName & "] ("
190                 For Each fld In idx.Fields
200                     strSQL = strSQL & "[" & fld.Name & "], "
210                 Next fld
220                 strSQL = Left$(strSQL, Len(strSQL) - 2) & ")"
230             End If
240         End If
250     End If

Cleanup:

260     On Error Resume Next
270     Set fld = Nothing
280     Set idx = Nothing
290     Set tdf = Nothing
300     Set db = Nothing
310     GenerateIndexSQL = strSQL
320     Exit Function

errHandler:

330     If Err.Number <> AppErr.ItemnotFound Then
340         Call GlblErrMsg(iLn:=Erl, _
                sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
                sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
                bLog:=True)
350     End If
360     Resume Cleanup
370     Resume

End Function
