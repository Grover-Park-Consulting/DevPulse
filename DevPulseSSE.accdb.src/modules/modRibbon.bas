Attribute VB_Name = "modRibbon"
Option Compare Database
Option Explicit

' Ribbon UI reference — retained for Invalidate() calls if needed
Public gobjRibbon As Object

' ---------------------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As Object)
    Set gobjRibbon = ribbon
End Sub

' ---------------------------------------------------------------------------

Public Function Ribbon_NodeClicks(ByVal intNodeIndex As Variant)
100       On Error GoTo errHandler

110       Select Case CLng(intNodeIndex)
              Case 2:  DoCmd.OpenForm "frmProject", acNormal, , , , acWindowNormal
120           Case 3:  DoCmd.OpenForm "frmSession", acNormal, , , , acWindowNormal
130           Case 4:  DoCmd.OpenReport "rptProjectSummary", acViewReport
140           Case 10: DoCmd.OpenForm "frmLookupProjectArea", acNormal, , , , acWindowNormal
150           Case 11: DoCmd.OpenForm "frmLookupStatus", acNormal, , , , acWindowNormal
160           Case 12: DoCmd.OpenForm "frmLookupStepType", acNormal, , , , acWindowNormal
170           Case 13: DoCmd.OpenForm "frmLookupTool", acNormal, , , , acWindowNormal
180           Case 14: DoCmd.OpenForm "frmLookupOutputType", acNormal, , , , acWindowNormal
190           Case 15: DoCmd.OpenForm "frmAppConfig", acNormal, , , , acWindowNormal
200           Case 21: Call BackupDevPulse
210           Case 99:
220               If MsgBox("Quit DevPulse?", vbQuestion + vbYesNo, "DevPulse") = vbYes Then
230                   Application.Quit
240               End If
250       End Select

Cleanup:
260       Exit Function

errHandler:
270       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
280       Resume Cleanup
290       Resume
End Function

' ---------------------------------------------------------------------------

Public Function BackupDevPulse() As Boolean
100   On Error GoTo errHandler

      Dim fso As Object
      Dim strSourceFile As String
      Dim strBaseName As String
      Dim strExt As String
      Dim strBackupFolder As String
      Dim strBackupFile As String
      Dim strTimestamp As String

110   Set fso = CreateObject("Scripting.FileSystemObject")
120   strSourceFile = CurrentProject.FullName
130   strBaseName = fso.GetBaseName(strSourceFile)
140   strExt = fso.GetExtensionName(strSourceFile)
150   strBackupFolder = CurrentProject.path & "\Backups"
160   strTimestamp = Format(Now(), "YYYYMMDD-HHNNSS")
170   strBackupFile = strBackupFolder & "\" & strBaseName & "_" & strTimestamp & "." & strExt

180   If Not fso.FolderExists(strBackupFolder) Then fso.CreateFolder strBackupFolder

      ' -----------------------------------------------------------------------
      ' SQL Server back end requires no file backup - data lives on the server.
      ' To enable back-end backup for a split accdb design, uncomment and adapt:
      '
      ' Dim strBEFile As String, strBEBackup As String
      ' strBEFile   = CurrentProject.Path & "\DevPulse_be.accdb"
      ' strBEBackup = strBackupFolder & "\DevPulse_be_" & strTimestamp & ".accdb"
      ' If fso.FileExists(strBEFile) Then fso.CopyFile strBEFile, LCase(strBEBackup)
      ' -----------------------------------------------------------------------

190   fso.CopyFile strSourceFile, LCase(strBackupFile)
200   BackupDevPulse = True
210   MsgBox "Front end backed up to:" & vbCrLf & strBackupFile, _
          vbInformation, "Backup Complete"

Cleanup:
220   Set fso = Nothing
230   Exit Function

errHandler:
240   BackupDevPulse = False
250   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
260   Resume Cleanup
270   Resume
End Function
