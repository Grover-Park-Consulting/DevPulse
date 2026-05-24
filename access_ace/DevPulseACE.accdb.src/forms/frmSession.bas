Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18960
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =9360
    Top =3216
    Right =28080
    Bottom =14412
    OrderBy ="[qrySession].[SessionDate]"
    RecSrcDt = Begin
        0x4974469daf89e640
    End
    RecordSource ="qrySession"
    Caption ="Sessions"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1080
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =90
                    Width =900
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblProject"
                    Caption ="Project"
                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =540
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1350
                    Top =90
                    Width =6390
                    Height =450
                    FontSize =14
                    Name ="cboProjectFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TOP 1 0,'(All Projects)' FROM USysAppString UNION ALL SELECT ProjectID, P"
                        "rojectName FROM tblProject ORDER BY 2"
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =1350
                    LayoutCachedTop =90
                    LayoutCachedWidth =7740
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =7830
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStep"
                    Caption ="Step"
                    LayoutCachedLeft =7830
                    LayoutCachedTop =90
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =540
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8640
                    Top =90
                    Width =6390
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="cboStepFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TOP 1 0,'(All Steps)' FROM USysAppString UNION ALL SELECT s.StepID, IIf(s"
                        "ess.StepID Is Null, '* ' & s.StepName, s.StepName) FROM tblStep s LEFT JOIN (SEL"
                        "ECT DISTINCT StepID FROM tblSession) sess ON s.StepID = sess.StepID WHERE (TempV"
                        "ars!lngProjectID=0 OR s.ProjectID=TempVars!lngProjectID) ORDER BY 1"
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =8640
                    LayoutCachedTop =90
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15120
                    Top =90
                    Width =3240
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="cmdNewSession"
                    Caption ="New Session"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =15120
                    LayoutCachedTop =90
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =630
                    Width =1884
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrDate"
                    Caption ="Date"
                    LayoutCachedLeft =360
                    LayoutCachedTop =630
                    LayoutCachedWidth =2244
                    LayoutCachedHeight =970
                End
                Begin Label
                    OverlapFlags =85
                    Left =2295
                    Top =630
                    Width =1524
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrStart"
                    Caption ="Start"
                    LayoutCachedLeft =2295
                    LayoutCachedTop =630
                    LayoutCachedWidth =3819
                    LayoutCachedHeight =970
                End
                Begin Label
                    OverlapFlags =85
                    Left =3870
                    Top =630
                    Width =1524
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrEnd"
                    Caption ="End"
                    LayoutCachedLeft =3870
                    LayoutCachedTop =630
                    LayoutCachedWidth =5394
                    LayoutCachedHeight =970
                End
                Begin Label
                    OverlapFlags =85
                    Left =5445
                    Top =630
                    Width =900
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrReset"
                    Caption ="Reset"
                    LayoutCachedLeft =5445
                    LayoutCachedTop =630
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =970
                End
                Begin Label
                    OverlapFlags =85
                    Left =6396
                    Top =630
                    Width =1350
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrRevisions"
                    Caption ="Revisions"
                    LayoutCachedLeft =6396
                    LayoutCachedTop =630
                    LayoutCachedWidth =7746
                    LayoutCachedHeight =970
                End
                Begin Label
                    OverlapFlags =85
                    Left =7797
                    Top =630
                    Width =11160
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrNotes"
                    Caption ="Notes"
                    LayoutCachedLeft =7797
                    LayoutCachedTop =630
                    LayoutCachedWidth =18957
                    LayoutCachedHeight =970
                End
            End
        End
        Begin Section
            Height =480
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =30
                    Width =1884
                    Height =450
                    FontSize =14
                    Name ="txtSessionDate"
                    ControlSource ="SessionDate"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x00000000b8dcff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530065007300730069006f006e00490044005d003d00540065006d007000 ,
                        0x76006100720073004c006f006e006700280022006c006e006700430075007200 ,
                        0x720065006e007400730065007300730069006f006e0049004400220029000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =360
                    LayoutCachedTop =30
                    LayoutCachedWidth =2244
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000b8dcff002f0000005b00 ,
                        0x530065007300730069006f006e00490044005d003d00540065006d0070007600 ,
                        0x6100720073004c006f006e006700280022006c006e0067004300750072007200 ,
                        0x65006e007400730065007300730069006f006e00490044002200290000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2295
                    Top =30
                    Width =1524
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="txtSessionStartTime"
                    ControlSource ="SessionStartTime"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x0000000077c0d400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530065007300730069006f006e00490044005d003d00540065006d007000 ,
                        0x56006100720073004c006f006e006700280022006c006e006700430075007200 ,
                        0x720065006e007400530065007300730069006f006e0049004400220029000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =2295
                    LayoutCachedTop =30
                    LayoutCachedWidth =3819
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010000000000000077c0d4002f0000005b00 ,
                        0x530065007300730069006f006e00490044005d003d00540065006d0070005600 ,
                        0x6100720073004c006f006e006700280022006c006e0067004300750072007200 ,
                        0x65006e007400530065007300730069006f006e00490044002200290000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3870
                    Top =30
                    Width =1524
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtSessionEndTime"
                    ControlSource ="SessionEndTime"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000c8000000010000000100000000000000000000003200000001000000 ,
                        0x00000000b8dcff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530065007300730069006f006e00490044005d0020003d00200054006500 ,
                        0x6d00700056006100720073004c006f006e006700280022006c006e0067004300 ,
                        0x75007200720065006e007400530065007300730069006f006e00490044002200 ,
                        0x2900000000000000
                    End

                    LayoutCachedLeft =3870
                    LayoutCachedTop =30
                    LayoutCachedWidth =5394
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000b8dcff00310000005b00 ,
                        0x530065007300730069006f006e00490044005d0020003d002000540065006d00 ,
                        0x700056006100720073004c006f006e006700280022006c006e00670043007500 ,
                        0x7200720065006e007400530065007300730069006f006e004900440022002900 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5445
                    Top =30
                    Width =900
                    Height =450
                    TabIndex =3
                    Name ="chkContextReset"
                    ControlSource ="ContextReset"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =5445
                    LayoutCachedTop =30
                    LayoutCachedWidth =6345
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6396
                    Top =30
                    Width =1350
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtRevisionCycles"
                    ControlSource ="RevisionCycles"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000c2000000010000000100000000000000000000003000000001000000 ,
                        0x00000000b8dcff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530065007300730069006f006e00490044005d003d00540065006d007000 ,
                        0x56006100720073004c006f006e006700280022006c006e006700430075007200 ,
                        0x720065006e007400530065007300730069006f006e0049004400220029000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =6396
                    LayoutCachedTop =30
                    LayoutCachedWidth =7746
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000b8dcff002f0000005b00 ,
                        0x530065007300730069006f006e00490044005d003d00540065006d0070005600 ,
                        0x6100720073004c006f006e006700280022006c006e0067004300750072007200 ,
                        0x65006e007400530065007300730069006f006e00490044002200290000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7797
                    Top =30
                    Width =11160
                    Height =450
                    FontSize =14
                    TabIndex =5
                    Name ="txtNotes"
                    ControlSource ="Notes"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000c8000000010000000100000000000000000000003200000001000000 ,
                        0x00000000b8dcff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00530065007300730069006f006e00490044005d0020003d00200054006500 ,
                        0x6d00700056006100720073004c006f006e006700280022006c006e0067004300 ,
                        0x75007200720065006e007400530065007300730069006f006e00490044002200 ,
                        0x2900000000000000
                    End

                    LayoutCachedLeft =7797
                    LayoutCachedTop =30
                    LayoutCachedWidth =18957
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000b8dcff00310000005b00 ,
                        0x530065007300730069006f006e00490044005d0020003d002000540065006d00 ,
                        0x700056006100720073004c006f006e006700280022006c006e00670043007500 ,
                        0x7200720065006e007400530065007300730069006f006e004900440022002900 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =100
                    Height =100
                    TabIndex =6
                    Name ="txtStepID"
                    ControlSource ="StepID"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedWidth =100
                    LayoutCachedHeight =100
                End
            End
        End
        Begin FormFooter
            Height =2700
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =90
                    Width =7200
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblTools"
                    Caption ="Tools for selected session"
                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =360
                End
                Begin Subform
                    OverlapFlags =85
                    Left =360
                    Top =450
                    Width =7200
                    Height =2160
                    Name ="sfrmSessionTools"
                    SourceObject ="Form.sfrmSessionTool"
                    LinkChildFields ="SessionID"
                    LinkMasterFields ="SessionID"

                    LayoutCachedLeft =360
                    LayoutCachedTop =450
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =2610
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mControlProperties As FormControlHighlighting
Private Sub ApplyCF()
100   On Error GoTo errHandler
      Dim aCtl(4) As control
      Dim i       As Integer
      Dim oFC     As FormatCondition
110   Set aCtl(0) = Me.txtSessionDate
120   Set aCtl(1) = Me.txtSessionStartTime
130   Set aCtl(2) = Me.txtSessionEndTime
140   Set aCtl(3) = Me.txtRevisionCycles
150   Set aCtl(4) = Me.txtNotes
160   For i = 0 To 4
170       With aCtl(i).FormatConditions
180           .Delete
190           Set oFC = .Add(acExpression, , "[SessionID] = TempVarsLong(""lngCurrentSessionID"")")
200           oFC.BackColor = RGB(255, 255, 153)
210       End With
220   Next i
Cleanup:
230   On Error Resume Next
240   Set oFC = Nothing
250   Exit Sub
errHandler:
260   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
270   Resume Cleanup
280   Resume
End Sub
Private Sub UpdateNavButtons()
100   On Error GoTo errHandler
110   Me.NavigationButtons = (TempVarsLong("lngStepID") = 0)
Cleanup:
120   Exit Sub
errHandler:
130   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
140   Resume Cleanup
150   Resume
End Sub

Private Sub Form_Load()
100       On Error GoTo errHandler

      Dim lngSessionID As Long

110   Set mControlProperties = New FormControlHighlighting
120   mControlProperties.Initialize Me
      Dim lngStepID    As Long
      Dim lngProjectID As Long

130       lngSessionID = Nz(DMax("SessionID", "tblSession"), 0)
140       If lngSessionID > 0 Then
150           lngStepID = Nz(DLookup("StepID", "tblSession", "SessionID=" & lngSessionID), 0)
160           lngProjectID = Nz(DLookup("ProjectID", "tblStep", "StepID=" & lngStepID), 0)
170       End If

180       SetTempVar "lngProjectID", lngProjectID
190       SetTempVar "lngStepID", lngStepID

200       Me.cboProjectFilter = IIf(lngProjectID > 0, lngProjectID, 0)
210       Me.cboStepFilter.Requery
220       Me.cboStepFilter = IIf(lngStepID > 0, lngStepID, 0)
230       Me.Requery
240       Call UpdateNavButtons
250       Call ApplyCF

Cleanup:
260       Exit Sub

errHandler:
270       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
280       Resume Cleanup
290       Resume
End Sub
Private Sub Form_Current()
100   On Error GoTo errHandler
110   SetTempVar "lngCurrentSessionID", Nz(Me.SessionID, 0)
120   Me.Refresh  ' .Refresh (not .Repaint) required to trigger CF re-evaluation
Cleanup:
130   Exit Sub
errHandler:
140   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
150   Resume Cleanup
160   Resume
End Sub
Private Sub Form_BeforeInsert(Cancel As Integer)
100   On Error GoTo errHandler
      Dim lngStepID As Long
110   lngStepID = TempVarsLong("lngStepID")
120   If lngStepID > 0 Then Me.txtStepID = lngStepID
Cleanup:
130   Exit Sub
errHandler:
140   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
150   Resume Cleanup
160   Resume
End Sub

Private Sub cboProjectFilter_AfterUpdate()
100       On Error GoTo errHandler

110       SetTempVar "lngProjectID", Nz(Me.cboProjectFilter, 0)
120       Me.cboStepFilter.Requery
130       Me.cboStepFilter = Null
140       SetTempVar "lngStepID", 0
150       Me.Requery
160       Call UpdateNavButtons

Cleanup:
170       Exit Sub

errHandler:
180       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
190       Resume Cleanup
200       Resume
End Sub

Private Sub cboStepFilter_AfterUpdate()
100       On Error GoTo errHandler

110       SetTempVar "lngStepID", Nz(Me.cboStepFilter, 0)
120       Me.Requery
130       Call UpdateNavButtons

Cleanup:
140       Exit Sub

errHandler:
150       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
160       Resume Cleanup
170       Resume
End Sub

Private Sub cmdNewSession_Click()
100       On Error GoTo errHandler

110       If TempVarsLong("lngStepID") = 0 Then
120           MsgBox "Please select a Step before adding a new session.", vbExclamation, "DevPulse"
130           Me.cboStepFilter.SetFocus
140           GoTo Cleanup
150       End If
160       DoCmd.GoToRecord , , acNewRec

Cleanup:
170       Exit Sub

errHandler:
180       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
190       Resume Cleanup
200       Resume
End Sub



Private Sub cboProjectFilter_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("cboProjectFilter")
End Sub

Private Sub cboProjectFilter_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("cboProjectFilter")
End Sub

Private Sub cboStepFilter_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("cboStepFilter")
End Sub

Private Sub cboStepFilter_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("cboStepFilter")
End Sub

Private Sub txtSessionDate_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtSessionDate")
End Sub

Private Sub txtSessionDate_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtSessionDate")
End Sub

Private Sub txtSessionStartTime_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtSessionStartTime")
End Sub

Private Sub txtSessionStartTime_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtSessionStartTime")
End Sub

Private Sub txtSessionEndTime_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtSessionEndTime")
End Sub

Private Sub txtSessionEndTime_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtSessionEndTime")
End Sub

Private Sub chkContextReset_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("chkContextReset")
End Sub

Private Sub chkContextReset_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("chkContextReset")
End Sub

Private Sub txtRevisionCycles_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtRevisionCycles")
End Sub

Private Sub txtRevisionCycles_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtRevisionCycles")
End Sub

Private Sub txtNotes_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtNotes")
End Sub

Private Sub txtNotes_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtNotes")
End Sub

Private Sub txtStepID_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtStepID")
End Sub

Private Sub txtStepID_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtStepID")
End Sub
