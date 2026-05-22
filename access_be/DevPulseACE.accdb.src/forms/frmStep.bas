Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18720
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =5051
    Top =3938
    Right =24278
    Bottom =13095
    RecSrcDt = Begin
        0xeb673aea0889e640
    End
    RecordSource ="qryStep"
    Caption ="Steps"
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
            Height =1170
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1350
                    Top =630
                    Width =17010
                    Height =450
                    FontSize =14
                    Name ="cboStepSearch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TOP 1 0,'(All Steps)' FROM USysAppString UNION ALL SELECT StepID, StepNam"
                        "e FROM tblStep WHERE (TempVars!lngProjectID=0 OR ProjectID=TempVars!lngProjectID"
                        ") ORDER BY 2"
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1350
                    LayoutCachedTop =630
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15120
                    Top =90
                    Width =3240
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="cmdNewStep"
                    Caption ="New Step"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =15120
                    LayoutCachedTop =90
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =540
                End
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
                    Width =13680
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="cboProjectFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TOP 1 0,'(All Projects)' FROM USysAppString UNION ALL SELECT ProjectID, P"
                        "rojectName FROM tblProject ORDER BY 2"
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1350
                    LayoutCachedTop =90
                    LayoutCachedWidth =15030
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =630
                    Width =900
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStep"
                    Caption ="Step"
                    LayoutCachedLeft =360
                    LayoutCachedTop =630
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =1080
                End
            End
        End
        Begin Section
            Height =7380
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =90
                    Width =1350
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStepName"
                    Caption ="Step Name"
                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =1710
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =90
                    Width =16560
                    Height =450
                    FontSize =14
                    Name ="txtStepName"
                    ControlSource ="StepName"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =90
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =630
                    Width =900
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStepTypeID"
                    Caption ="Type"
                    LayoutCachedLeft =360
                    LayoutCachedTop =630
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1350
                    Top =630
                    Width =4500
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="cboStepTypeID"
                    ControlSource ="StepTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT StepTypeID, StepTypeName FROM lcl_tlkpStepType WHERE InactiveDate IS NULL"
                        " ORDER BY SortOrder;"
                    ColumnWidths ="0;4500"

                    LayoutCachedLeft =1350
                    LayoutCachedTop =630
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    Left =5940
                    Top =630
                    Width =900
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStatusID"
                    Caption ="Status"
                    LayoutCachedLeft =5940
                    LayoutCachedTop =630
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6930
                    Top =630
                    Width =3600
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="cboStatusID"
                    ControlSource ="StatusID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT StatusID, StatusName FROM lcl_tlkpStatus WHERE InactiveDate IS NULL ORDER"
                        " BY SortOrder;"
                    ColumnWidths ="0;3600"

                    LayoutCachedLeft =6930
                    LayoutCachedTop =630
                    LayoutCachedWidth =10530
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    Left =10620
                    Top =630
                    Width =1800
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblPredecessorStepID"
                    Caption ="Preceded by"
                    LayoutCachedLeft =10620
                    LayoutCachedTop =630
                    LayoutCachedWidth =12420
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =12510
                    Top =630
                    Width =5850
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="cboPredecessorStepID"
                    ControlSource ="PredecessorStepID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT StepID, StepName FROM tblStep WHERE (TempVars!lngProjectID=0 OR ProjectID"
                        "=TempVars!lngProjectID) ORDER BY StepName; "
                    ColumnWidths ="0;5850"

                    LayoutCachedLeft =12510
                    LayoutCachedTop =630
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =1170
                    Width =1238
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStartDate"
                    Caption ="Started"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1170
                    LayoutCachedWidth =1598
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1688
                    Top =1170
                    Width =2250
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtStartDate"
                    ControlSource ="StartDate"

                    LayoutCachedLeft =1688
                    LayoutCachedTop =1170
                    LayoutCachedWidth =3938
                    LayoutCachedHeight =1620
                End
                Begin Label
                    OverlapFlags =85
                    Left =4028
                    Top =1170
                    Width =1688
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblCompletionDate"
                    Caption ="Completed"
                    LayoutCachedLeft =4028
                    LayoutCachedTop =1170
                    LayoutCachedWidth =5716
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5806
                    Top =1170
                    Width =2250
                    Height =450
                    FontSize =14
                    TabIndex =5
                    Name ="txtCompletionDate"
                    ControlSource ="CompletionDate"

                    LayoutCachedLeft =5806
                    LayoutCachedTop =1170
                    LayoutCachedWidth =8056
                    LayoutCachedHeight =1620
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =1710
                    Width =18000
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblDescription"
                    Caption ="Description"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1710
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =1980
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =2070
                    Width =18000
                    Height =900
                    FontSize =14
                    TabIndex =6
                    Name ="txtDescription"
                    ControlSource ="StepDescription"

                    LayoutCachedLeft =360
                    LayoutCachedTop =2070
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =2970
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =3060
                    Width =18000
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblNotes"
                    Caption ="Notes"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3060
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =3330
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =3420
                    Width =18000
                    Height =630
                    FontSize =14
                    TabIndex =7
                    Name ="txtNotes"
                    ControlSource ="Notes"

                    LayoutCachedLeft =360
                    LayoutCachedTop =3420
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =4050
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =360
                    Top =4140
                    Width =3000
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =8
                    Name ="cmdSessions"
                    Caption ="Sessions"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =4140
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =4590
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =3360
                    Top =4140
                    Width =3000
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =9
                    Name ="cmdOutputs"
                    Caption ="Outputs"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =4140
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =4590
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =6360
                    Top =4140
                    Width =3000
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =10
                    Name ="cmdTools"
                    Caption ="Tools"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6360
                    LayoutCachedTop =4140
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =4590
                    Overlaps =1
                End
                Begin Subform
                    OverlapFlags =87
                    Left =360
                    Top =4590
                    Width =18360
                    Height =2700
                    TabIndex =11
                    Name ="sfrmActive"
                    SourceObject ="Form.sfrmSession"
                    LinkChildFields ="StepID"
                    LinkMasterFields ="StepID"

                    LayoutCachedLeft =360
                    LayoutCachedTop =4590
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =7290
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =100
                    Height =100
                    TabIndex =12
                    Name ="txtProjectID"
                    ControlSource ="ProjectID"

                    LayoutCachedWidth =100
                    LayoutCachedHeight =100
                End
            End
        End
        Begin FormFooter
            Height =360
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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

Private Const HIGHLIGHT_COLOR As Long = 5000786
Private Const NORMAL_COLOR As Long = -2147483633
Private Sub HighlightActiveButton(ByRef cmdActive As CommandButton)
    ' Adding a nav button requires adding a Set line below and incrementing the array bound.
Dim cmds(2) As CommandButton
Dim i As Integer
Dim bActive As Boolean
    Set cmds(0) = Me.cmdSessions
    Set cmds(1) = Me.cmdOutputs
    Set cmds(2) = Me.cmdTools
    For i = 0 To 2
        bActive = (cmds(i).Name = cmdActive.Name)
        cmds(i).BackColor = IIf(bActive, HIGHLIGHT_COLOR, NORMAL_COLOR)
        cmds(i).ForeColor = IIf(bActive, vbWhite, vbBlack)
    Next i
End Sub
Private Sub UpdateNavButtons()
Dim bShowAll As Boolean
    bShowAll = (TempVarsLong("lngStepID") = 0)
    Me.NavigationButtons = bShowAll
End Sub
Private Sub Form_Load()
100       On Error GoTo errHandler

      Dim lngID        As Long
      Dim lngProjectID As Long

110       lngProjectID = TempVarsLong("lngProjectID")
120       Me.cboProjectFilter = IIf(lngProjectID > 0, lngProjectID, 0)
130       If lngProjectID > 0 Then
140           lngID = Nz(DMax("StepID", "tblStep", "ProjectID=" & lngProjectID), 0)
150       Else
160           lngID = Nz(DMax("StepID", "tblStep"), 0)
170       End If
180       SetTempVar "lngStepID", lngID
190       Me.Requery
200       Me.cboStepSearch = lngID
210       Me.sfrmActive.SourceObject = "sfrmSession"
220       Call HighlightActiveButton(Me.cmdSessions)
230       Call UpdateNavButtons

Cleanup:
240       Exit Sub

errHandler:
250       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
260       Resume Cleanup
270       Resume
End Sub
Private Sub cboStepSearch_AfterUpdate()
100       On Error GoTo errHandler

110       SetTempVar "lngStepID", Nz(Me.cboStepSearch, 0)
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

Private Sub cmdNewStep_Click()
100       On Error GoTo errHandler

110       DoCmd.GoToRecord , , acNewRec

Cleanup:
120       Exit Sub

errHandler:
130       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
140       Resume Cleanup
150       Resume
End Sub

Private Sub cmdSessions_Click()
    Me.sfrmActive.SourceObject = "sfrmSession"
    Call HighlightActiveButton(Me.cmdSessions)
End Sub

Private Sub cmdOutputs_Click()
    Me.sfrmActive.SourceObject = "sfrmOutput"
    Call HighlightActiveButton(Me.cmdOutputs)
End Sub

Private Sub cmdTools_Click()
    Me.sfrmActive.SourceObject = "sfrmTool"
    Call HighlightActiveButton(Me.cmdTools)
End Sub

Private Sub cboProjectFilter_AfterUpdate()
100       On Error GoTo errHandler

110       SetTempVar "lngProjectID", Nz(Me.cboProjectFilter, 0)
120       Me.cboStepSearch.Requery
130       SetTempVar "lngStepID", 0
140       Me.cboStepSearch = 0
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

Private Sub Form_BeforeInsert(Cancel As Integer)
Dim lngProjectID As Long
    lngProjectID = TempVarsLong("lngProjectID")
    If lngProjectID > 0 Then Me.txtProjectID = lngProjectID
End Sub
