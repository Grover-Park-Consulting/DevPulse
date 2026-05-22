Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
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
    ItemSuffix =17
    Top =484
    Right =10665
    Bottom =11318
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0xb0a8fe450789e640
    End
    RecordSource ="qryProject"
    Caption ="Projects"
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
            Height =630
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
                    Left =360
                    Top =90
                    Width =14580
                    Height =450
                    FontSize =14
                    Name ="cboProjectSearch"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT TOP 1 0,'(All Projects)' FROM USysAppString UNION ALL SELECT ProjectID, P"
                        "rojectName FROM tblProject ORDER BY 2"
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =15120
                    Top =90
                    Width =3240
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="cmdNewProject"
                    Caption ="New Project"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =15120
                    LayoutCachedTop =90
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =540
                End
            End
        End
        Begin Section
            Height =9023
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
                    Width =1800
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblProjectName"
                    Caption ="Project Name"
                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2250
                    Top =90
                    Width =16110
                    Height =450
                    FontSize =14
                    Name ="txtProjectName"
                    ControlSource ="ProjectName"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =2250
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
                    Name ="lblProjectAreaID"
                    Caption ="Area"
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
                    Width =5400
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="cboProjectAreaID"
                    ControlSource ="ProjectAreaID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ProjectAreaID, ProjectAreaName FROM lcl_tlkpProjectArea WHERE InactiveDat"
                        "e IS NULL ORDER BY SortOrder;"
                    ColumnWidths ="0;3060"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =1350
                    LayoutCachedTop =630
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =1080
                End
                Begin Label
                    OverlapFlags =85
                    Left =6840
                    Top =630
                    Width =1125
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStatusID"
                    Caption ="Status"
                    LayoutCachedLeft =6840
                    LayoutCachedTop =630
                    LayoutCachedWidth =7965
                    LayoutCachedHeight =1080
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8055
                    Top =630
                    Width =10305
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="cboStatusID"
                    ControlSource ="StatusID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT StatusID, StatusName FROM lcl_tlkpStatus WHERE InactiveDate IS NULL ORDER"
                        " BY SortOrder;"
                    ColumnWidths ="0;4590"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =8055
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1688
                    Top =1170
                    Width =1500
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="txtStartDate"
                    ControlSource ="StartDate"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =1688
                    LayoutCachedTop =1170
                    LayoutCachedWidth =3188
                    LayoutCachedHeight =1620
                End
                Begin Label
                    OverlapFlags =85
                    Left =4478
                    Top =1170
                    Width =1688
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblCompletionDate"
                    Caption ="Completed"
                    LayoutCachedLeft =4478
                    LayoutCachedTop =1170
                    LayoutCachedWidth =6166
                    LayoutCachedHeight =1620
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6256
                    Top =1170
                    Width =1500
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtCompletionDate"
                    ControlSource ="CompletionDate"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =6256
                    LayoutCachedTop =1170
                    LayoutCachedWidth =7756
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
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =3285
                    Width =18000
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblNotes"
                    Caption ="Notes"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3285
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =3555
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =2070
                    Width =18000
                    Height =1125
                    FontSize =14
                    TabIndex =5
                    Name ="txtDescription"
                    ControlSource ="ProjectDescription"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =2070
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =3195
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =3645
                    Width =18000
                    Height =788
                    FontSize =14
                    TabIndex =6
                    Name ="txtNotes"
                    ControlSource ="Notes"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =360
                    LayoutCachedTop =3645
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =4433
                End
                Begin Subform
                    OverlapFlags =85
                    Left =360
                    Top =4523
                    Width =18360
                    Height =4500
                    TabIndex =7
                    Name ="sfrmStep"
                    SourceObject ="Form.sfrmStep"
                    LinkChildFields ="ProjectID"
                    LinkMasterFields ="ProjectID"

                    LayoutCachedLeft =360
                    LayoutCachedTop =4523
                    LayoutCachedWidth =18720
                    LayoutCachedHeight =9023
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
Private mControlProperties As FormControlHighlighting

Private Sub Form_Load()
100       On Error GoTo errHandler

      Dim lngID As Long

110   Set mControlProperties = New FormControlHighlighting
120   mControlProperties.Initialize Me

130       lngID = Nz(DMax("ProjectID", "tblProject"), 0)
140       SetTempVar "lngProjectID", lngID
150       Me.Requery
160       Me.cboProjectSearch = lngID
170       Call UpdateNavButtons

Cleanup:
180       Exit Sub

errHandler:
190       Call GlblErrMsg(iLn:=Erl, _
              sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
              sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
              bLog:=True)
200       Resume Cleanup
210       Resume
End Sub
Private Sub cboProjectSearch_AfterUpdate()
100       On Error GoTo errHandler

110       SetTempVar "lngProjectID", Nz(Me.cboProjectSearch, 0)
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

Private Sub cmdNewProject_Click()
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
Private Sub UpdateNavButtons()
100   On Error GoTo errHandler
      Dim bShowAll As Boolean
110   bShowAll = (TempVarsLong("lngProjectID") = 0)
120   Me.NavigationButtons = bShowAll
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

Private Sub cboProjectSearch_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("cboProjectSearch")
End Sub

Private Sub cboProjectSearch_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("cboProjectSearch")
End Sub

Private Sub txtProjectName_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtProjectName")
End Sub

Private Sub txtProjectName_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtProjectName")
End Sub

Private Sub cboProjectAreaID_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("cboProjectAreaID")
End Sub

Private Sub cboProjectAreaID_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("cboProjectAreaID")
End Sub

Private Sub cboStatusID_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("cboStatusID")
End Sub

Private Sub cboStatusID_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("cboStatusID")
End Sub

Private Sub txtStartDate_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtStartDate")
End Sub

Private Sub txtStartDate_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtStartDate")
End Sub

Private Sub txtCompletionDate_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtCompletionDate")
End Sub

Private Sub txtCompletionDate_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtCompletionDate")
End Sub

Private Sub txtDescription_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtDescription")
End Sub

Private Sub txtDescription_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtDescription")
End Sub

Private Sub txtNotes_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtNotes")
End Sub

Private Sub txtNotes_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtNotes")
End Sub

Private Sub Form_AfterInsert()
100       On Error GoTo errHandler

110       Me.cboProjectSearch.Requery

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
