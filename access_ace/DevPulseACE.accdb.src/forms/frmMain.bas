Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
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
    ItemSuffix =18
    Left =8066
    Top =3060
    Right =26786
    Bottom =12870
    RecSrcDt = Begin
        0x0a421660b389e640
    End
    Caption ="DevPulse"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin ModernChart
            OldBorderStyle =0
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackThemeColorIndex =1
            ChartTitle ="Chart Title"
            ChartTitleFontName ="Calibri"
            ThemeFontIndex =1
            ChartSubtitle ="Subtitle"
        End
        Begin FormHeader
            Height =720
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
                    Top =135
                    Width =720
                    Height =450
                    FontSize =14
                    Name ="lblLogo"
                    Caption ="[Logo]"
                    LayoutCachedLeft =360
                    LayoutCachedTop =135
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =585
                End
                Begin Label
                    OverlapFlags =85
                    Left =1260
                    Top =135
                    Width =17100
                    Height =450
                    FontSize =18
                    FontWeight =700
                    Name ="lblTitle"
                    Caption ="DevPulse — AI-Assisted Development Tracker"
                    LayoutCachedLeft =1260
                    LayoutCachedTop =135
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =585
                End
            End
        End
        Begin Section
            Height =9000
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
                    Width =12000
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblOpenProjects"
                    Caption ="Open Projects"
                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =360
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =3330
                    Width =12000
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblCompletedProjects"
                    Caption ="Completed Projects"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3330
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =3600
                End
                Begin Label
                    OverlapFlags =85
                    Left =12720
                    Top =90
                    Width =5640
                    Height =270
                    FontSize =14
                    FontWeight =700
                    Name ="lblMetrics"
                    Caption ="What We’ve Done"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =90
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =360
                End
                Begin Subform
                    OverlapFlags =85
                    Left =360
                    Top =450
                    Width =12000
                    Height =2700
                    Name ="sfrmOpenProjects"
                    SourceObject ="Form.sfrmOpenProjects"

                    LayoutCachedLeft =360
                    LayoutCachedTop =450
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =3150
                End
                Begin Subform
                    OverlapFlags =85
                    Left =360
                    Top =3690
                    Width =12000
                    Height =2700
                    TabIndex =1
                    Name ="sfrmCompletedProjects"
                    SourceObject ="Form.sfrmCompletedProjects"

                    LayoutCachedLeft =360
                    LayoutCachedTop =3690
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =6390
                End
                Begin Label
                    OverlapFlags =85
                    Left =12720
                    Top =450
                    Width =3600
                    Height =450
                    FontSize =14
                    Name ="lblMProjects"
                    Caption ="Projects"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =450
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =900
                End
                Begin Label
                    OverlapFlags =85
                    Left =12720
                    Top =990
                    Width =3600
                    Height =450
                    FontSize =14
                    Name ="lblMSessions"
                    Caption ="Sessions"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =990
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =1980
                End
                Begin Label
                    OverlapFlags =85
                    Left =12720
                    Top =1530
                    Width =3600
                    Height =450
                    FontSize =14
                    Name ="lblMRevisions"
                    Caption ="Revision Cycles"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =1530
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OverlapFlags =85
                    Left =12720
                    Top =2070
                    Width =3600
                    Height =450
                    FontSize =14
                    Name ="lblMResets"
                    Caption ="Context Resets"
                    LayoutCachedLeft =12720
                    LayoutCachedTop =2070
                    LayoutCachedWidth =16320
                    LayoutCachedHeight =3060
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16410
                    Top =450
                    Width =1950
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtMProjects"
                    ControlSource ="=DCount(\"*\",\"tblProject\")"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =16410
                    LayoutCachedTop =450
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16410
                    Top =990
                    Width =1950
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="txtMSessions"
                    ControlSource ="=DCount(\"*\",\"tblSession\")"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =16410
                    LayoutCachedTop =990
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =1980
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16410
                    Top =1530
                    Width =1950
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtMRevisions"
                    ControlSource ="=DSum(\"RevisionCycles\",\"tblSession\")"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =16410
                    LayoutCachedTop =1530
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =2520
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =16410
                    Top =2070
                    Width =1950
                    Height =450
                    FontSize =14
                    TabIndex =5
                    Name ="txtMResets"
                    ControlSource ="=DCount(\"*\",\"tblSession\",\"ContextReset=True\")"
                    OnGotFocus ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"

                    LayoutCachedLeft =16410
                    LayoutCachedTop =2070
                    LayoutCachedWidth =18360
                    LayoutCachedHeight =3060
                End
                Begin ModernChart
                    Top =2760
                    Left =12720
                    Width =5400
                    Height =5940
                    Name ="chrtStepsByProject"
                    RowSource ="qryChartStepsByProject"
                    OverlapFlags =85
                    ChartAxis ="[ProjectName]"
                    ChartLegend ="[StepCategory]"
                    ChartValues ="[StepCount]"
                    ChartType =4
                    ChartTitle ="Progress by Project"
                    TabIndex =6
                    ChartLegendData = Begin
                        0x02000000960000000f00000043006f006d0070006c0065007400650064002000 ,
                        0x5300740065007000730000000000005b9bd5005b9bd500000000000000000000 ,
                        0x000000000000009f6f00000f00000043006f006d0070006c0065007400650064 ,
                        0x0020005300740065007000730001000000000000000001000000000000000000 ,
                        0x00000000000000000000000000000000000000000000a5a5a500ed7d31008200 ,
                        0x00000a0000004f00700065006e002000530074006500700073000000000000ed ,
                        0x7d3100ed7d3100000000000000000000000000000000009f6f00000a0000004f ,
                        0x00700065006e0020005300740065007000730001000000000000000001000000 ,
                        0x00000000000000000000000000000000000000000000000000000000a5a5a500 ,
                        0xed7d3100
                    End
                    PrimaryValuesAxisFormat ="0"
                    PrimaryValuesAxisFontColor =8355711
                    SecondaryValuesAxisFontColor =8355711
                    CategoryAxisFontColor =8355711
                    ChartTitleFontColor =8355711
                    ChartTitleFontSize =14
                    LegendTextFontColor =8355711
                    ChartSubtitleFontColor =8355711
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
100   On Error GoTo errHandler
110   Set mControlProperties = New FormControlHighlighting
120   mControlProperties.Initialize Me
      gblnMainFormOpen = True
      If Not gobjRibbon Is Nothing Then gobjRibbon.InvalidateControl "btnClearData"
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



Private Sub txtMProjects_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtMProjects")
End Sub

Private Sub txtMProjects_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtMProjects")
End Sub

Private Sub txtMSessions_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtMSessions")
End Sub

Private Sub txtMSessions_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtMSessions")
End Sub

Private Sub txtMRevisions_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtMRevisions")
End Sub

Private Sub txtMRevisions_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtMRevisions")
End Sub

Private Sub txtMResets_GotFocus()
    mControlProperties.ControlGotFocus Me.Controls("txtMResets")
End Sub

Private Sub txtMResets_LostFocus()
    mControlProperties.ControlLostFocus Me.Controls("txtMResets")
End Sub

Private Sub Form_Close()
    gblnMainFormOpen = False
    If Not gobjRibbon Is Nothing Then gobjRibbon.InvalidateControl "btnClearData"
End Sub
