Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18720
    DatasheetFontHeight =11
    ItemSuffix =10
    Top =484
    Right =10665
    Bottom =7594
    RecSrcDt = Begin
        0xd6b9f208b289e640
    End
    RecordSource ="tlkpOutputType"
    Caption ="Manage Output Types"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
        Begin FormHeader
            Height =630
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
                    Width =11940
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrName"
                    Caption ="Name"
                    LayoutCachedLeft =360
                    LayoutCachedTop =90
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =12390
                    Top =90
                    Width =1620
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrSortOrder"
                    Caption ="Sort Order"
                    LayoutCachedLeft =12390
                    LayoutCachedTop =90
                    LayoutCachedWidth =14010
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =14100
                    Top =90
                    Width =1620
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrInactiveDate"
                    Caption ="Inactive Date"
                    LayoutCachedLeft =14100
                    LayoutCachedTop =90
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =540
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
                    Top =15
                    Width =11940
                    Height =450
                    FontSize =14
                    Name ="txtOutputTypeName"
                    ControlSource ="OutputTypeName"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =360
                    LayoutCachedTop =15
                    LayoutCachedWidth =12300
                    LayoutCachedHeight =465
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12390
                    Top =15
                    Width =1620
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="txtSortOrder"
                    ControlSource ="SortOrder"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =12390
                    LayoutCachedTop =15
                    LayoutCachedWidth =14010
                    LayoutCachedHeight =465
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14100
                    Top =15
                    Width =1620
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtInactiveDate"
                    ControlSource ="InactiveDate"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =14100
                    LayoutCachedTop =15
                    LayoutCachedWidth =15720
                    LayoutCachedHeight =465
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
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

Private Sub Form_Close()
100   On Error GoTo errHandler
110   Call CloseLookupForm("lcl_tlkpOutputType")
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
