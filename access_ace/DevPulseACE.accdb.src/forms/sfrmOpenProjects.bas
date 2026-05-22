Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12000
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =754
    Top =2835
    Right =12488
    Bottom =5265
    RecSrcDt = Begin
        0xc43b221bb689e640
    End
    RecordSource ="qryOpenProjects"
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
                    Top =90
                    Width =5850
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblProjectName"
                    Caption ="Project Name"
                    LayoutCachedTop =90
                    LayoutCachedWidth =5850
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =5940
                    Top =90
                    Width =2700
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStatusName"
                    Caption ="Status"
                    LayoutCachedLeft =5940
                    LayoutCachedTop =90
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =8730
                    Top =90
                    Width =1620
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStartDate"
                    Caption ="Started"
                    LayoutCachedLeft =8730
                    LayoutCachedTop =90
                    LayoutCachedWidth =10350
                    LayoutCachedHeight =540
                End
                Begin Label
                    OverlapFlags =85
                    Left =10440
                    Top =90
                    Width =840
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblStepCount"
                    Caption ="Steps"
                    LayoutCachedLeft =10440
                    LayoutCachedTop =90
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =540
                End
            End
        End
        Begin Section
            Height =480
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =5850
                    Height =450
                    FontSize =14
                    Name ="txtProjectName"
                    ControlSource ="ProjectName"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedWidth =5850
                    LayoutCachedHeight =450
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5940
                    Width =2700
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="txtStatusName"
                    ControlSource ="StatusName"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5940
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =450
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8730
                    Width =1620
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtStartDate"
                    ControlSource ="StartDate"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =8730
                    LayoutCachedWidth =10350
                    LayoutCachedHeight =450
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10440
                    Width =840
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="txtStepCount"
                    ControlSource ="=DCount(\"*\",\"tblStep\",\"ProjectID=\" & [ProjectID])"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =10440
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =450
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
