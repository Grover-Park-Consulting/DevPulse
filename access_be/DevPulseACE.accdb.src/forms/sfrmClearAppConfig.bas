Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =11
    ItemSuffix =3
    Top =450
    Right =10676
    Bottom =10069
    RecSrcDt = Begin
        0xed58f5c84c8ae640
    End
    RecordSource ="tblAppConfig"
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
        Begin Section
            Height =420
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =30
                    Width =1680
                    Height =360
                    FontSize =14
                    Name ="txtConfigName"
                    ControlSource ="ConfigName"

                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1800
                    Top =30
                    Width =1500
                    Height =360
                    FontSize =14
                    TabIndex =1
                    Name ="txtConfigValue"
                    ControlSource ="ConfigValue"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =30
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Top =30
                    Width =3000
                    Height =360
                    FontSize =14
                    TabIndex =2
                    Name ="txtConfigDescription"
                    ControlSource ="ConfigDescription"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =30
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =390
                End
            End
        End
    End
End
