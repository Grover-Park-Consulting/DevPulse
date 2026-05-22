Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =18000
    DatasheetFontHeight =11
    ItemSuffix =12
    Top =484
    Right =10665
    Bottom =11318
    RecSrcDt = Begin
        0x50bff7df0889e640
    End
    RecordSource ="tblSession"
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
        Begin FormHeader
            Height =460
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
                    Top =60
                    Width =1620
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrDate"
                    Caption ="Date"
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    Left =2070
                    Top =60
                    Width =1260
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrStart"
                    Caption ="Start"
                    LayoutCachedLeft =2070
                    LayoutCachedTop =60
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    Left =3420
                    Top =60
                    Width =1260
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrEnd"
                    Caption ="End"
                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    Left =4770
                    Top =60
                    Width =900
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrReset"
                    Caption ="Reset"
                    LayoutCachedLeft =4770
                    LayoutCachedTop =60
                    LayoutCachedWidth =5670
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    Left =5760
                    Top =60
                    Width =1350
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrRevisions"
                    Caption ="Revisions"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =7110
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    Left =7200
                    Top =60
                    Width =10530
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrNotes"
                    Caption ="Notes"
                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =17730
                    LayoutCachedHeight =400
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
                    Left =360
                    Top =30
                    Width =1620
                    Height =450
                    FontSize =14
                    Name ="txtSessionDate"
                    ControlSource ="SessionDate"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =360
                    LayoutCachedTop =30
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2070
                    Top =30
                    Width =1260
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="txtSessionStartTime"
                    ControlSource ="SessionStartTime"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =2070
                    LayoutCachedTop =30
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3420
                    Top =30
                    Width =1260
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtSessionEndTime"
                    ControlSource ="SessionEndTime"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =3420
                    LayoutCachedTop =30
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4770
                    Top =30
                    Width =900
                    Height =450
                    TabIndex =3
                    Name ="chkContextReset"
                    ControlSource ="ContextReset"

                    LayoutCachedLeft =4770
                    LayoutCachedTop =30
                    LayoutCachedWidth =5670
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5760
                    Top =30
                    Width =1350
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtRevisionCycles"
                    ControlSource ="RevisionCycles"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5760
                    LayoutCachedTop =30
                    LayoutCachedWidth =7110
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7200
                    Top =30
                    Width =10530
                    Height =450
                    FontSize =14
                    TabIndex =5
                    Name ="txtNotes"
                    ControlSource ="Notes"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =7200
                    LayoutCachedTop =30
                    LayoutCachedWidth =17730
                    LayoutCachedHeight =480
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
