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
    ItemSuffix =10
    Top =484
    Right =10665
    Bottom =11318
    RecSrcDt = Begin
        0x8ef7f8410789e640
    End
    RecordSource ="tblStep"
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
                    TextAlign =1
                    Left =360
                    Top =60
                    Width =7560
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrStepName"
                    Caption ="Step Name"
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8010
                    Top =60
                    Width =3150
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrStepType"
                    Caption ="Type"
                    LayoutCachedLeft =8010
                    LayoutCachedTop =60
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =11250
                    Top =60
                    Width =2880
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrStatus"
                    Caption ="Status"
                    LayoutCachedLeft =11250
                    LayoutCachedTop =60
                    LayoutCachedWidth =14130
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =14220
                    Top =60
                    Width =1620
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrStartDate"
                    Caption ="Start"
                    LayoutCachedLeft =14220
                    LayoutCachedTop =60
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =400
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =16200
                    Top =60
                    Width =1530
                    Height =340
                    FontSize =14
                    FontWeight =700
                    Name ="lblHdrDone"
                    Caption ="Done"
                    LayoutCachedLeft =16200
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
                    Width =7560
                    Height =450
                    FontSize =14
                    Name ="txtStepName"
                    ControlSource ="StepName"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =360
                    LayoutCachedTop =30
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8010
                    Top =30
                    Width =3150
                    Height =450
                    FontSize =14
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Name ="cboStepTypeID"
                    ControlSource ="StepTypeID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT StepTypeID, StepTypeName FROM lcl_tlkpStepType WHERE InactiveDate IS NULL"
                        " ORDER BY SortOrder;"
                    ColumnWidths ="0;1800"

                    LayoutCachedLeft =8010
                    LayoutCachedTop =30
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11250
                    Top =30
                    Width =2880
                    Height =450
                    FontSize =14
                    TabIndex =2
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Name ="cboStatusID"
                    ControlSource ="StatusID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT StatusID, StatusName FROM lcl_tlkpStatus WHERE InactiveDate IS NULL ORDER"
                        " BY SortOrder;"
                    ColumnWidths ="0;1620"

                    LayoutCachedLeft =11250
                    LayoutCachedTop =30
                    LayoutCachedWidth =14130
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14220
                    Top =30
                    Width =1620
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="txtStartDate"
                    ControlSource ="StartDate"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =14220
                    LayoutCachedTop =30
                    LayoutCachedWidth =15840
                    LayoutCachedHeight =480
                    ConditionalFormat14 = Begin
                        0x01000100000002000000000000000100000000000000ffff9900000000000000 ,
                        0x00000000000000ffff99000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =16200
                    Top =30
                    Width =1530
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtCompletionDate"
                    ControlSource ="CompletionDate"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000200000000000000000000000200000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =16200
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
