Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    Top =450
    RecSrcDt = Begin
        0x57ad73d6c989e640
    End
    RecordSource ="qryRptProjectSummary"
    Caption ="GPC Data AI Assisted Development Projects"
    FilterOnLoad =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Calibri"
        End
        Begin TextBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Calibri"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="StatusGroupLabel"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            KeepTogether =2
            ControlSource ="ProjectID"
        End
        Begin PageHeader
            Height =810
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =180
                    Top =90
                    Width =10620
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="lblTitle"
                    Caption ="GPC Data AI Assisted Development Projects"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10800
                    Top =225
                    Width =3420
                    Height =450
                    FontSize =14
                    Name ="txtRptDate"
                    ControlSource ="=Format(Now(),\"mmmm d, yyyy\")"

                End
            End
        End
        Begin BreakHeader
            Height =540
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =180
                    Top =90
                    Width =14040
                    Height =360
                    FontSize =16
                    FontWeight =700
                    Name ="txtStatusGroupLabel"
                    ControlSource ="StatusGroupLabel"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4770
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin Label
                    Left =180
                    Top =90
                    Width =1080
                    Height =450
                    FontSize =14
                    Name ="lblProjName"
                    Caption ="Project"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1350
                    Top =90
                    Width =12870
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="txtProjectName"
                    ControlSource ="ProjectName"

                End
                Begin Label
                    Left =180
                    Top =630
                    Width =720
                    Height =450
                    FontSize =14
                    Name ="lblArea"
                    Caption ="Area"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =990
                    Top =630
                    Width =3870
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="txtProjectAreaName"
                    ControlSource ="ProjectAreaName"

                End
                Begin Label
                    Left =5040
                    Top =630
                    Width =990
                    Height =450
                    FontSize =14
                    Name ="lblStatus"
                    Caption ="Status"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6120
                    Top =630
                    Width =8100
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="txtProjectStatus"
                    ControlSource ="ProjectStatus"

                End
                Begin Label
                    Left =180
                    Top =1170
                    Width =990
                    Height =450
                    FontSize =14
                    Name ="lblStarted"
                    Caption ="Started"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1260
                    Top =1170
                    Width =2700
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    Name ="txtProjectStartDate"
                    ControlSource ="ProjectStartDate"

                End
                Begin Label
                    Left =4140
                    Top =1170
                    Width =1440
                    Height =450
                    FontSize =14
                    Name ="lblCompleted"
                    Caption ="Completed"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5670
                    Top =1170
                    Width =2700
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =5
                    Name ="txtProjectCompletionDate"
                    ControlSource ="ProjectCompletionDate"

                End
                Begin Label
                    Left =180
                    Top =1800
                    Width =14040
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="lblDescription"
                    Caption ="Description"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =180
                    Top =2250
                    Width =14040
                    Height =720
                    FontSize =14
                    TabIndex =6
                    Name ="txtProjectDescription"
                    ControlSource ="ProjectDescription"

                End
                Begin Label
                    Left =180
                    Top =3060
                    Width =14040
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="lblNotes"
                    Caption ="Notes"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =180
                    Top =3510
                    Width =14040
                    Height =720
                    FontSize =14
                    Name ="txtProjectNotes"
                    ControlSource ="ProjectNotes"

                End
                Begin Label
                    Left =180
                    Top =4320
                    Width =5940
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblColStep"
                    Caption ="Step Name"
                End
                Begin Label
                    Left =6210
                    Top =4320
                    Width =2070
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblColType"
                    Caption ="Type"
                End
                Begin Label
                    Left =8370
                    Top =4320
                    Width =2070
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblColStatus"
                    Caption ="Status"
                End
                Begin Label
                    TextAlign =3
                    Left =10530
                    Top =4320
                    Width =1620
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblColStarted"
                    Caption ="Started"
                End
                Begin Label
                    TextAlign =3
                    Left =12240
                    Top =4320
                    Width =1620
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblColCompleted"
                    Caption ="Completed"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =480
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =180
                    Width =5940
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="txtStepName"
                    ControlSource ="StepName"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =6210
                    Width =2070
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtStepTypeName"
                    ControlSource ="StepTypeName"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =8370
                    Width =2070
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="txtStepStatus"
                    ControlSource ="StepStatus"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =10530
                    Width =1620
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="txtStepStartDate"
                    ControlSource ="StepStartDate"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Left =12240
                    Width =1620
                    Height =450
                    FontSize =14
                    Name ="txtStepCompletionDate"
                    ControlSource ="StepCompletionDate"

                End
            End
        End
        Begin BreakFooter
            Height =630
            BreakLevel =1
            Name ="GroupFooter0"
            Begin
                Begin Label
                    Left =180
                    Top =90
                    Width =1350
                    Height =450
                    FontSize =14
                    Name ="lblFtrSteps"
                    Caption ="Steps:"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1620
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="txtFtrSteps"
                    ControlSource ="=Count([StepID])"

                End
                Begin Label
                    Left =2520
                    Top =90
                    Width =1620
                    Height =450
                    FontSize =14
                    Name ="lblFtrSessions"
                    Caption ="Sessions:"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4230
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="txtFtrSessions"
                    ControlSource ="=Sum([SessionCount])"

                End
                Begin Label
                    Left =5130
                    Top =90
                    Width =2520
                    Height =450
                    FontSize =14
                    Name ="lblFtrRevisions"
                    Caption ="Revision Cycles:"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7740
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="txtFtrRevisions"
                    ControlSource ="=Sum([RevisionTotal])"

                End
                Begin Label
                    Left =8640
                    Top =90
                    Width =2520
                    Height =450
                    FontSize =14
                    Name ="lblFtrResets"
                    Caption ="Context Resets:"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =11250
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="txtFtrResets"
                    ControlSource ="=Sum([ResetCount])"

                End
            End
        End
        Begin PageFooter
            Height =1080
            Name ="PageFooterSection"
            Begin
                Begin Label
                    Left =180
                    Top =90
                    Width =1620
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblGrandSteps"
                    Caption ="Total Steps:"
                    LayoutCachedLeft =180
                    LayoutCachedTop =90
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1890
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="txtGrandSteps"
                    ControlSource ="=DCount(\"*\",\"tblStep\")"

                    LayoutCachedLeft =1890
                    LayoutCachedTop =90
                    LayoutCachedWidth =2610
                    LayoutCachedHeight =540
                End
                Begin Label
                    Left =2790
                    Top =90
                    Width =1260
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblGrandSessions"
                    Caption ="Sessions:"
                    LayoutCachedLeft =2790
                    LayoutCachedTop =90
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4140
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="txtGrandSessions"
                    ControlSource ="=DSum(\"SessionCount\",\"vw_RptProjectSummary\")"

                    LayoutCachedLeft =4140
                    LayoutCachedTop =90
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =540
                End
                Begin Label
                    Left =5040
                    Top =90
                    Width =2250
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblGrandRevisions"
                    Caption ="Revision Cycles:"
                    LayoutCachedLeft =5040
                    LayoutCachedTop =90
                    LayoutCachedWidth =7290
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7380
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="txtGrandRevisions"
                    ControlSource ="=DSum(\"RevisionTotal\",\"vw_RptProjectSummary\")"

                    LayoutCachedLeft =7380
                    LayoutCachedTop =90
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =540
                End
                Begin Label
                    Left =8280
                    Top =90
                    Width =2160
                    Height =450
                    FontSize =14
                    FontWeight =700
                    Name ="lblGrandResets"
                    Caption ="Context Resets:"
                    LayoutCachedLeft =8280
                    LayoutCachedTop =90
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10530
                    Top =90
                    Width =720
                    Height =450
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    Name ="txtGrandResets"
                    ControlSource ="=DSum(\"ResetCount\",\"vw_RptProjectSummary\")"

                    LayoutCachedLeft =10530
                    LayoutCachedTop =90
                    LayoutCachedWidth =11250
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11520
                    Top =630
                    Width =2700
                    Height =360
                    FontSize =12
                    Name ="txtPageNumber"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"

                End
            End
        End
    End
End
