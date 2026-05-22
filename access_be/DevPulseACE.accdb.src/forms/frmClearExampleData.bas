Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7680
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =-14472
    Top =4140
    Right =-3796
    Bottom =14996
    RecSrcDt = Begin
        0x7e1821cc4c8ae640
    End
    Caption ="Clear Example Data"
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =540
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =60
                    Width =7200
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="lblTitle"
                    Caption ="CLEAR EXAMPLE DATA"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =540
                End
            End
        End
        Begin Section
            Height =3240
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1680
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="lblColName"
                    Caption ="Setting"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =420
                End
                Begin Label
                    OverlapFlags =85
                    Left =1800
                    Top =60
                    Width =1500
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="lblColValue"
                    Caption ="Value"
                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =420
                End
                Begin Label
                    OverlapFlags =85
                    Left =3360
                    Top =60
                    Width =3000
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="lblColDesc"
                    Caption ="Description"
                    LayoutCachedLeft =3360
                    LayoutCachedTop =60
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =420
                End
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =480
                    Width =7560
                    Height =1380
                    Name ="sfrmConfig"
                    SourceObject ="Form.sfrmClearAppConfig"

                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =1860
                End
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =1980
                    Width =7200
                    Height =480
                    FontSize =14
                    FontWeight =700
                    ForeColor =255
                    Name ="lblWarning"
                    Caption ="This cannot be undone.  Back up first (Admin › Backup)."
                    LayoutCachedLeft =60
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =2460
                    ForeThemeColorIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =2580
                    Height =480
                    TabIndex =1
                    Name ="btnClearData"
                    Caption ="Clear Data"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =2580
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =3060
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1620
                    Top =2580
                    Height =480
                    TabIndex =2
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =3060
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


Private Sub btnClearData_Click()
100   On Error GoTo errHandler

110   If MsgBox("Delete ALL example data? This cannot be undone.", _
              vbExclamation + vbYesNo, "Clear Example Data") <> vbYes Then
120       Exit Sub
      End If

130   If Me.sfrmConfig.Form.Dirty Then Me.sfrmConfig.Form.Dirty = False

140   CurrentDb.Execute "DELETE FROM tblOutput", dbFailOnError
150   CurrentDb.Execute "DELETE FROM tblSessionTool", dbFailOnError
160   CurrentDb.Execute "DELETE FROM tblSession", dbFailOnError
170   CurrentDb.Execute "DELETE FROM tblStep", dbFailOnError
180   CurrentDb.Execute "DELETE FROM tblProject", dbFailOnError

190   CurrentDb.Execute "DELETE FROM tlkpTool WHERE InactiveDate IS NOT NULL", dbFailOnError
200   CurrentDb.Execute "DELETE FROM tlkpStatus WHERE InactiveDate IS NOT NULL", dbFailOnError
210   CurrentDb.Execute "DELETE FROM tlkpStepType WHERE InactiveDate IS NOT NULL", dbFailOnError
220   CurrentDb.Execute "DELETE FROM tlkpProjectArea WHERE InactiveDate IS NOT NULL", dbFailOnError
230   CurrentDb.Execute "DELETE FROM tlkpOutputType WHERE InactiveDate IS NOT NULL", dbFailOnError

240   MsgBox "Example data cleared. You are ready to start.", vbInformation, "Done"
250   DoCmd.Close acForm, Me.Name
260   If CurrentProject.AllForms("frmMain").IsLoaded Then Forms!frmMain.Requery

Cleanup:
270   Exit Sub

errHandler:
280   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
290   Resume Cleanup
300   Resume
End Sub

' ---------------------------------------------------------------------------

Private Sub btnCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub
