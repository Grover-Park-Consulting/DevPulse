Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    RecordLocks =2
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5760
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =-9927
    Top =7808
    Right =749
    Bottom =17426
    RecSrcDt = Begin
        0x45d084e6748ae640
    End
    Caption ="DevPulse — Connection Setup"
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
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =360
                    Top =360
                    Width =1440
                    Height =450
                    FontSize =14
                    Name ="lblServerAddress"
                    Caption ="SQL Server address:"
                    LayoutCachedLeft =360
                    LayoutCachedTop =360
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =810
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =360
                    Width =3480
                    Height =450
                    FontSize =14
                    Name ="txtServerAddress"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =360
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =810
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =360
                    Top =1080
                    Width =1440
                    Height =450
                    FontSize =14
                    Name ="lblUID"
                    Caption ="User ID:"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =1530
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1080
                    Width =3480
                    Height =450
                    FontSize =14
                    TabIndex =1
                    Name ="txtUID"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1530
                    ThemeFontIndex =-1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =360
                    Top =1800
                    Width =1440
                    Height =450
                    FontSize =14
                    Name ="lblPWD"
                    Caption ="Password:"
                    LayoutCachedLeft =360
                    LayoutCachedTop =1800
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =2250
                    ThemeFontIndex =-1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =1800
                    Width =3480
                    Height =450
                    FontSize =14
                    TabIndex =2
                    Name ="txtPWD"
                    InputMask ="Password"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =1800
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =2250
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Top =2700
                    Width =1560
                    Height =450
                    FontSize =14
                    TabIndex =3
                    Name ="cmdConnect"
                    Caption ="Connect"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2700
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =3150
                    ThemeFontIndex =-1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3600
                    Top =2700
                    Height =450
                    FontSize =14
                    TabIndex =4
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =3600
                    LayoutCachedTop =2700
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =3150
                    ThemeFontIndex =-1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =360
                    Top =3360
                    Width =5040
                    Height =600
                    FontSize =14
                    ForeColor =255
                    Name ="lblStatus"
                    LayoutCachedLeft =360
                    LayoutCachedTop =3360
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =3960
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                End
            End
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


Private Sub cmdConnect_Click()
10    Dim strServer As String
20    Dim strUID As String
30    Dim strPWD As String

40    On Error GoTo errHandler

50    strServer = Nz(Me.txtServerAddress, "")
60    strUID = Nz(Me.txtUID, "")
70    strPWD = Nz(Me.txtPWD, "")

80    If strServer = "" Or strUID = "" Or strPWD = "" Then
90        Me.lblStatus.Caption = "Please enter server address, user ID, and password."
100       Me.lblStatus.Visible = True
110       Exit Sub
    End If

120   Me.lblStatus.Caption = "Connecting..."
130   Me.lblStatus.Visible = True
140   DoEvents

150   If InitConnect(strServer, "DevPulse", strUID, strPWD, True) Then
160       DoCmd.Close acForm, Me.Name
    Else
170       Me.lblStatus.Caption = "Connection failed. Check server address and credentials."
    End If

Cleanup:
180   Exit Sub
errHandler:
190   Call GlblErrMsg(iLn:=Erl, _
          sFrm:=Application.VBE.ActiveCodePane.CodeModule, _
          sCtl:=Application.VBE.ActiveCodePane.CodeModule.ProcOfLine(Erl, 0), _
          bLog:=True)
200   Resume Cleanup
210   Resume
End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub
