VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Appconstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False




Option Compare Database

Option Explicit



' ============================================================

' Appconstants — trimmed for DevPulse standalone distribution.

' Full version in GPC CodeArchive.

' Properties included: those required by GlblErrMsg only.

' All values read from USysAppString via DLookup.

' ============================================================



Public Property Get AppName() As String

Static strAppName As String

    strAppName = DLookup("AppStringValue", TempVars("AppString"), "AppStringName='AppName'")

    AppName = strAppName

    TempVars.Add Name:="AppName", Value:=strAppName

End Property



Public Property Get booDebug() As Boolean

Static booDBPrint As Boolean

    booDBPrint = DLookup("AppStringValue", AppString(), "AppStringName='booDebug'")

    booDebug = booDBPrint

    TempVars.Add Name:="booDebug", Value:=booDBPrint

End Property



Public Property Get MBErr() As String

Static strMBErr As String

    strMBErr = DLookup("AppStringValue", TempVars("AppString"), "AppStringName='AppName'") & _
               " " & DLookup("AppStringValue", TempVars("AppString"), "AppStringName='MBErr'")

    MBErr = strMBErr

    TempVars.Add Name:="MBErr", Value:=strMBErr

End Property



Public Property Get MBYNBtn() As String

Static strMBYNBtn As String

    strMBYNBtn = DLookup("AppStringValue", TempVars("AppString"), "AppStringName='MBYNBtn'")

    MBYNBtn = strMBYNBtn

    TempVars.Add Name:="MBYNBtn", Value:=strMBYNBtn

End Property



Public Property Get SupportEmail() As String

Static strSupportEmail As String

    strSupportEmail = DLookup("AppStringValue", TempVars("AppString"), "AppStringName='SupportEmail'")

    SupportEmail = strSupportEmail

    TempVars.Add Name:="SupportEmail", Value:=strSupportEmail

End Property



Public Property Get SupportPerson() As String

Static strSupportPerson As String

    strSupportPerson = DLookup("AppStringValue", TempVars("AppString"), "AppStringName='SupportPerson'")

    SupportPerson = strSupportPerson

    TempVars.Add Name:="SupportPerson", Value:=strSupportPerson

End Property
