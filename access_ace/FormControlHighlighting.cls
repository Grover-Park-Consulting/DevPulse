VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormControlHighlighting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




Option Compare Database

Option Explicit



' ============================================================

' FormControlHighlighting  Per-form focus highlighting

'

' Setup (do once per form):

'

'   Step 1   Declare in the form module:

'       Private mControlProperties As FormControlHighlighting

'

'   Step 2  In Form_Load (or Form_Open):

'       Set mControlProperties = New FormControlHighlighting

'       mControlProperties.Initialize Me

'       ' Optional -- override the default highlight colors:

'       ' mControlProperties.HighlightBackColor = RGB(200, 230, 255)

'       ' mControlProperties.HighlightForeColor = vbBlack

'

'   Step 3  Run once from the Immediate Window to inject

'       GotFocus/LostFocus stubs into every eligible control:

'       codearchive.ForceInjectFocusHighlighting "frmYourFormName"

'       (Only needed once; re-run only if controls are added.)

'

' Continuous forms (DefaultView=1) and Datasheet forms (DefaultView=2)

' are silently skipped -- Initialize exits without enabling the class.

' ============================================================



Private mForm             As Form

Private mBackColors       As Collection

Private mForeColors       As Collection

Private mHighlightBackColor As Long

Private mHighlightForeColor As Long

Private mEnabled          As Boolean



Private Sub Class_Initialize()

    mHighlightBackColor = RGB(255, 255, 153)

    mHighlightForeColor = vbBlack

    mEnabled = False

End Sub



Public Sub Initialize(ByVal frm As Form)

    Dim ctl As Control



    ' Continuous and Datasheet forms: skip silently.

    ' Continuous: all row instances share one control object, so GotFocus

    ' highlighting would light every row. Datasheet has its own row selection.

    If frm.DefaultView = 1 Or frm.DefaultView = 2 Then Exit Sub



    Set mForm = frm

    Set mBackColors = New Collection

    Set mForeColors = New Collection



    For Each ctl In mForm.Controls

        Select Case ctl.ControlType

            Case acTextBox, acComboBox, acListBox, acCheckBox, acOptionButton

                On Error Resume Next

                mBackColors.Add ctl.BackColor, ctl.Name

                mForeColors.Add ctl.ForeColor, ctl.Name

                On Error GoTo 0

        End Select

    Next ctl



    mEnabled = True

End Sub



Public Property Let HighlightBackColor(ByVal Value As Long)

    mHighlightBackColor = Value

End Property



Public Property Get HighlightBackColor() As Long

    HighlightBackColor = mHighlightBackColor

End Property



Public Property Let HighlightForeColor(ByVal Value As Long)

    mHighlightForeColor = Value

End Property



Public Property Get HighlightForeColor() As Long

    HighlightForeColor = mHighlightForeColor

End Property



Public Sub ControlGotFocus(ByRef ctl As Control)

    If Not mEnabled Then Exit Sub

    On Error Resume Next

    ctl.BackColor = mHighlightBackColor

    ctl.ForeColor = mHighlightForeColor

End Sub



Public Sub ControlLostFocus(ByRef ctl As Control)

    If Not mEnabled Then Exit Sub

    On Error Resume Next

    ctl.BackColor = mBackColors(ctl.Name)

    ctl.ForeColor = mForeColors(ctl.Name)

End Sub



' ---- Control property utilities ----



Public Sub SetCaption(ByVal controlName As String, ByVal Caption As String)

    Dim ctrl As Control

    Set ctrl = mForm.Controls(controlName)

    If Not ctrl Is Nothing Then ctrl.Caption = Caption

End Sub



Public Sub SetEnabled(ByVal controlName As String, ByVal Enabled As Boolean)

    Dim ctrl As Control

    Set ctrl = mForm.Controls(controlName)

    If Not ctrl Is Nothing Then ctrl.Enabled = Enabled

End Sub



Public Sub SetVisible(ByVal controlName As String, ByVal Visible As Boolean)

    Dim ctrl As Control

    Set ctrl = mForm.Controls(controlName)

    If Not ctrl Is Nothing Then ctrl.Visible = Visible

End Sub



Public Sub SetValue(ByVal controlName As String, ByVal Value As Variant)

    Dim ctrl As Control

    Set ctrl = mForm.Controls(controlName)

    If Not ctrl Is Nothing Then ctrl.Value = Value

End Sub



Public Sub SetBackColor(ByVal controlName As String, ByVal Value As Long)

    Dim ctrl As Control

    Set ctrl = mForm.Controls(controlName)

    If Not ctrl Is Nothing Then ctrl.BackColor = Value

End Sub



Public Sub SetForeColor(ByVal controlName As String, ByVal Value As Long)

    Dim ctrl As Control

    Set ctrl = mForm.Controls(controlName)

    If Not ctrl Is Nothing Then ctrl.ForeColor = Value

End Sub



Private Sub Class_Terminate()

    Set mForm = Nothing

    Set mBackColors = Nothing

    Set mForeColors = Nothing

End Sub
