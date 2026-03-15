VERSION 5.00
Begin VB.Form FComments 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2385
   ClientLeft      =   1725
   ClientTop       =   1740
   ClientWidth     =   4950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FComments.frx":0000
      ToolTipText     =   "(Ctrl + Enter) to save -- (Esc) to cancel"
      Top             =   240
      Width           =   2715
   End
End
Attribute VB_Name = "FComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCancel As Boolean
Private mString As String

' TAG: Displays a modal comment editor positioned near a buddy control, normalizes line endings, and returns the edited string and whether the user confirmed.
' CONVERT: Replace with a Blazor modal dialog component (e.g., HfDialog) with a textarea bound to the comment string.
' CONVERT: Use CSS/JS interop to position the dialog near the triggering element, or use a standard centered modal as in FPricingWorkSheetOpus.razor.
' CONVERT: Handle Editable flag by setting the textarea to readonly and MaxLength via the maxlength HTML attribute.
' CONVERT: Return the edited string and confirmation status via an EventCallback or async Task<(bool, string)> method pattern.
' CONVERT: Normalize line endings using C# string.Replace before populating the textarea.
Public Function Edit(s As String, BuddyCtrl As Control, Optional Editable As Boolean = True, Optional Title As String = "Comments", Optional MaxLength As Long) As Boolean
    On Error Resume Next
    Dim p As POINTAPI
    
    Load Me
    
    p.X = BuddyCtrl.colPos(BuddyCtrl.Col) / Screen.TwipsPerPixelX
    p.Y = (BuddyCtrl.RowPos(BuddyCtrl.Row) + BuddyCtrl.RowHeight(BuddyCtrl.Row)) / Screen.TwipsPerPixelY
    Call ClientToScreen(BuddyCtrl.hwnd, p)
    p.X = p.X * Screen.TwipsPerPixelX
    p.Y = p.Y * Screen.TwipsPerPixelY
    'adjust position if form runs off screen
    If p.X + Me.Width > Screen.Width Then p.X = Screen.Width - Me.Width
    If p.Y + Me.Height > Screen.Height Then p.Y = Screen.Height - Me.Height
    Me.Move p.X, p.Y
    
        
    mString = Replace(Replace(Replace(s, vbLf, vbCr), vbCr & vbCr, vbCr), vbCr, vbCrLf)
    Text1.Locked = Not Editable
    If Not Editable Then Text1.ToolTipText = ""
    Text1.MaxLength = MaxLength
    Text1 = mString
    mCancel = False
    Me.Caption = Title
    Me.Show vbModal
    Edit = Not mCancel
    s = mString
End Function

' TAG: Handles keyboard shortcuts: Escape cancels and hides the form, Ctrl+Enter confirms and hides the form.
' CONVERT: Implement as @onkeydown handler on the Blazor dialog or use JS interop for global key capture.
' CONVERT: On Escape, set a cancel flag and close the dialog; on Ctrl+Enter, set confirmed and close.
' CONVERT: Reference the pattern used in FAssemblyOpus.razor and FItemsOpus.razor dialog OK/Cancel buttons, adding keyboard shortcut support.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyEscape:                        mCancel = True:  Me.Hide
        Case KeyCode = vbKeyReturn And Shift = vbCtrlMask: mCancel = False: Me.Hide
    End Select
End Sub

' TAG: Handles form unload by either unloading (if triggered by code) or hiding (if triggered by user closing the window).
' CONVERT: In Blazor, dialog visibility is controlled by a bool flag; no direct equivalent of Unload vs Hide distinction is needed.
' CONVERT: Set the dialog's @bind-Visible to false when the user closes via the X button or cancel action.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormCode Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

' TAG: Resizes the TextBox to fill the entire form client area when the form is resized.
' CONVERT: Use CSS (width: 100%; height: 100%) on the textarea within the Blazor dialog to achieve responsive sizing.
' CONVERT: No explicit resize handler is needed; CSS flex or grid layout handles this automatically.
Private Sub Form_Resize()
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

' TAG: Updates the internal string variable whenever the text box content changes.
' CONVERT: Use Blazor two-way binding (@bind) on the textarea element to automatically sync the value to a C# string property.
' CONVERT: No separate change handler is needed; the bound property is always current.
Private Sub Text1_Change()
    mString = Text1
End Sub