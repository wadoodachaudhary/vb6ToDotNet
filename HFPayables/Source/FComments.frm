VERSION 5.00
Begin VB.Form FComments 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Comments"
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


Public Function Edit(s As String, BuddyCtrl As Control, Optional Editable As Boolean = True, Optional Title As String = "Comments", Optional MaxLength As Long, Optional Tip As String) As Boolean
    On Error Resume Next
    Dim p As POINTAPI
    
    Load Me
    
    If BuddyCtrl Is Nothing Then
        Call IniGetForm(Me)
    Else
        p.X = BuddyCtrl.colPos(BuddyCtrl.Col) / Screen.TwipsPerPixelX
        p.Y = (BuddyCtrl.RowPos(BuddyCtrl.Row) + BuddyCtrl.RowHeight(BuddyCtrl.Row)) / Screen.TwipsPerPixelY
        If err.Description <> "Object doesn't support this property or method" Then
            Call ClientToScreen(BuddyCtrl.hwnd, p)
            p.X = p.X * Screen.TwipsPerPixelX
            p.Y = p.Y * Screen.TwipsPerPixelY
            'adjust position if form runs off screen
            If p.X + Me.Width > Screen.Width Then p.X = Screen.Width - Me.Width
            If p.Y + Me.Height > Screen.Height Then p.Y = Screen.Height - Me.Height
            Me.Move p.X, p.Y
        End If
    End If
    
    mString = Replace(Replace(Replace(s, vbLf, vbCr), vbCr & vbCr, vbCr), vbCr, vbCrLf)
    Text1.Locked = Not Editable
    
    If Tip <> "" Then Text1.ToolTipText = Tip
    
    If Not Editable Then Text1.ToolTipText = ""
    Text1.MaxLength = MaxLength
    Text1 = mString
    mCancel = False
    Me.Caption = Title
    Me.Show vbModal
    Edit = Not mCancel
    s = mString
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case KeyCode = vbKeyEscape:                        mCancel = True:  Me.Hide
        Case KeyCode = vbKeyReturn And Shift = vbCtrlMask: mCancel = False: Me.Hide
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormCode Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call IniPutForm(Me)
End Sub

Private Sub text1_Change()
    mString = Text1
End Sub
