VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form DCalendar 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4350
   ClientLeft      =   3045
   ClientTop       =   2055
   ClientWidth     =   5085
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   171245569
      CurrentDate     =   37995
   End
End
Attribute VB_Name = "DCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SRCFILE = "DCalendar::"
Private mValue As Date
Private mBuddy As Control
Attribute mBuddy.VB_VarHelpID = -1

Public Event Changed()

Private Sub Form_Load()
    mvDate.Move 0, 0
    Me.Width = mvDate.Width + Me.Width - Me.ScaleWidth
    Me.Height = mvDate.Height + Me.Height - Me.ScaleHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyReturn
            mValue = mvDate.value
            Call UpdateBuddy
            RaiseEvent Changed
            Unload Me
    End Select
End Sub

Private Sub Form_Resize()
    Dim r As Long
    Dim c As Long
    
    Me.Height = Max(2310, Min(Me.Height, 8745))
    Me.Width = Max(2670, Min(Me.Width, 8190))
    
    Select Case Me.Height
        Case Is > 6960: r = 4
        Case Is > 4620: r = 3
        Case Is > 2310: r = 2
        Case Else:      r = 1
    End Select
    
    Select Case Me.Width
        Case Is > 5430: c = 3
        Case Is > 2670: c = 2
        Case Else:      c = 1
    End Select

    mvDate.MonthColumns = 1 'c
    mvDate.MonthRows = 1 'r
    
    
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    mValue = mvDate.value
    Call UpdateBuddy
    RaiseEvent Changed
    Unload Me
End Sub

Private Sub mvDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.Width = mvDate.Width + Me.Width - Me.ScaleWidth
    Me.Height = mvDate.Height + Me.Height - Me.ScaleHeight
End Sub

Private Sub mvDate_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)
    mValue = mvDate.value
    Call UpdateBuddy
    RaiseEvent Changed
End Sub

Private Sub vbalDropDownClient1_DeactivateForm()
    Unload Me
End Sub

Public Property Get value() As Date
    value = mValue
End Property
Public Property Let value(RHS As Date)
    mValue = RHS
    mvDate.value = mValue
End Property
Private Sub UpdateBuddy()
On Error Resume Next
    mBuddy = format(mValue, "medium date")
End Sub



Public Sub Popup(Buddy As Control, Optional Top As Single, Optional Left As Single)
On Error GoTo eh
    
    Dim tR   As RECT
    
    Set mBuddy = Buddy
    Call GetWindowRect(mBuddy.hwnd, tR)
    
    
    If Top = 0 And Left = 0 Then
        'not given so assume bottom left corner
        Me.Move tR.Left * Screen.TwipsPerPixelX, tR.Bottom * Screen.TwipsPerPixelY + Screen.TwipsPerPixelY
    Else
        'top left corner + top/left parameters
        Me.Move tR.Left * Screen.TwipsPerPixelX + Left, tR.Top * Screen.TwipsPerPixelY + Top
    End If
    
    
    If IsDate(mBuddy.Text) Then
        Me.value = DateValue(mBuddy.Text)
    Else
        Me.value = Now
    End If
        
    Me.Show vbModal
        
Exit Sub
eh: Call errHandler(SRCFILE & "Popup")
End Sub

Private Sub Shangle_DblClick()
    If Me.Height = 8745 And Me.Width = 8190 Then
        Me.Height = 2310
        Me.Width = 2670
    Else
        Me.Height = 8745
        Me.Width = 8190
    End If
    Call Form_Resize
End Sub

Private Sub Shangle_DoneSizing()
On Error Resume Next
    Me.Width = mvDate.Width
    Me.Height = mvDate.Height
End Sub
