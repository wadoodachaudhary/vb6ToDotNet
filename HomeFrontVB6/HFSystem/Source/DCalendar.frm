VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form DCalendar 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1830
   ClientLeft      =   3045
   ClientTop       =   2055
   ClientWidth     =   2355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin HFSystem.Shangle Shangle 
      Height          =   240
      Left            =   3300
      Top             =   1680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   840
      ScaleHeight     =   15
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   3960
      ScaleHeight     =   2055
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   480
      Width           =   15
   End
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
            mValue = mvDate.Value
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

    mvDate.MonthColumns = c
    mvDate.MonthRows = r
    
    Shangle.Move Me.ScaleWidth - Shangle.Width - Screen.TwipsPerPixelX, Me.ScaleHeight - Shangle.Height - Screen.TwipsPerPixelY
    Picture1.Move Me.ScaleWidth - Screen.TwipsPerPixelX, 0, Screen.TwipsPerPixelX, Me.ScaleHeight
    Picture2.Move 0, Me.ScaleHeight - Screen.TwipsPerPixelY, Me.ScaleWidth - Screen.TwipsPerPixelY
    
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    mValue = mvDate.Value
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
    mValue = mvDate.Value
    Call UpdateBuddy
    RaiseEvent Changed
End Sub

Private Sub vbalDropDownClient1_DeactivateForm()
    Unload Me
End Sub

Public Property Get Value() As Date
    Value = mValue
End Property
Public Property Let Value(RHS As Date)
    mValue = RHS
    mvDate.Value = mValue
End Property
Private Sub UpdateBuddy()
On Error Resume Next
    mBuddy = Format(mValue, "medium date")
End Sub



Public Sub Popup(Buddy As Control, Optional Top As Single, Optional left As Single)
On Error GoTo eh
    
    Dim tR   As RECT
    
    Set mBuddy = Buddy
    Call GetWindowRect(mBuddy.hwnd, tR)
    
    
    If Top = 0 And left = 0 Then
        'not given so assume bottom left corner
        Me.Move tR.left * Screen.TwipsPerPixelX, tR.Bottom * Screen.TwipsPerPixelY + Screen.TwipsPerPixelY
    Else
        'top left corner + top/left parameters
        Me.Move tR.left * Screen.TwipsPerPixelX + left, tR.Top * Screen.TwipsPerPixelY + Top
    End If
    
    
    If IsDate(mBuddy.Text) Then
        Me.Value = DateValue(mBuddy.Text)
    Else
        Me.Value = Now
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
