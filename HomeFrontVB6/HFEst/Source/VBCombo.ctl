VERSION 5.00
Begin VB.UserControl VBCombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   1065
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   825
      Width           =   2025
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   420
      Width           =   2025
   End
End
Attribute VB_Name = "VBCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Change()
Event Click()
Event DropDown()
Event KeyDown(KeyCode As Integer, Shift As Integer)

'do not use the actual values as they screw up on some resolutions
Const ScreenTwipsPerPixelX = 15
Const ScreenTwipsPerPixelY = 15



Private mStyle As Integer
Private mBorder As Boolean
Private mDropDownWidth As Long

Private Sub Combo1_Change(Index As Integer)
    RaiseEvent Change
End Sub

Private Sub Combo1_Click(Index As Integer)
    If Index = 2 Then
        Combo1(0).ListIndex = Combo1(2).ListIndex
    Else
        RaiseEvent Click
    End If
End Sub

Private Sub Combo1_DropDown(Index As Integer)
    RaiseEvent DropDown
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Public Property Let Tag(RHS As String)
    Combo1(0).Tag = RHS
End Property
Public Property Get Tag() As String
    Tag = Combo1(0).Tag
End Property


Public Property Let Border(RHS As Boolean)
    mBorder = RHS
    Call UserControl_Resize
End Property
Public Property Get Border() As Boolean
    Border = mBorder
End Property


Public Property Let Style(RHS As ComboBoxConstants)
    mStyle = RHS
    If mStyle = 1 Then mStyle = 2
    Combo1(0).Visible = mStyle = 0
    Combo1(2).Visible = mStyle = 2
End Property
Public Property Get Style() As ComboBoxConstants
    Style = mStyle
End Property

Public Property Let Text(RHS As String)
    Combo1(0).Text = RHS
    'text is readonly on lists
    'Combo1(2).Text = RHS
End Property
Public Property Get Text() As String
    Text = Combo1(0).Text
End Property

Public Property Let ListIndex(RHS As Integer)
    Combo1(0).ListIndex = RHS
    Combo1(2).ListIndex = RHS
End Property
Public Property Get ListIndex() As Integer
    ListIndex = Combo1(0).ListIndex
End Property

Public Property Let Enabled(RHS As Boolean)
    Combo1(0).Enabled = RHS
    Combo1(2).Enabled = RHS
End Property
Public Property Get Enabled() As Boolean
    Enabled = Combo1(0).Enabled
End Property

Public Sub Clear()
    Combo1(0).Clear
    Combo1(2).Clear
End Sub

Public Sub AddItem(Item As String, Optional Index)
    Call Combo1(0).AddItem(Item, Index)
    Call Combo1(2).AddItem(Item, Index)
End Sub

Private Sub UserControl_Resize()
    If Border Then
        UserControl.Height = Combo1(0).Height
        Combo1(0).Move 0, 0, UserControl.Width
    Else
        UserControl.Height = Combo1(0).Height - 5 * ScreenTwipsPerPixelY
        Combo1(0).Move -ScreenTwipsPerPixelX * 2, -ScreenTwipsPerPixelY * 2, UserControl.Width + ScreenTwipsPerPixelX * 5
    End If
    Combo1(2).Move Combo1(0).Left, Combo1(0).Top, Combo1(0).Width
End Sub

Private Sub UserControl_InitProperties()
    Style = 0
    Border = False
    DropDownWidth = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Style = PropBag.ReadProperty("Style", 0)
    Text = PropBag.ReadProperty("Text", "")
    Enabled = PropBag.ReadProperty("Enabled", True)
    Border = PropBag.ReadProperty("Border", False)
    DropDownWidth = PropBag.ReadProperty("DropDownWidth", False)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Style", Style, 0)
    Call PropBag.WriteProperty("Text", Text, "")
    Call PropBag.WriteProperty("Border", Border, False)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("DropDownWidth", DropDownWidth, 0)
End Sub

Public Property Get ListCount() As Integer
    ListCount = Combo1(0).ListCount
End Property

Public Property Get List(Index) As String
    List = Combo1(0).List(Index)
End Property
Public Property Let List(Index, RHS As String)
    Combo1(0).List(Index) = RHS
    Combo1(2).List(Index) = RHS
End Property

Public Property Get ItemData(Index) As Long
    If Index > -1 Then ItemData = Combo1(0).ItemData(Index)
End Property
Public Property Let ItemData(Index, RHS As Long)
    Combo1(0).ItemData(Index) = RHS
    Combo1(2).ItemData(Index) = RHS
End Property

Public Property Get NewIndex() As Integer
    NewIndex = Combo1(0).NewIndex
End Property


Public Property Get SelStart() As Long
    SelStart = Combo1(0).SelStart
End Property
Public Property Let SelStart(RHS As Long)
    Combo1(0).SelStart = RHS
    Combo1(2).SelStart = RHS
End Property

Public Property Get SelLength() As Long
    SelLength = Combo1(0).SelLength
End Property
Public Property Let SelLength(RHS As Long)
    Combo1(0).SelLength = RHS
    Combo1(2).SelLength = RHS
End Property

Public Property Get DropDownWidth() As Long
    DropDownWidth = mDropDownWidth
End Property
Public Property Let DropDownWidth(RHS As Long)
    mDropDownWidth = RHS
    SendMessage Combo1(0).hwnd, CB_SETDROPPEDWIDTH, mDropDownWidth / Screen.TwipsPerPixelX, ByVal 0&
    SendMessage Combo1(2).hwnd, CB_SETDROPPEDWIDTH, mDropDownWidth / Screen.TwipsPerPixelX, ByVal 0&
End Property


