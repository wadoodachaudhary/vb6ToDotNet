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

Private mStyle As Integer
Private mBorder As Boolean
Private mDropDownWidth As Integer

Public Sub ClearList()
    Dim i As Long
    For i = Combo1(0).ListCount - 1 To 0 Step -1
        Combo1(0).RemoveItem (i)
        Combo1(2).RemoveItem (i)
    Next

End Sub
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

Public Sub AddItem(item As String, Optional Index)
    Call Combo1(0).AddItem(item, Index)
    Call Combo1(2).AddItem(item, Index)
End Sub

Private Sub UserControl_Resize()
    If Border Then
        UserControl.Height = Combo1(0).Height
        Combo1(0).Move 0, 0, UserControl.Width
    Else
        UserControl.Height = Combo1(0).Height - 5 * Screen.TwipsPerPixelY
        Combo1(0).Move -Screen.TwipsPerPixelX * 2, -Screen.TwipsPerPixelY * 2, UserControl.Width + Screen.TwipsPerPixelX * 5
    End If
    Combo1(2).Move Combo1(0).left, Combo1(0).Top, Combo1(0).Width
End Sub

Private Sub UserControl_InitProperties()
    Style = 0
    Border = False
    DropDownWidth = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    DropDownWidth = PropBag.ReadProperty("DropDownWidth", 0)
    Style = PropBag.ReadProperty("Style", 0)
    Text = PropBag.ReadProperty("Text", "")
    Enabled = PropBag.ReadProperty("Enabled", True)
    Border = PropBag.ReadProperty("Border", False)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DropDownWidth", DropDownWidth, 0)
    Call PropBag.WriteProperty("Style", Style, 0)
    Call PropBag.WriteProperty("Text", Text, "")
    Call PropBag.WriteProperty("Border", Border, False)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
End Sub

Public Property Get DropDownWidth() As Integer
    DropDownWidth = mDropDownWidth
End Property

Public Property Let DropDownWidth(RHS As Integer)
    mDropDownWidth = RHS
    If RHS <> 0 Then
        SetComboDropDownWidth Combo1(0), RHS
        SetComboDropDownWidth Combo1(2), RHS
    End If
End Property


Public Property Get ListCount() As Integer
    ListCount = Combo1(0).ListCount
End Property

Public Property Get list(Index) As String
    list = Combo1(0).list(Index)
End Property
Public Property Let list(Index, RHS As String)
    Combo1(0).list(Index) = RHS
    Combo1(2).list(Index) = RHS
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
End Property

Public Property Get SelLength() As Long
    SelLength = Combo1(0).SelLength
End Property
Public Property Let SelLength(RHS As Long)
    Combo1(0).SelLength = RHS
End Property


