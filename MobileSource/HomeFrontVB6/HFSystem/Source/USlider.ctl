VERSION 5.00
Begin VB.UserControl Slider 
   CanGetFocus     =   0   'False
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   120
   MousePointer    =   7  'Size N S
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   8
   ToolboxBitmap   =   "USlider.ctx":0000
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum slOrientation
    Vertical = 0
    Horizontal = 1
End Enum

'Property Variables
Private mOrientation As slOrientation
Private mMin  As Single
Private mMax  As Single

'Property Defaults
Private Const mOrientation_def = Vertical

'Private Junk
Private mOrigX As Double
Private mOrigY As Double

Public Event Move()

Public Property Get Min() As Double
    Min = mMin
End Property
Public Property Let Min(RHS As Double)
    mMin = RHS
    PropertyChanged "Min"
End Property

Public Property Get Max() As Double
    Max = mMax
End Property
Public Property Let Max(RHS As Double)
    mMax = RHS
    PropertyChanged "Max"
End Property


Public Property Get Orientation() As slOrientation
    Orientation = mOrientation
End Property
Public Property Let Orientation(RHS As slOrientation)
    Dim od As Single
    If RHS = Horizontal Then
        UserControl.MousePointer = vbSizeNS
    Else
        UserControl.MousePointer = vbSizeWE
    End If
    mOrientation = RHS
    PropertyChanged "Orientation"
End Property


Private Sub UserControl_AmbientChanged(PropertyName As String)
    UserControl.ScaleMode = ParentScaleMode
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mOrigX = x
    mOrigY = y
    UserControl.BackColor = vbApplicationWorkspace
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim NewValue As Double
    
    If Button = vbLeftButton Then
        If mOrientation = Vertical Then
            NewValue = UserControl.Extender.left + x - mOrigX
            NewValue = MaxVal(NewValue, mMin)
            NewValue = MinVal(NewValue, mMax - 2 * UserControl.ScaleWidth)
            UserControl.Extender.left = NewValue
        Else
            NewValue = UserControl.Extender.Top + y - mOrigY
            NewValue = MaxVal(NewValue, mMin)
            NewValue = MinVal(NewValue, mMax - UserControl.ScaleHeight)
            UserControl.Extender.Top = NewValue
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl.BackColor = vb3DFace
    RaiseEvent Move
End Sub

Private Sub UserControl_Paint()
    If Not Ambient.UserMode Then
        With UserControl
            .ScaleMode = vbPixels
            UserControl.Line (0, 0)-(.ScaleWidth - 1, .ScaleHeight - 1), vb3DShadow, B
            UserControl.Line (0, 0)-(.ScaleWidth - 1, 0), vb3DHighlight
            UserControl.Line (0, 0)-(0, .ScaleHeight - 1), vb3DHighlight
            .ScaleMode = ParentScaleMode
        End With
    End If
End Sub

Private Sub UserControl_InitProperties()
    mOrientation = mOrientation_def
    mMin = UserControl.Parent.ScaleLeft
    mMax = UserControl.Parent.ScaleWidth
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Orientation = PropBag.ReadProperty("Orientation", mOrientation_def)
    mMin = PropBag.ReadProperty("Min", UserControl.Parent.ScaleLeft)
    mMax = PropBag.ReadProperty("Max", UserControl.Parent.ScaleWidth)
    UserControl.ScaleMode = ParentScaleMode
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Orientation", mOrientation, 0)
    Call PropBag.WriteProperty("Min", mMin, UserControl.Parent.ScaleLeft)
    Call PropBag.WriteProperty("Max", mMax, UserControl.Parent.ScaleWidth)
End Sub

Private Function ParentScaleMode() As ScaleModeConstants
    Select Case Ambient.ScaleUnits
        Case "User":        ParentScaleMode = vbUser
        Case "Twip":        ParentScaleMode = vbTwips
        Case "Point":       ParentScaleMode = vbPoints
        Case "Pixel":       ParentScaleMode = vbPixels
        Case "Character":   ParentScaleMode = vbCharacters
        Case "Inch":        ParentScaleMode = vbInches
        Case "Millimeter":  ParentScaleMode = vbMillimeters
        Case "Centimeter":  ParentScaleMode = vbCentimeters
    End Select
End Function
Private Function MinVal(a, b)
    If a < b Then
        MinVal = a
    Else
        MinVal = b
    End If
End Function
Private Function MaxVal(a, b)
    If a > b Then
        MaxVal = a
    Else
        MaxVal = b
    End If
End Function

