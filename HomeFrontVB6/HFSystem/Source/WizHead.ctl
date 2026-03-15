VERSION 5.00
Begin VB.UserControl WizHead 
   Alignable       =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4845
      Stretch         =   -1  'True
      Top             =   195
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -8
      X2              =   15974.67
      Y1              =   58
      Y2              =   58
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillColor       =   &H00800000&
      Height          =   690
      Left            =   4740
      Top             =   90
      Width           =   690
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export vendor pricelists to Excel spreadsheets"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   420
      Width           =   3240
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Export Vendor Pricelists"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2040
   End
End
Attribute VB_Name = "WizHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Description() As String
    Description = lblDescription.Caption
End Property

Public Property Let Description(RHS As String)
    lblDescription.Caption = RHS
    PropertyChanged "Description"
End Property
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property
Public Property Let Caption(RHS As String)
    lblCaption.Caption = RHS
    PropertyChanged "Caption"
End Property
Public Property Get Picture() As StdPicture
    Set Picture = Image1.Picture
End Property
Public Property Set Picture(RHS As StdPicture)
    Set Image1.Picture = RHS
    PropertyChanged "Picture"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "Caption")
    Description = PropBag.ReadProperty("Description", "Description")
    Set Picture = PropBag.ReadProperty("Icon")
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 60 * Screen.TwipsPerPixelY
    Shape1.Move UserControl.ScaleWidth - Shape1.Width - 6, 6
    Image1.Move Shape1.left + 7, Shape1.Top + 7
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, "Caption")
    Call PropBag.WriteProperty("Description", lblDescription.Caption, "Description")
    Call PropBag.WriteProperty("Icon", Image1.Picture)
End Sub
