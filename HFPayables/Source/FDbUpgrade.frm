VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FDbUpgrade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Upgrade - <Application Name>"
   ClientHeight    =   5340
   ClientLeft      =   7170
   ClientTop       =   870
   ClientWidth     =   7410
   Icon            =   "FDbUpgrade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNav 
      Caption         =   "< Previous"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   22
      Top             =   4860
      Width           =   1155
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   1
      Left            =   8160
      Picture         =   "FDbUpgrade.frx":0E42
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   -120
      Visible         =   0   'False
      Width           =   7395
      Begin VB.OptionButton optBackup 
         BackColor       =   &H80000005&
         Caption         =   "No, I have not backed up my data"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4080
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optBackup 
         BackColor       =   &H80000005&
         Caption         =   "Yes, I have backed up my data"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2820
         Picture         =   "FDbUpgrade.frx":269FC
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":272C6
         Height          =   3495
         Index           =   1
         Left            =   3480
         TabIndex        =   20
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   3675
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Before you begin the database upgrade..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Index           =   1
         Left            =   2640
         TabIndex        =   21
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   4515
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   4
      Left            =   8760
      Picture         =   "FDbUpgrade.frx":27551
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":4D10B
         Height          =   3135
         Index           =   4
         Left            =   3480
         TabIndex        =   16
         Top             =   1380
         UseMnemonic     =   0   'False
         Width           =   3675
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   2820
         Picture         =   "FDbUpgrade.frx":4D370
         Top             =   1500
         Width           =   480
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "The database upgrade failed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Index           =   4
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   4515
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   3
      Left            =   420
      Picture         =   "FDbUpgrade.frx":4DC3A
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   9600
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":737F4
         Height          =   3135
         Index           =   3
         Left            =   2700
         TabIndex        =   12
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   4455
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "The database upgrade completed successfully"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Index           =   3
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   4515
      End
   End
   Begin VB.CommandButton cmdNav 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   4
      Top             =   4860
      Width           =   1155
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Next >"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   4860
      TabIndex        =   3
      Top             =   4860
      Width           =   1155
   End
   Begin VB.PictureBox picFrame 
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   2
      Left            =   8520
      Picture         =   "FDbUpgrade.frx":73908
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4740
      Visible         =   0   'False
      Width           =   7395
      Begin VSFlex8Ctl.VSFlexGrid gStatements 
         Height          =   2595
         Left            =   660
         TabIndex        =   2
         Top             =   1920
         Width           =   6195
         _cx             =   10927
         _cy             =   4577
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FDbUpgrade.frx":8960A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Image imgNull 
         Height          =   240
         Left            =   5760
         Picture         =   "FDbUpgrade.frx":8965A
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Left            =   5520
         Picture         =   "FDbUpgrade.frx":897A4
         Top             =   1440
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasks"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   18
         Top             =   1680
         Width           =   5955
      End
      Begin VB.Label lblUpgradeDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "These modifications will be made to your database."
         Height          =   510
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   5955
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   600
         Picture         =   "FDbUpgrade.frx":898EE
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":8A5B8
         Height          =   435
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   5415
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Required Changes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   120
         UseMnemonic     =   0   'False
         Width           =   6135
      End
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4710
      Index           =   0
      Left            =   0
      Picture         =   "FDbUpgrade.frx":8A64A
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   493
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   7395
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"FDbUpgrade.frx":B0204
         Height          =   3495
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   4515
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the <Application Name> Database Upgrade Wizard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4275
         Index           =   0
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   4515
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   -200
      X2              =   9719
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   -12
      X2              =   9707
      Y1              =   314
      Y2              =   314
   End
End
Attribute VB_Name = "FDbUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mConnection       As Connection
Private mCompleted        As Boolean
Private mConnectionString As String
Private mUserID           As String
Private mDbRevision       As Long

Private Type StatementType
    Description  As String  'displayed in listbox
    SqlStatement As String  'query to run
    IgnoreErrors As Boolean
End Type
Private mStatements() As StatementType

Private Sub LoadRevisions()
    Dim c As String
    Dim s As String
    ReDim mStatements(0)
    
c = "Initialize Database Schema"
    Sql c, "CREATE TABLE DBRevisions(" & vbCrLf & _
           "  AppKey       VARCHAR(12)" & vbCrLf & _
           " ,Revision     INTEGER" & vbCrLf & _
           " ,RevisionDate TIMESTAMP" & vbCrLf & _
           " ,UserID       VARCHAR(50)" & vbCrLf & _
           " ,Description  VARCHAR(50)" & vbCrLf & _
           " ,SqlStatement LONGVARCHAR)"
           
    
    s = ""
    s = s & "CREATE TABLE AppOptions(" & vbCrLf
    s = s & "  UID           VARCHAR(10) NOT NULL" & vbCrLf
    s = s & " ,OptionName    VARCHAR(35)" & vbCrLf
    s = s & " ,OptionValue   LONGVARCHAR)" & vbCrLf
    Sql c, s
    Sql c, "CREATE UNIQUE INDEX PK_AppOptions ON AppOptions(UID,OptionName)" & vbCrLf
           
           
    s = ""
    s = s & "CREATE TABLE Users(" & vbCrLf
    s = s & "  UID           VARCHAR(10) NOT NULL" & vbCrLf
    s = s & " ,Pswd          VARCHAR(20)" & vbCrLf
    s = s & " ,Name          VARCHAR(20)" & vbCrLf
    s = s & " ,SecurityLevel INTEGER" & vbCrLf
    s = s & " ,OverridePcnt  DOUBLE" & vbCrLf
    s = s & " ,OverrideMax   DOUBLE)" & vbCrLf
    Sql c, s
    Sql c, "CREATE UNIQUE INDEX PK_Users ON Users(UID)" & vbCrLf
    Sql c, "INSERT INTO Users(UID,Pswd,Name) VALUES('Admin','','System Administrator')"
    
    
    s = ""
    s = s & "CREATE TABLE Invoices(" & vbCrLf
    s = s & "  InvoiceID       IDENTITY" & vbCrLf
    s = s & " ,Vendor          VARCHAR(10) NOT NULL CASE" & vbCrLf
    s = s & " ,VendorName      VARCHAR(30)" & vbCrLf
    s = s & " ,Invoice         VARCHAR(15) NOT NULL CASE" & vbCrLf
    s = s & " ,Job             VARCHAR(10)" & vbCrLf
    s = s & " ,JobDesc         VARCHAR(30)" & vbCrLf
    s = s & " ,Status          VARCHAR(10)" & vbCrLf
    s = s & " ,PreTax          DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,Tax             DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,Discount        DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,DiscountDate    DATE" & vbCrLf
    s = s & " ,ReceivedDate    DATE" & vbCrLf
    s = s & " ,InvoiceDate     DATE" & vbCrLf
    s = s & " ,PaymentDate     DATE" & vbCrLf
    s = s & " ,AccountingDate  DATE" & vbCrLf
    s = s & " ,Description     VARCHAR(30)" & vbCrLf
    s = s & " ,UStmp           VARCHAR(10)" & vbCrLf
    s = s & " ,DStmp           DATE" & vbCrLf
    s = s & " ,TStmp           TIME)" & vbCrLf
    Sql c, s
    Sql c, "CREATE UNIQUE INDEX AK1_Invoices ON Invoices(Vendor,Invoice)"
    Sql c, "CREATE INDEX AK2_Invoices ON Invoices(Status,Vendor,Invoice)"
    
    s = ""
    s = s & "CREATE TABLE InvoiceItems(" & vbCrLf
    s = s & "  ItemID           IDENTITY" & vbCrLf
    s = s & " ,InvoiceID        INTEGER NOT NULL" & vbCrLf
    s = s & " ,Vendor           VARCHAR(10) NOT NULL CASE" & vbCrLf
    s = s & " ,Invoice          VARCHAR(15) NOT NULL CASE" & vbCrLf
    s = s & " ,Commitment       VARCHAR(15) CASE" & vbCrLf
    s = s & " ,CommitmentDesc   VARCHAR(30) CASE" & vbCrLf
    s = s & " ,CommitmentItem   INTEGER" & vbCrLf
    s = s & " ,CommitmentVendor VARCHAR(10) CASE" & vbCrLf
    s = s & " ,Job              VARCHAR(10) CASE" & vbCrLf
    s = s & " ,JobDesc          VARCHAR(30) CASE" & vbCrLf
    s = s & " ,Extra            VARCHAR(10) CASE" & vbCrLf
    s = s & " ,ExtraDesc        VARCHAR(30) CASE" & vbCrLf
    s = s & " ,Phase            VARCHAR(12) CASE" & vbCrLf
    s = s & " ,PhaseDesc        VARCHAR(30) CASE" & vbCrLf
    s = s & " ,Category         VARCHAR(3) CASE" & vbCrLf
    s = s & " ,CategoryDesc     VARCHAR(30) CASE" & vbCrLf
    s = s & " ,DebitAccount     VARCHAR(25) CASE" & vbCrLf
    s = s & " ,DebitAccountDesc VARCHAR(30) CASE" & vbCrLf
    s = s & " ,TaxGroup         VARCHAR(6) CASE" & vbCrLf
    s = s & " ,TaxGroupDesc     VARCHAR(30) CASE" & vbCrLf
    s = s & " ,TaxRate          DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,PreTax           DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,Tax              DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,Retainage        DOUBLE DEFAULT 0" & vbCrLf
    s = s & " ,Description      VARCHAR(30)" & vbCrLf
    s = s & " ,UStmp            VARCHAR(10)" & vbCrLf
    s = s & " ,DStmp            DATE" & vbCrLf
    s = s & " ,TStmp            TIME)" & vbCrLf
    Sql c, s
    Sql c, "CREATE INDEX FK1_InvoiceItems ON InvoiceItems(InvoiceID)"
    
c = "Add web invoice support"
    Sql c, "ALTER TABLE Invoices MODIFY Status VARCHAR(15)"
    Sql c, "ALTER TABLE Invoices ADD Errors VARCHAR(250)"
    Sql c, "ALTER TABLE InvoiceItems ADD Errors VARCHAR(250)"
    Sql c, "ALTER TABLE Invoices ADD Source VARCHAR(15)"
    Sql c, "UPDATE Invoices SET Source='Desk Entry'"

c = "Add invoice posting date"
    Sql c, "ALTER TABLE Invoices ADD PostingDate TIMESTAMP"

c = "Add web invoice statuses"
    Sql c, "ALTER TABLE Invoices MODIFY Status VARCHAR(15)"

c = "Enlarge all description fields"
    Sql c, "alter table invoiceitems alter column commitmentdesc   varchar(50)"
    Sql c, "alter table invoiceitems alter column jobdesc          varchar(50)"
    Sql c, "alter table invoiceitems alter column extradesc        varchar(50)"
    Sql c, "alter table invoiceitems alter column phasedesc        varchar(50)"
    Sql c, "alter table invoiceitems alter column categorydesc     varchar(50)"
    Sql c, "alter table invoiceitems alter column debitaccountdesc varchar(50)"
    Sql c, "alter table invoiceitems alter column taxgroupdesc     varchar(50)"
    Sql c, "alter table invoiceitems alter column description      varchar(50)"
    Sql c, "alter table invoices     alter column jobdesc          varchar(50)"
    Sql c, "alter table invoices     alter column description      varchar(50)"
    '-- well that didn't work. do it again dopey ------------------
    
    
    Sql c, "update invoices set vendor = left(vendor ,10)"
    Sql c, "update invoices set vendorname = left(vendorname ,30)"
    Sql c, "update invoices set invoice = left(invoice ,15)"
    Sql c, "update invoices set job = left(job ,10)"
    Sql c, "update invoices set jobdesc = left(jobdesc ,30)"
    Sql c, "update invoices set status = left(status ,10)"
    Sql c, "update invoices set errors = left(errors ,1000)"
    Sql c, "update invoices set source = left(source ,15)"
    Sql c, "update invoices set description = left(description ,30)"
    Sql c, "update invoiceitems set vendor = left(vendor ,10)"
    Sql c, "update invoiceitems set invoice = left(invoice ,15)"
    Sql c, "update invoiceitems set commitment = left(commitment ,12)"
    Sql c, "update invoiceitems set commitmentdesc = left(commitmentdesc ,30)"
    Sql c, "update invoiceitems set commitmentvendor = left(commitmentvendor ,10)"
    Sql c, "update invoiceitems set job = left(job ,10)"
    Sql c, "update invoiceitems set jobdesc = left(jobdesc ,30)"
    Sql c, "update invoiceitems set extra = left(extra ,10)"
    Sql c, "update invoiceitems set extradesc = left(extradesc ,30)"
    Sql c, "update invoiceitems set phase = left(phase ,12)"
    Sql c, "update invoiceitems set phasedesc = left(phasedesc ,30)"
    Sql c, "update invoiceitems set category = left(category ,3)"
    Sql c, "update invoiceitems set categorydesc = left(categorydesc ,30)"
    Sql c, "update invoiceitems set debitaccount = left(debitaccount ,25)"
    Sql c, "update invoiceitems set debitaccountdesc = left(debitaccountdesc ,40)"
    Sql c, "update invoiceitems set taxgroup = left(taxgroup ,6)"
    Sql c, "update invoiceitems set taxgroupdesc = left(taxgroupdesc ,30)"
    Sql c, "update invoiceitems set description = left(description ,30)"
    Sql c, "update invoiceitems set ustmp = left(ustmp ,10)"
    Sql c, "update invoiceitems set errors = left(errors ,1000)"
    
    
    Sql c, "alter table invoices alter column vendor varchar(10)"
    Sql c, "alter table invoices alter column vendorname varchar(30)"
    Sql c, "alter table invoices alter column invoice varchar(15)"
    Sql c, "alter table invoices alter column job varchar(10)"
    Sql c, "alter table invoices alter column jobdesc varchar(30)"
    Sql c, "alter table invoices alter column status varchar(15)"
    Sql c, "alter table invoices alter column errors varchar(254)"
    Sql c, "alter table invoices alter column source varchar(15)"
    Sql c, "alter table invoices alter column description varchar(30)"
    Sql c, "alter table invoiceitems alter column vendor varchar(10)"
    Sql c, "alter table invoiceitems alter column invoice varchar(15)"
    
    Sql c, "drop index invoiceitems.ak1_invoiceitems", True
    Sql c, "drop index ak1_invoiceitems", True
    Sql c, "alter table invoiceitems alter column commitment varchar(12)"
    Sql c, "create index ak1_invoiceitems on InvoiceItems(Commitment,CommitmentItem)"
    
    Sql c, "alter table invoiceitems alter column commitmentdesc varchar(30)"
    Sql c, "alter table invoiceitems alter column commitmentvendor varchar(10)"
    Sql c, "alter table invoiceitems alter column job varchar(10)"
    Sql c, "alter table invoiceitems alter column jobdesc varchar(30)"
    Sql c, "alter table invoiceitems alter column extra varchar(10)"
    Sql c, "alter table invoiceitems alter column extradesc varchar(30)"
    Sql c, "alter table invoiceitems alter column phase varchar(12)"
    Sql c, "alter table invoiceitems alter column phasedesc varchar(30)"
    Sql c, "alter table invoiceitems alter column category varchar(3)"
    Sql c, "alter table invoiceitems alter column categorydesc varchar(30)"
    Sql c, "alter table invoiceitems alter column debitaccount varchar(25)"
    Sql c, "alter table invoiceitems alter column debitaccountdesc varchar(40)"
    Sql c, "alter table invoiceitems alter column taxgroup varchar(6)"
    Sql c, "alter table invoiceitems alter column taxgroupdesc varchar(30)"
    Sql c, "alter table invoiceitems alter column description varchar(30)"
    Sql c, "alter table invoiceitems alter column ustmp varchar(10)"
    Sql c, "alter table invoiceitems alter column errors varchar(254)"
    
c = "Enlarge userid Field"
    Sql c, "alter table Users        alter column UID   varchar(20)"
    Sql c, "alter table invoices     alter column UStmp varchar(20)"
    Sql c, "alter table invoiceitems alter column UStmp varchar(20)"
    Sql c, "alter table appoptions   alter column UID   varchar(20)"
    
c = "Add support for additional invoice codes"
    Sql c, "alter table invoices add InvoiceCode1 varchar(10)"
    Sql c, "alter table invoices add InvoiceCode2 varchar(10)"
    
c = "Add optional support for unit measurements"
    Sql c, "ALTER TABLE InvoiceItems DROP COLUMN Units", True
    Sql c, "ALTER TABLE InvoiceItems DROP COLUMN UnitCost", True
    Sql c, "ALTER TABLE InvoiceItems ADD CommittedQuantity  INTEGER DEFAULT 0"
    Sql c, "ALTER TABLE InvoiceItems ADD CommittedUnitPrice DOUBLE DEFAULT 0"
    Sql c, "ALTER TABLE InvoiceItems ADD InvoicedQuantity   INTEGER DEFAULT 0"
    Sql c, "ALTER TABLE InvoiceItems ADD InvoicedUnitPrice  DOUBLE DEFAULT 0"
    
c = "Add unique constraints (optional) to misc invoice codes"
    Sql c, "CREATE INDEX AK3_Invoices on Invoices(InvoiceCode1)"
    Sql c, "CREATE INDEX AK4_Invoices on Invoices(InvoiceCode2)"
        
c = "Add invoice approval processes"
    Sql c, "ALTER TABLE Invoices ADD Approver         VARCHAR(20)"
    Sql c, "ALTER TABLE Invoices ADD HoldReason       VARCHAR(250)"
    Sql c, "ALTER TABLE Invoices ADD ApproverComments VARCHAR(250)"
    Sql c, "ALTER TABLE Invoices ADD DateApproved     DATE"
    Sql c, "ALTER TABLE Users ADD EmailAddress VARCHAR(250)"
        
c = "Add advance payment invoicing"
    Sql c, "ALTER TABLE Invoices ADD Advance BIT DEFAULT 0"
    Sql c, "UPDATE Invoices SET Advance=0"
    
    Sql c, "ALTER TABLE InvoiceItems ADD RetainageRate DOUBLE DEFAULT 0"
    Sql c, "UPDATE InvoiceItems SET RetainageRate=0"
    
    Sql c, "ALTER TABLE InvoiceItems ADD TaxRetainageRate DOUBLE DEFAULT 0"
    Sql c, "UPDATE InvoiceItems SET TaxRetainageRate=0"
    
    Sql c, "ALTER TABLE Invoices ADD BatchNumber VARCHAR(30)"
    
    
c = "Add invoice splitting feature"
    Sql c, "ALTER TABLE InvoiceItems ADD IsSplit BIT DEFAULT 0"
    Sql c, "UPDATE InvoiceItems SET IsSplit=0"
    
c = "Add support for misc vendor deductions"
    Sql c, "ALTER TABLE InvoiceItems ADD IsWCB BIT DEFAULT 0"
    Sql c, "UPDATE InvoiceItems SET IsWCB=0"
    Sql c, "ALTER TABLE InvoiceItems drop column IsWCB", True
    Sql c, "ALTER TABLE InvoiceItems drop column IsSplit", True
    
c = "Add import invoice function"
    Sql c, "alter table invoiceitems add CreditAccount varchar(25)"
    Sql c, "alter table invoiceitems drop column CreditAccount"
    
    
Sql c, "ALTER TABLE Invoices ADD DateApproved DATE ", True
Sql c, "ALTER TABLE Invoices drop column DatetimeApproved", True
        
    
End Sub

Private Sub Sql(Description As String, SqlStatement As String, Optional IgnoreErrors As Boolean)
    Dim i As Long
    i = UBound(mStatements) + 1
    ReDim Preserve mStatements(i)
    mStatements(i).Description = Description
    mStatements(i).SqlStatement = SqlStatement
    mStatements(i).IgnoreErrors = IgnoreErrors
End Sub

Public Function Upgrade(ConnectionString As String, Optional UserID As String) As Boolean
    Dim r      As Long
    Dim prev   As String
    
    'open connection
    Set mConnection = New Connection
    mConnection.Open ConnectionString
    If Not IsCorrectApp() Then
        On Error Resume Next
        mConnection.Close
        VBA.MsgBox "This doesn't appear to be an " & App.ProductName & " database." & vbCrLf & _
                   "Please verify that this DSN points to the correct location.", vbExclamation, App.ProductName
        Upgrade = False
        Exit Function
    End If
    
    'save parameters
    mConnectionString = ConnectionString
    mUserID = UserID
        
    'reposition and initialize frames
    For r = picFrame.LBound To picFrame.UBound
        picFrame(r).Move 0, 0
    Next
    Call InitLabels
    Call LoadRevisions
    
    'get revision number
    On Error Resume Next
    mDbRevision = mConnection.Execute("SELECT MAX(Revision) FROM DBRevisions WHERE AppKey=" & DbQuote(Str, APPKEY))(0)
    
    'show screen if there are 'un-run' statements
    If mDbRevision < UBound(mStatements) Then
        
        'load statements into grid
        gStatements.Rows = 0
        For r = mDbRevision + 1 To UBound(mStatements)
            If mStatements(r).Description <> prev Then
                gStatements.AddItem vbTab & mStatements(r).Description & vbTab & r
                Set gStatements.Cell(flexcpPicture, r, 0) = imgNull.Picture
                prev = mStatements(r).Description
            End If
        Next
        
        'show form
        mCompleted = False
        Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
        CurPage = 0
        Screen.MousePointer = vbNormal
        Me.Show vbModal
        Upgrade = mCompleted
        
    Else
        Upgrade = True
    End If
    
    On Error Resume Next
    mConnection.Close
    Set mConnection = Nothing
    
    Unload Me
End Function


Private Sub cmdNav_Click(Index As Integer)
    Select Case Index
        Case 0 'previous
            CurPage = Max(0, CurPage - 1)
            
        Case 1 'next
            If CurPage = 2 Then
                cmdNav(0).Enabled = False
                cmdNav(1).Enabled = False
                cmdNav(2).Enabled = False
                mCompleted = RunStatements()
                If mCompleted Then
                    CurPage = 3 'succeed
                Else
                    CurPage = 4 'fail
                End If
            Else
                'move next
                CurPage = Min(2, CurPage + 1)
            End If
            
        Case 2 'cancel
            Unload Me
            
    End Select
End Sub

Private Function RunStatements() As Boolean
    Dim i As Long
    Dim r As Long
    Dim s As String
    
    Screen.MousePointer = vbHourglass
    
    lblTitle(2) = "Upgrading Database"
    lblTitle(2).Refresh
    
    lblDescription(2) = "Please wait while your " & AppName & " database is upgraded."
    lblDescription(2).Refresh
    
    With gStatements
        For i = 0 To .Rows - 1
            Call .Select(i, 0)
            Call .ShowCell(i, 0)
            
            r = Val(.TextMatrix(i, 2))
            On Error Resume Next
            While mStatements(r).Description = .TextMatrix(i, 1)
                
                If r > UBound(mStatements) Then
                    Screen.MousePointer = vbDefault
                    RunStatements = True
                    Exit Function
                Else
                    lblUpgradeDesc = "Upgrading database... " & vbCrLf & "(step " & r + 1 - Val(.TextMatrix(0, 2)) & " of " & UBound(mStatements) + 1 - Val(.TextMatrix(0, 2)) & ")"
                    lblUpgradeDesc.Refresh
                
                    s = mStatements(r).SqlStatement
                    Call App.ConvertSQL(s, mConnection)
                    mConnection.CommandTimeout = 600
                    mConnection.Execute s
                    If Err.Number = 0 Or mStatements(r).IgnoreErrors Then
                        
                        s = ""
                        s = s & "INSERT INTO DBRevisions(AppKey,Revision,RevisionDate,UserID,Description,SqlStatement)" & vbCrLf
                        s = s & "VALUES(" & DbQuote(Str, APPKEY) & vbCrLf
                        s = s & "      ," & r & vbCrLf
                        s = s & "      ,NOW()" & vbCrLf
                        s = s & "      ," & DbQuote(Str, mUserID) & vbCrLf
                        s = s & "      ," & DbQuote(Str, Mid(mStatements(r).Description, 1, 50)) & vbCrLf
                        s = s & "      ," & DbQuote(Str, mStatements(r).SqlStatement) & ")"
                        Call App.ConvertSQL(s, mConnection)
                        mConnection.Execute s

                    Else
                        Screen.MousePointer = vbDefault
                        MsgBox AppName & " Database Upgrade Wizard has encountered a problem" & vbCrLf & _
                               "and cannot complete the requested action. Ensure all other users are logged" & vbCrLf & _
                               "out of the system then run this wizard again. If the problem persists contact" & vbCrLf & _
                               "HomeFront support." & vbCrLf & vbCrLf & _
                               "Error: " & Trim(Parse(Err.Description, Parse(Err.Description, , "]"), "]")) & vbCrLf _
                               , vbExclamation, AppName & " Database Upgrade Wizard"
                        Exit Function
                    End If
                    r = r + 1
                End If
            Wend
            Set .Cell(flexcpPicture, i, 0) = imgCheck.Picture
        Next
    End With
    
    
End Function

Private Property Get CurPage() As Integer
    Dim i As Integer
    For i = picFrame.LBound To picFrame.UBound
        If picFrame(i).Visible Then
            CurPage = i
            Exit Property
        End If
    Next
End Property

Private Property Let CurPage(RHS As Integer)
    
    Dim i As Integer
    For i = picFrame.LBound To picFrame.UBound
        picFrame(i).Visible = i = RHS
    Next
    
    Select Case RHS
        
        Case 0 'welcome
            cmdNav(1).Caption = "< Previous":    cmdNav(0).Enabled = False:    cmdNav(0).Visible = False
            cmdNav(1).Caption = "Next >":        cmdNav(1).Enabled = True:     cmdNav(1).Visible = True
            cmdNav(2).Caption = "Cancel":        cmdNav(2).Enabled = True:     cmdNav(2).Visible = True
            
        Case 1  'warning
            cmdNav(1).Caption = "< Previous":    cmdNav(0).Enabled = True:     cmdNav(0).Visible = True
            cmdNav(1).Caption = "Next >":        cmdNav(1).Enabled = False:    cmdNav(1).Visible = True
            cmdNav(2).Caption = "Cancel":        cmdNav(2).Enabled = True:     cmdNav(2).Visible = True
            optBackup(1).Value = True
            
        Case 2 'upgrading
            cmdNav(1).Caption = "< Previous":    cmdNav(0).Enabled = True:     cmdNav(0).Visible = True
            cmdNav(1).Caption = "OK":            cmdNav(1).Enabled = True:     cmdNav(1).Visible = True
            cmdNav(2).Caption = "Cancel":        cmdNav(2).Enabled = True:     cmdNav(2).Visible = True
        
        Case 3 'finished
            cmdNav(1).Caption = "< Previous":    cmdNav(0).Enabled = False:    cmdNav(0).Visible = False
            cmdNav(1).Caption = "Next >":        cmdNav(1).Enabled = False:    cmdNav(1).Visible = False
            cmdNav(2).Caption = "Close":         cmdNav(2).Enabled = True:     cmdNav(2).Visible = True
            cmdNav(2).SetFocus
            
        Case 4 'failed
            cmdNav(1).Caption = "< Previous":    cmdNav(0).Enabled = False:    cmdNav(0).Visible = False
            cmdNav(1).Caption = "Next >":        cmdNav(1).Enabled = False:    cmdNav(1).Visible = False
            cmdNav(2).Caption = "Close":         cmdNav(2).Enabled = True:     cmdNav(2).Visible = True
            cmdNav(2).SetFocus
            
    End Select
    
End Property

Private Property Get AppName() As String
    AppName = App.ProductName
End Property

Private Property Get AppVersion() As String
    AppVersion = App.Major & "." & App.Minor & ".0." & App.Revision
End Property

Private Sub InitLabels()

    Me.Caption = AppName & " - Database Upgrade"

    lblTitle(0) = "Welcome to the " & AppName & " Database Upgrade Wizard"
    lblDescription(0) = "The database you have selected must be upgraded before it can be used by this version of " & AppName & "." & vbCrLf & vbCrLf & _
                        "This wizard will upgrade your " & AppName & " database to the current version specifications (version " & AppVersion & ")" & vbCrLf & vbCrLf & _
                        "These upgrades are fully backward compatible. After running the upgrade wizard, previous versions of " & AppName & " will continue to function as they did before." & vbCrLf & vbCrLf & _
                        "Click Next to continue, or Cancel to exit."
    
    lblTitle(1) = "Before you begin the database upgrade..."
    lblDescription(1) = "Before continuing you must close all open licenses of " & AppName & ". The database cannot be upgraded if anyone is logged on to the system." & vbCrLf & vbCrLf & _
                        "BACK UP YOUR DATA!!!" & vbCrLf & _
                        "Every effort has been made to ensure the integrity of this upgrade but there is always the risk of a failure causing database corruption. In such an event HomeFront will not be held liable and likely won't be able to recover your damaged data." & vbCrLf & vbCrLf & _
                        "Click Next to continue, or Cancel to exit."
    
    lblTitle(2) = "Required Changes"
    lblDescription(2) = "Review the list of database changes required, then click OK to begin the upgrade or Cancel to exit."
    lblUpgradeDesc = "These modifications will be made to your database."
    
    lblTitle(3) = "The database upgrade completed successfully"
    lblDescription(3) = "The " & AppName & " database upgrade wizard has finished upgrading your database."
    
    lblTitle(4) = "The database upgrade failed"
    lblDescription(4) = "The " & AppName & " database upgrade wizard encountered problems while upgrading your database." & vbCrLf & vbCrLf & _
                        "Database tables may be locked by another user or you may not have file system permissions to make the required modifications." & vbCrLf & vbCrLf & _
                        "Ensure all other users are logged out of the system then run this wizard again while logged in with administrator permissions."

End Sub

Private Function Min(This, That)
    Min = IIf(This < That, This, That)
End Function
Private Function Max(This, That)
    Max = IIf(This > That, This, That)
End Function




Private Sub Form_Initialize()
'    HookAttach
End Sub
Private Sub Form_Load()
'    HookDetach
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Me.Hide
End Sub

Private Sub optBackup_Click(Index As Integer)
    cmdNav(1).Enabled = optBackup(0).Value
End Sub

Private Function IsCorrectApp() As Boolean
On Error Resume Next
    IsCorrectApp = True 'APPKEY = UCase(mConnection.Execute("SELECT OptionValue FROM AppOptions WHERE OptionName='APPKEY'")(0))
End Function


