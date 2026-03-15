VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm FMain 
   BackColor       =   &H8000000C&
   Caption         =   "Precision Builder"
   ClientHeight    =   7545
   ClientLeft      =   3375
   ClientTop       =   2685
   ClientWidth     =   12360
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picWarningBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   12360
      TabIndex        =   3
      Top             =   0
      Width           =   12360
      Begin VB.Image Image1 
         Height          =   240
         Left            =   105
         Picture         =   "FMain.frx":08CA
         Top             =   60
         Width           =   240
      End
      Begin VB.Label lblWarningBanner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"FMain.frx":0E54
         Height          =   195
         Left            =   420
         TabIndex        =   4
         Top             =   90
         Width           =   11550
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   4560
   End
   Begin VB.PictureBox CommandPanel 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7155
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   390
      Width           =   3780
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   264
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2952
         _ExtentX        =   5212
         _ExtentY        =   476
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "close"
               ImageKey        =   "close"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "expand"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "shrink"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin cPopMenu6.PopMenu PopMenu 
         Left            =   570
         Top             =   675
         _ExtentX        =   1058
         _ExtentY        =   1058
         HighlightCheckedItems=   0   'False
         TickIconIndex   =   0
      End
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   0
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   68
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":0F02
               Key             =   "SaveAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":17DC
               Key             =   "snapshots"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1819E
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":18A78
               Key             =   "AssemblyCosts"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":19352
               Key             =   "MassChange"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":19C2C
               Key             =   "FieldPOs"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1A506
               Key             =   "NewRFQ"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1ADE0
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1B6BA
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1BF94
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1C86E
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1D148
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1DA22
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1E2FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1EBD6
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1F4B0
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":1FD8A
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":20664
               Key             =   "LookupPublished"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":20F3E
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":21818
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":220F2
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":229CC
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":232A6
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":23B80
               Key             =   "SendPOs"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2445A
               Key             =   "SendRFQs"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":24D34
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2560E
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":25EE8
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":267C2
               Key             =   "TakeoffPlanSwift"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2709C
               Key             =   "TakeoffPipeline"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":27976
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":28250
               Key             =   "TakeoffItemChart"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":28B2A
               Key             =   "New"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":29404
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":29CDE
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2A5B8
               Key             =   "Import"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2AE92
               Key             =   "Export"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2B76C
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2C046
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2C920
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2D1FA
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2DAD4
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2E3AE
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2EC88
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2F562
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":2FE3C
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":30716
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":30FF0
               Key             =   "View"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":318CA
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":321A4
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":32A7E
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":33358
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":33C32
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3450C
               Key             =   "OptionWiz"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3688E
               Key             =   ""
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":37168
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":37A42
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3831C
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":38BF6
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":394D0
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":39DAA
               Key             =   "MasterBuilder"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3A684
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3AF5E
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3B838
               Key             =   "Quickbooks"
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3C112
               Key             =   "Sage50"
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3C9EC
               Key             =   "Approve"
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3D2C6
               Key             =   "Decline"
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3DBA0
               Key             =   "BuildPro"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList SmallIcons 
         Left            =   0
         Top             =   1350
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   82
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3E47A
               Key             =   "option"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3EA14
               Key             =   "QBO"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3F2EE
               Key             =   "combo1"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3F448
               Key             =   "Custom Requests"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":3FD22
               Key             =   "MassChange"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":402BC
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":40856
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":40DF0
               Key             =   "sendreceive"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4138A
               Key             =   "communitystandards"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":41924
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":421FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":42AD8
               Key             =   "pricelists"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":433B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":43C8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":44566
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":44E40
               Key             =   "MB"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":453DA
               Key             =   "Intacct"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":45CB4
               Key             =   "QB"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4624E
               Key             =   "custom"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":467E8
               Key             =   "assembly"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":46D82
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4731C
               Key             =   "EditIntersection"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":478B6
               Key             =   "links"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":47E50
               Key             =   "ItemDB"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":483EA
               Key             =   "sendpos"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":48984
               Key             =   "HelpSearch"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":48F1E
               Key             =   "Items"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":494B8
               Key             =   "New"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":49A52
               Key             =   "Edit"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":49FEC
               Key             =   "HelpContents"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4A586
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4AB20
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4B0BA
               Key             =   "shrink"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4B654
               Key             =   "preview"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4BBEE
               Key             =   "customer"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4C188
               Key             =   "close"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4C722
               Key             =   "expand"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4CCBC
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4D256
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4D7F0
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4DD8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4E324
               Key             =   "estimating"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4E8BE
               Key             =   "jobcost"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4EE58
               Key             =   "error"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4F3F2
               Key             =   "salesworksheet"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4F98C
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":4FF26
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":504C0
               Key             =   "newworksheet"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":50A5A
               Key             =   "worksheet"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":50FF4
               Key             =   "purchaseorder"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5158E
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":51B28
               Key             =   "groupphase"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":520C2
               Key             =   "costcode"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5265C
               Key             =   "job"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":52BF6
               Key             =   "information"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":53190
               Key             =   "item"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5372A
               Key             =   "itemchecked"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":53CC4
               Key             =   "phase"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5425E
               Key             =   "vendor"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":547F8
               Key             =   "warning"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":54D92
               Key             =   "question"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5532C
               Key             =   "category"
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":558C6
               Key             =   "folder"
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":55E60
               Key             =   "AddItems"
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":563FA
               Key             =   "model"
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":56994
               Key             =   "area"
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":56F2E
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":574C8
               Key             =   "Underline"
            EndProperty
            BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":57622
               Key             =   "Bold"
            EndProperty
            BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5777C
               Key             =   "AlignCenter"
            EndProperty
            BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":578D6
               Key             =   "Italic"
            EndProperty
            BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":57A30
               Key             =   "AlignLeft"
            EndProperty
            BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":57B8A
               Key             =   "Bullet"
            EndProperty
            BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":57CE4
               Key             =   "BulletNumber"
            EndProperty
            BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":57E3E
               Key             =   "AlignRight"
            EndProperty
            BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":57F98
               Key             =   "printer"
            EndProperty
            BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":58532
               Key             =   "defaultprinter"
            EndProperty
            BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":58ACC
               Key             =   "costcodes"
            EndProperty
            BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":593A6
               Key             =   "defaultvendors"
            EndProperty
            BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":59C80
               Key             =   "cellcomments"
            EndProperty
            BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5A21A
               Key             =   "editpos"
            EndProperty
            BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5AAF4
               Key             =   "fieldpos"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList MultiStateIcons 
         Left            =   0
         Top             =   2040
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5B3CE
               Key             =   "K00"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5B528
               Key             =   "K01"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5B682
               Key             =   "K02"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5B7DC
               Key             =   "K10"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5B936
               Key             =   "K11"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5BA90
               Key             =   "K12"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5BBEA
               Key             =   "K20"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5BD44
               Key             =   "K21"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5BE9E
               Key             =   "K22"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5BFF8
               Key             =   "E"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5C592
               Key             =   "EV"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5CB2C
               Key             =   "none"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMain.frx":5D0C6
               Key             =   "V"
            EndProperty
         EndProperty
      End
      Begin VSFlex8Ctl.VSFlexGrid CommandBar 
         Height          =   5355
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   3555
         _cx             =   1976834175
         _cy             =   1976837350
         Appearance      =   1
         BorderStyle     =   0
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
         BackColor       =   -2147483645
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483645
         BackColorAlternate=   -2147483645
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483645
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FMain.frx":5D660
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   6
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   1
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuFileSub 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuConstruction 
      Caption         =   "Tasks"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuConstructionSub 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuSalesPricing 
      Caption         =   "Sales Pricing"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuSalesPricingSub 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Setup"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuSetupSub 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuInquiries 
      Caption         =   "Inquiries"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuInquiriesSub 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuReportsSub 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolLinks 
         Caption         =   "(dynamic)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowSub 
         Caption         =   "Show Task Panel"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuWindowSub 
         Caption         =   "Show Workflow"
         Index           =   1
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpSub 
         Caption         =   "Help"
         Index           =   3
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "Support"
         Index           =   4
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "About..."
         Index           =   6
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   ""
      Begin VB.Menu mnuTakeoff 
         Caption         =   "<mnuTakeOff>"
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "Find..."
            Index           =   0
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "View Assemblies"
            Index           =   3
            Begin VB.Menu mnuTakeOffModelsView 
               Caption         =   "Assemblies"
               Index           =   0
            End
            Begin VB.Menu mnuTakeOffModelsView 
               Caption         =   "Assembly Specific Extras"
               Index           =   1
            End
            Begin VB.Menu mnuTakeOffModelsView 
               Caption         =   "Global Extras"
               Index           =   2
            End
            Begin VB.Menu mnuTakeOffModelsView 
               Caption         =   "Design Center"
               Index           =   3
            End
         End
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "View Quotes && Jobs"
            Index           =   4
            Begin VB.Menu mnuTakeOffJobView 
               Caption         =   "Quotes"
               Index           =   4
            End
            Begin VB.Menu mnuTakeOffJobView 
               Caption         =   "Open Jobs"
               Index           =   5
            End
            Begin VB.Menu mnuTakeOffJobView 
               Caption         =   "Custom Options"
               Index           =   6
            End
         End
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "View Items"
            Index           =   6
            Begin VB.Menu mnuTakeOffItemView 
               Caption         =   "Sort by Group and Phase"
               Index           =   8
            End
            Begin VB.Menu mnuTakeOffItemView 
               Caption         =   "Sort by Purchase Order"
               Index           =   9
            End
            Begin VB.Menu mnuTakeOffItemView 
               Caption         =   "Sort by Cost Code"
               Index           =   10
            End
            Begin VB.Menu mnuTakeOffItemView 
               Caption         =   "Sort by Description"
               Index           =   11
            End
         End
         Begin VB.Menu mnuTakeOffEstimateView 
            Caption         =   "View Estimates"
            Index           =   12
         End
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "<mnuGrid>"
         Begin VB.Menu mnuGridSub 
            Caption         =   "Hide this column"
            Index           =   4
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Insert a column"
            Index           =   5
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Rename this column..."
            Index           =   6
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Print..."
            Index           =   8
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Save As..."
            Index           =   9
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Attachments"
            Index           =   10
         End
      End
      Begin VB.Menu mnuGrid2 
         Caption         =   "<mnuGrid2>"
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Group on this column"
            Index           =   0
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Expand All"
            Index           =   1
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Collapse All"
            Index           =   2
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Hide this column"
            Index           =   4
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Insert a column"
            Index           =   5
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Rename this column..."
            Index           =   6
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Print..."
            Index           =   8
         End
         Begin VB.Menu mnuGrid2Sub 
            Caption         =   "Save As..."
            Index           =   9
         End
      End
      Begin VB.Menu mnuItems 
         Caption         =   "<FItems>"
         Begin VB.Menu mnuFItemsPhases 
            Caption         =   "<Phases Tree>"
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "Rename..."
               Index           =   0
            End
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "Renumber..."
               Index           =   1
            End
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "-"
               Index           =   2
            End
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "New Group..."
               Index           =   3
            End
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "New Phase..."
               Index           =   4
            End
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "-"
               Index           =   6
            End
            Begin VB.Menu mnuFItemsPhasesSub 
               Caption         =   "Delete"
               Index           =   7
            End
         End
         Begin VB.Menu mnuFItemsItems 
            Caption         =   "<Item Grid>"
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "Renumber..."
               Index           =   0
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "New Item..."
               Index           =   1
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "Make a Copy..."
               Index           =   2
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "Attachments..."
               Index           =   4
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "-"
               Index           =   7
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "Delete"
               Index           =   8
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "-"
               Index           =   9
            End
            Begin VB.Menu mnuFItemsItemsSub 
               Caption         =   "Price links"
               Index           =   10
               Begin VB.Menu mnuFItemsPriceGroupSub 
                  Caption         =   "Make Group"
                  Index           =   0
               End
               Begin VB.Menu mnuFItemsPriceGroupSub 
                  Caption         =   "Join Group..."
                  Index           =   1
               End
               Begin VB.Menu mnuFItemsPriceGroupSub 
                  Caption         =   "Ungroup"
                  Index           =   2
               End
            End
         End
      End
      Begin VB.Menu mnuEstimateItems 
         Caption         =   "<FEstimateItems>"
         Begin VB.Menu mnuInvoices 
            Caption         =   "<mnuInvoices>"
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "Preview..."
               Index           =   0
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "Print..."
               Index           =   1
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "-"
               Index           =   2
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "New..."
               Index           =   3
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "Edit..."
               Index           =   4
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "Delete"
               Index           =   5
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "Void"
               Index           =   6
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "-"
               Index           =   7
            End
            Begin VB.Menu mnuInvoicesSub 
               Caption         =   "Post"
               Index           =   8
            End
         End
         Begin VB.Menu mnuEstimateItemsAssembly 
            Caption         =   "<Assemblies>"
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "New Change Order"
               Index           =   0
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "New Change Request"
               Index           =   1
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "New Assembly"
               Index           =   2
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "Delete"
               Index           =   4
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "Rename..."
               Index           =   5
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "-"
               Index           =   6
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "Finalize Budgets"
               Index           =   7
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "Attach Quote..."
               Index           =   8
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "Copy From Job..."
               Index           =   9
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "-"
               Index           =   10
            End
            Begin VB.Menu mnuEstimateItemsAssemblySub 
               Caption         =   "Print..."
               Index           =   11
            End
         End
         Begin VB.Menu mnuEstimateItemsPO 
            Caption         =   "<POs>"
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Preview..."
               Index           =   0
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Print..."
               Index           =   1
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Send..."
               Index           =   2
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "PO Price Update Wizard..."
               Index           =   4
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "-"
               Index           =   5
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Cancel PO..."
               Index           =   6
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "-"
               Index           =   7
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Edit Vendor..."
               Index           =   8
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Change Vendor..."
               Index           =   9
            End
         End
         Begin VB.Menu mnuEstimateItemsVendors 
            Caption         =   "<Vendors>"
            Begin VB.Menu mnuEstimateItemsVendorsSub 
               Caption         =   "Edit Vendor..."
               Index           =   0
            End
         End
         Begin VB.Menu mnuEstimateItemsGrid 
            Caption         =   "<gItems>"
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Make a Copy"
               Index           =   0
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Split items..."
               Index           =   1
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Substitute Item..."
               Index           =   2
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Remove Items"
               Index           =   3
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Modify Item"
               Index           =   4
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "-"
               Index           =   5
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Attachments..."
               Index           =   6
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "-"
               Index           =   7
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Cancel Budget..."
               Index           =   8
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "-"
               Index           =   9
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Save to phase/item db..."
               Index           =   10
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Update Pricelist..."
               Index           =   11
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Compare Pricing..."
               Index           =   12
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Formatting Rules..."
               Index           =   13
            End
         End
         Begin VB.Menu mnuEstimateBidsGrid 
            Caption         =   "<gBids>"
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "Add..."
               Index           =   0
            End
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "Remove"
               Index           =   1
            End
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "-"
               Index           =   2
            End
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "Export Bid Sheet..."
               Index           =   3
            End
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "Import Bid Sheet..."
               Index           =   4
            End
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "-"
               Index           =   5
            End
            Begin VB.Menu mnuEstimateBidsGridSub 
               Caption         =   "Vendor..."
               Index           =   6
            End
         End
         Begin VB.Menu mnuEstimateRFPs 
            Caption         =   "<RFP List>"
            Begin VB.Menu mnuEstimateRFPsSub 
               Caption         =   "Delete"
               Index           =   0
            End
         End
         Begin VB.Menu mnuEstimateRFPItemsGrid 
            Caption         =   "<RFP Items>"
            Begin VB.Menu mnuEstimateRFPItemsGridSub 
               Caption         =   "Add..."
               Index           =   0
            End
            Begin VB.Menu mnuEstimateRFPItemsGridSub 
               Caption         =   "Remove"
               Index           =   1
            End
         End
         Begin VB.Menu mnuEstimateBidItemsGrid 
            Caption         =   "<gBidItems>"
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Accepted"
               Index           =   0
               Shortcut        =   ^A
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Declined"
               Index           =   1
               Shortcut        =   ^D
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Comments..."
               Index           =   2
               Shortcut        =   {F4}
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Export Bid Sheet..."
               Index           =   4
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Import Bid Sheet..."
               Index           =   5
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "-"
               Index           =   6
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Format Rules..."
               Index           =   7
            End
            Begin VB.Menu mnuEstimateBidItemsGridSub 
               Caption         =   "Vendor..."
               Index           =   8
            End
         End
         Begin VB.Menu mnuEstimateItemViews 
            Caption         =   "<views>"
            Begin VB.Menu mnuEstimateItemViewsSub 
               Caption         =   "(none)"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuPriceList 
         Caption         =   "<FPriceList>"
         Begin VB.Menu mnuPriceListViews 
            Caption         =   "<Views>"
            Begin VB.Menu mnuPriceListViewsSub 
               Caption         =   "(none)"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuSalesSheet 
         Caption         =   "<FSalesSheet>"
         Begin VB.Menu mnuSalesSheetSub 
            Caption         =   "Copy to Community..."
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSalesSheetSub 
            Caption         =   "Refresh Costs"
            Index           =   1
         End
         Begin VB.Menu mnuSalesSheetSub 
            Caption         =   "Adjust Prices..."
            Index           =   2
         End
         Begin VB.Menu mnuSalesSheetSub 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuSalesSheetSub 
            Caption         =   "Remove"
            Index           =   4
         End
      End
      Begin VB.Menu mnuAssembly 
         Caption         =   "<FAssembly>"
         Begin VB.Menu mnuAssemblyComponents 
            Caption         =   "<COMPONENTS>"
            Begin VB.Menu mnuAssemblyComponentsSub 
               Caption         =   "Add..."
               Index           =   0
            End
            Begin VB.Menu mnuAssemblyComponentsSub 
               Caption         =   "Remove"
               Index           =   1
            End
         End
         Begin VB.Menu mnuAssemblyNew 
            Caption         =   "<NEW>"
            Begin VB.Menu mnuAssemblyNewSub 
               Caption         =   "Model"
               Index           =   0
            End
            Begin VB.Menu mnuAssemblyNewSub 
               Caption         =   "Model Specific Option"
               Index           =   1
            End
            Begin VB.Menu mnuAssemblyNewSub 
               Caption         =   "Global Option"
               Index           =   2
            End
            Begin VB.Menu mnuAssemblyNewSub 
               Caption         =   "Design Center Option"
               Index           =   3
            End
            Begin VB.Menu mnuAssemblyNewSub 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnuAssemblyNewSub 
               Caption         =   "Copy of Current"
               Index           =   5
            End
         End
         Begin VB.Menu mnuAssemblyItems 
            Caption         =   "<ITEMS>"
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "Make a Copy"
               Index           =   0
            End
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "Substitute Item..."
               Index           =   1
            End
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "Remove"
               Index           =   2
            End
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "Edit Item Chart..."
               Index           =   3
            End
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "Attachments..."
               Index           =   5
            End
         End
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "<Delete>"
         Begin VB.Menu mnuDeleteSub 
            Caption         =   "Delete"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSnapShots 
         Caption         =   "<SNAPSHOTS>"
         Begin VB.Menu mnuSnapShotsSub 
            Caption         =   "Initial"
            Index           =   1
         End
         Begin VB.Menu mnuSnapShotsSub 
            Caption         =   "Confirmed"
            Index           =   2
         End
         Begin VB.Menu mnuSnapShotsSub 
            Caption         =   "Modified"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Const SRCFILE = "FMain::"
Const TASKPANELWIDTH = 3420
Private mTimerTask As String 'stupid menus
Private mCurrentJob As String

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hMod As Long


'custom descriptions
Public CD_Community As String
Public CD_Communities As String
Public CD_Series As String
Public CD_SqrFootage As String
Public CD_Color As String
 
Public Enum TakeoffSystems
    tsNone = 0
    tsOnScreen = 1
    tsPlanSwift = 2
End Enum

'command bar rows
Private InboxRow As Long
Private CustomRequestsRow As Long
Private POVendorAssignmentRow As Long




Public Property Let CurrentJob(RHS As String)
    mCurrentJob = RHS
    If IsFormLoaded("FHome2") Then Call FHome2.LoadCurJob
End Property

Public Property Get CurrentJob() As String
    CurrentJob = mCurrentJob
End Property

Public Property Get CurrentCustomer() As String
On Error Resume Next
    If mCurrentJob <> "" Then
        CurrentCustomer = "" & HFApp.SqlExec("select ar_customer_deposit from tblcustomers where job_no=" & DbQuote(Str, mCurrentJob))(0)
    End If
End Property

Public Property Get CurrentCustomerDesc() As String
On Error Resume Next
    Dim s As String
    If mCurrentJob <> "" Then
        s = ""
        s = s & "select a.description" & vbCrLf
        s = s & "from tblcustomers c join arcustomers a on c.ar_customer_deposit=a.arcustomer" & vbCrLf
        s = s & "where c.job_no=" & DbQuote(Str, mCurrentJob)
        CurrentCustomerDesc = "" & HFApp.SqlExec(s)(0)
    End If
End Property

Public Property Get CurrentJobDesc() As String
On Error Resume Next
    CurrentJobDesc = "" & HFApp.SqlExec("select description from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(Str, mCurrentJob))(0)
End Property

Private Sub MDIForm_Initialize()
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControls
End Sub


Private Sub MDIForm_Load()
On Error GoTo eh
Dim s As String

s = "init"
    mnuHidden.Enabled = InIde
    PopMenu.SubClassMenu Me
    Call IniGetForm(Me)
    Call SetToolbarIcons(Toolbar, SmallIcons)
    Toolbar.Visible = True
    
    Dim b As Boolean
    b = Not HFApp.Options.ValueByName("IsProductionDatabase")
    picWarningBanner.Visible = b
    If Not HFApp.Options.ValueByName("IsProductionVerified") Then
        lblWarningBanner.Caption = "Cannot determine if this is your live database as the license server was not reachable. Try restarting the application."
        picWarningBanner.BackColor = &H80C0FF
    End If
    
    Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomefront)(0)
    Call LoadCustomDescriptions
s = "loadinterface"

    Call LoadInterface
    
s = "read settings"

    'show taskpanel??
    CommandPanel.Visible = IniGet(AppIni, "Options", "Taskbar Displayed", "True") = "True"
    If IniGet(AppIni, "Options", "Taskbar Minimized", "False") = "True" Then
        Toolbar.Buttons("expand").Visible = True
        Toolbar.Buttons("shrink").Visible = False
        Toolbar.Move 0, 0, 360, TASKPANELWIDTH
        CommandPanel.Move 0, 0, Toolbar.Width, Me.ScaleHeight - Toolbar.Height + 4 * Screen.TwipsPerPixelY
        CommandBar.Visible = False
    Else
        Toolbar.Buttons("expand").Visible = False
        Toolbar.Buttons("shrink").Visible = True
        Toolbar.Move 0, 0, TASKPANELWIDTH, 360
        CommandPanel.Move 0, 0, TASKPANELWIDTH, Me.ScaleHeight - Toolbar.Height + 4 * Screen.TwipsPerPixelY
        CommandBar.Visible = True
    End If
    
s = "show workflow"
    'show workflow??
    If IniGet(AppIni, "Options", "HomeScreen Displayed", "True") = "True" Then
        FHome2.Show
        FHome2.WindowState = vbMaximized
    End If
    
    'show help screen?
    Select Case IniGet(AppIni, "Options", "Help Displayed")
        Case ""
        Case "Contents": Call mnuHelpSub_Click(0)
        Case "Search":   Call mnuHelpSub_Click(1)
    End Select
Exit Sub
eh: Call errHandler("MDIForm_Load", s)
End Sub

Private Sub LoadInterface()
On Error GoTo eh
    Dim dbug As String

    Dim Est As Boolean
    Dim TL As Boolean
    Dim MB As Boolean
    Dim QB As Boolean
    Dim QBO As Boolean
    Dim ITCT As Boolean
    
    Dim s As String
    Dim i As Long
    
dbug = "get integrations: "

    IsMultiFamily = HFApp.Options(MultiFamily) = "true"
    PostPOQtyToAccounting = HFApp.Options.ValueByName("PostPOQtyToAccounting") = "true"

    Est = HFApp.Databases(dbEstimating).State = adStateOpen
    TL = HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asTimberline
    MB = HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asMasterBuilder
    QB = HFApp.Options(AccountingSystem) = asQuickBooks
    QBO = HFApp.Options(AccountingSystem) = asQuickbooksOnline
    ITCT = HFApp.Options(AccountingSystem) = asIntacct
        
    
dbug = "clear bar"
    CommandBar.Rows = 0
    
        
    With CommandBar
    .Redraw = flexRDNone
        
        dbug = "add purchasing items"
        Call CommandBarAddHeader("Purchasing Tasks")
            
           
        If HFApp.Options(SalesSystem) <> SalesSystems.asNone And HFApp.UserPermission("TaskEntQuote") Then
            InboxRow = CommandBarAddItem("Inbox", "job")
            CustomRequestsRow = CommandBarAddItem("Custom Requests", "job")
        End If
        
        If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" Then
            POVendorAssignmentRow = CommandBarAddItem("TBD Assignments", "job")
        End If
        
        
        If HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440 Then
            Call CommandBarAddItem("Retrieve Jobs from Sales Center", "sendreceive")
        End If
        If HFApp.UserPermission("IssueBudgets") Then
            Call CommandBarAddItem("Prepare Job Quote", "quote")
            Call CommandBarAddItem("Issue budgets", "costcodes")
        End If
        If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
            Call CommandBarAddItem("Issue PO's", "editpos")
            Call CommandBarAddItem("Create Manual PO's", "editpos")
            Call CommandBarAddItem("Field PO Requests", "fieldpos")
        End If
        If HFApp.UserPermission("PostBudgets") Then
            If TL Then
                Call CommandBarAddItem("Post budgets", "jobcost")
            ElseIf MB Then
                Call CommandBarAddItem("Post budgets", "MB")
            ElseIf QB Then
                Call CommandBarAddItem("Post budgets", "QB")
            ElseIf QBO Then
                Call CommandBarAddItem("Post budgets", "QBO")
            ElseIf ITCT Then
                Call CommandBarAddItem("Post budgets", "Intacct")
            End If
        End If
        
        If HFApp.UserPermission("PostCommitments") Then
            If TL Then
                Call CommandBarAddItem("Post purchase orders", "jobcost")
            ElseIf MB Then
                Call CommandBarAddItem("Post purchase orders", "MB")
            ElseIf QB Then
                Call CommandBarAddItem("Post purchase orders", "QB")
            ElseIf QBO Then
                Call CommandBarAddItem("Post purchase orders", "QBO")
            ElseIf ITCT Then
                Call CommandBarAddItem("Post purchase orders", "Intacct")
            End If
        End If
        
        If HFApp.UserPermission("SendPO") And HFApp.Options.ValueByName("BuildProCompanyCode") = "" Then
            Call CommandBarAddItem("Send Purchase Orders", "sendpos")
        End If
        
        If HFApp.Options.ValueByName("PostAssembliesAsSalesInvoices") = "true" Then
        If HFApp.UserPermission("PostAREstimates") Then
            If QB Then
                Call CommandBarAddItem("Post AR Estimates", "QB")
'            ElseIf TL Then
'                Call CommandBarAddItem("Post AR Estimates", "jobcost")
'            ElseIf MB Then
'                Call CommandBarAddItem("Post AR Estimates", "MB")
            End If
        End If
        End If
        
        
        dbug = "add setup items"
        Call CommandBarAddHeader("Setup", True)
        If HFApp.UserPermission("SetupProjectManager") Then Call CommandBarAddItem("Project Managers", "newworksheet")
        If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
            If HFApp.UserPermission("SetupCommunity") Then Call CommandBarAddItem(CD_Community & " setup", "area")
        Else
            If HFApp.UserPermission("SetupCommunity") Then Call CommandBarAddItem("Work Region Setup", "area")
        End If
        If HFApp.UserPermission("SetupJobs") Then Call CommandBarAddItem("Job setup", "job")
        If HFApp.UserPermission("SetupVendors") Then Call CommandBarAddItem("Vendor setup", "vendor")
        If HFApp.UserPermission("SetupPOIndex") Then Call CommandBarAddItem("PO Indexes", "purchaseorder")
        If HFApp.UserPermission("EditItemDb") Then Call CommandBarAddItem("Import items", "ExcelImport")
        If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
            If HFApp.UserPermission("EditAssemblies") Then Call CommandBarAddItem("Import models and options", "ExcelImport")
        Else
            If HFApp.UserPermission("EditAssemblies") Then Call CommandBarAddItem("Import Assemblies", "ExcelImport")
        End If
        If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
            If HFApp.UserPermission("EditAssemblies") Then Call CommandBarAddItem("Edit Models and Options", "EditAssembly")
        Else
            If HFApp.UserPermission("EditAssemblies") Then Call CommandBarAddItem("Edit Assemblies", "EditAssembly")
        End If
        If HFApp.UserPermission("EditAssemblies") Then Call CommandBarAddItem("Option Intersections", "EditIntersection")
        If HFApp.UserPermission("EditItemDb") Then Call CommandBarAddItem("Edit item database", "ItemDB")
        
        If HFApp.UserPermission("SetupDefaultVendor") Then Call CommandBarAddItem("Default Vendors", "defaultvendors")
        
        If "True" = HFApp.Options.ValueByName("UseCommunityStandards") Then
            If HFApp.UserPermission("EditAssemblies") Then Call CommandBarAddItem(CD_Community & " Standards", "communitystandards")
        End If
        
        
        
        dbug = "add vendor items"
        If HFApp.UserPermission("SetupVendorPricing") Then
            Call CommandBarAddHeader("Vendor Pricing", True)
            Call CommandBarAddItem("Edit item prices", "pricelists")
            Call CommandBarAddItem("Cost Forecasting", "Forecast")
            Call CommandBarAddItem("Export pricelists", "ExcelExport")
            Call CommandBarAddItem("Import pricelists", "ExcelImport")
        End If
        
        
        
        dbug = "add sales prcing items"
        If HFApp.UserPermission("OpenEstimatingWorksheets") Or HFApp.UserPermission("OpenMarketingWorksheets") Then
            Call CommandBarAddHeader("Sales Pricing", True)
            If Not HFApp.Options(UseEstimatingWorksheetsOnly) And HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                Call CommandBarAddItem("Open a marketing worksheet", "worksheet")
            End If
            If HFApp.UserPermission("OpenEstimatingWorksheets") Then Call CommandBarAddItem("Open an estimating worksheet", "worksheet")
            If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                If HFApp.UserPermission("OpenEstimatingWorksheets") And SeriesItemsConfigd Then
                    Call CommandBarAddItem("Open a design center worksheet", "worksheet")
                End If
            End If
            If HFApp.UserPermission("OpenEstimatingWorksheets") Then Call CommandBarAddItem("Review a published worksheet", "preview")
        End If
    
    
    
        dbug = "done - redraw"
        Call CommandBarFormat
        .Redraw = flexRDBuffered
    End With

    
    Select Case HFApp.Options.ValueByName("TakeoffSystem")
        Case "On-Screen":   TakeoffSystem = tsOnScreen
        Case "PlanSwift":   TakeoffSystem = tsPlanSwift
        Case Else:          TakeoffSystem = tsNone
    End Select
    
    'disable planswift integration if it is not installed
    If TakeoffSystem = tsPlanSwift Then
        If Not PlanSwiftInstalled Then
            TakeoffSystem = tsNone
        End If
    End If
    
    
    
    Timer1.Interval = 30000
    Timer1.Enabled = True

    
dbug = "finished"
Exit Sub:
eh: Call errHandler(SRCFILE & "LoadInterface", dbug)
End Sub

Private Function PlanSwiftInstalled() As Boolean
    PlanSwiftInstalled = RegSectionExists(HKEY_LOCAL_MACHINE, "SOFTWARE\PlanSwift")
End Function


Private Function SeriesItemsConfigd() As Boolean
    Dim s As String
    'DesignCenter pricing worksheets only show columns for series that have a reference_item assigned.
    'If no series has a reference_item then all columns are hidden and that bleeds over into the
    'estimating pricing worksheets as well. So disable DC pricing if not configured.
    s = "select count(*) from tblseries where divisionid=" & HFApp.DivisionID & " and isnull(reference_item,'')<>''"
    SeriesItemsConfigd = Val("" & HFApp.SqlExec(s)(0)) > 0
    
End Function



Public Sub RunTask(TaskName As String, CtrlKey As Boolean)
On Error Resume Next

    Dim s As String
    Dim f As Form
    Dim i As Long
    Dim rs As Recordset
    
    Select Case TaskName
        
        Case "Rooms"
            If HFApp.UserPermission("EditAssemblies") Then
                Call FDBGrid.ShowForm("Rooms", "select RoomID,DivisionID,Room,Inactive from RoomMaster where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by room", "RoomMaster", "RoomID", , "RoomID,DivisionID", , True)
                'Call FDBGrid.ShowForm("Rooms", "select RoomID,DivisionID,Room from RoomMaster where divisionid=" & DbQuote(Num, HFApp.DivisionID) & " order by room", "RoomMaster", "RoomID", , "RoomID,DivisionID", , True)
            End If
            
        Case "Dimension Categories":                  Call FDimensionCategories.ShowForm
        Case "Model Dimensions":                      Call FModelDimensions.ShowForm
        
        Case "RFI's":                                 Call HFApp.RunTask("RFIs|" & FMain.CurrentJob)
        
        Case "Field PO Requests", "Field PO's"
            If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
                FFieldPOs.Show
            End If
        
            
        Case "Job Cost Codes":                        Call HFApp.RunTask("EditJCCostCodes")
        Case "Job Cost Categories":                   Call HFApp.RunTask("EditJCCategories")
        Case CD_Community & " Setup", "Community Setup":  If HFApp.UserPermission("SetupCommunity") Then Call HFApp.RunTask("EditCommunities")
        Case "Work Region Setup":                     If HFApp.UserPermission("SetupCommunity") Then Call HFApp.RunTask("EditCommunities")
        
        Case CD_Community & " Phases Setup":
            If HFApp.UserPermission("SetupCommunity") Then Call FDBGrid.ShowForm("Community Phases", "select Community,CommunityPhase,Description,TarionBuilderNumber from CommunityPhase order by 1,2", "CommunityPhase", "Community,CommunityPhase", , IIf(HFApp.Options.ValueByName("EnableTarionFields") = "true", "", "TarionBuilderNumber"), , True)
        
        
        Case "Default Vendors":                       If HFApp.UserPermission("SetupDefaultVendor") Then Call HFApp.RunTask("EditDefaultVendors")
        Case "Series":                                If HFApp.UserPermission("Setupseries") Then Call FDBGrid.ShowForm("Series List", "select Series,Description from tblSeries where DivisionID = " & HFApp.DivisionID, "tblSeries", "divisionid,series", HFApp.Options(SalesSystem) <> SalesSystems.asHomeFront)
        Case "Option Groups":                         If HFApp.UserPermission("SetupMajorGroup") Then Call FDBGrid.ShowForm("Option Groups", "select Major_Group,Description from tblmajorgroups", "tblmajorgroups", "major_group")
        
        
        Case "Option Categories"
            If HFApp.UserPermission("SetupCategory") Then
                Set rs = HFApp.SqlExec("select major_group from tblmajorgroups order by 1", dbHomefront)
                s = "Group"
                While Not rs.EOF
                    s = s & "|" & Trim("" & rs(0)) & "|" & Trim("" & rs(0))
                    rs.MoveNext
                Wend
                If s = "Group" Then s = ""
                Call FDBGrid.ShowForm("Option Categories", "select Group_Code, Category,Description,COMarkup from tblcategories", "tblcategories", "category")
            End If
        
        
        Case "WriteToWebAll":                        Call FWebExport.WriteToWeb("All")
        Case "WriteToWebModels":                     Call FWebExport.WriteToWeb("Models")
        
        Case "ReadFromWeb", "Retrieve Jobs from Sales Center":
            Call FInboxJobs.Show(vbModal, Me)
            
        Case "Inbox":                      If HFApp.UserPermission("IssueBudgets") Then Call FInboxJobs.Show(vbModal, Me)
        Case "Custom Requests":            If HFApp.UserPermission("IssueBudgets") Then Call FInboxCustomQuote.Show(vbModal, Me)
        Case "TBD Assignments":            If HFApp.UserPermission("IssuePOs") Then Call FInboxTBDAssignments.Show(vbModal, Me)
        
        Case "Billing":                           Set f = New FEstimateItems:  Call f.ShowForm(fceBilling, FMain.CurrentJob)
        Case "Bids":                              Set f = New FEstimateItems:  Call f.ShowForm(fceBids, FMain.CurrentJob)
        Case "Contracts":                         Set f = New FEstimateItems:  Call f.ShowForm(fceContract, FMain.CurrentJob)
        Case "Issue budgets"
            If HFApp.UserPermission("IssueBudgets") Then
                Set f = New FEstimateItems:  Call f.ShowForm(fceBudget, FMain.CurrentJob)
            End If
        Case "Create Manual PO's"
            If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
                Set f = New FPurchaseOrder:  Call f.ShowForm(fcePO, FMain.CurrentJob)
            End If
        Case "Issue PO's"
            If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
                Set f = New FEstimateItems:  Call f.ShowForm(fcePO, FMain.CurrentJob)
            End If
        Case "Prepare Job Quote":                 If HFApp.UserPermission("TaskEntQuote") Then Set f = New FEstimateItems:  Call f.ShowForm(fceQuote, "")
        '----------------------------
        Case "Project Managers":                  If HFApp.UserPermission("SetupProjectManager") Then Call HFApp.RunTask("EditProjectManagerList")
        Case "Job setup":                         If HFApp.UserPermission("SetupJobs") Then Call HFApp.RunTask("EditJob|" & FMain.CurrentJob)
        Case "Vendor setup":                      If HFApp.UserPermission("SetupVendors") Then Call HFApp.RunTask("EditVendor")
        Case "Customer setup":                    Call HFApp.RunTask("EditCustomer|" & FMain.CurrentCustomer)
        Case "Purchase Orders", "PO Indexes":     If HFApp.UserPermission("SetupPOIndex") Then Call HFApp.RunTask("EditPOIndexes")
        
        
        Case FMain.CD_Community & " Standards":               FCommunityStandards.Show vbModal
        '----------------------------
        Case "Library Mass Change":               If HFApp.UserPermission("EditAssemblies") Then Call FMassChange.ShowForm
        
        Case "Import items"
            If HFApp.UserPermission("EditItemDb") Then
                FImportAssemblies.ImportMode = "Item":     Call FImportAssemblies.Show(vbModal, Me)
            End If
        Case "Import models and options"
            If HFApp.UserPermission("EditAssemblies") Then
                FImportAssemblies.ImportMode = "Assembly": Call FImportAssemblies.Show(vbModal, Me)
            End If
        Case "Import Assemblies"
            If HFApp.UserPermission("EditAssemblies") Then
                FImportAssemblies.ImportMode = "Assembly": Call FImportAssemblies.Show(vbModal, Me)
            End If
        Case "Edit models and options":
            If HFApp.UserPermission("EditAssemblies") Then
                Set f = New FAssembly
                Call f.ShowForm(False)
            End If
        Case "Edit Assemblies":
            If HFApp.UserPermission("EditAssemblies") Then
                Set f = New FAssembly
                Call f.ShowForm(False)
            End If
            
        Case "Option Intersections":              If HFApp.UserPermission("EditAssemblies") Then Call FIntersection.Show
        Case "Edit item database":                If HFApp.UserPermission("EditItemDb") Then Call fItems.Show(vbModal)
        '----------------------------
        Case "Edit item prices":
            If HFApp.UserPermission("SetupVendorPricing") Then
                Call FPriceList.Show
                Call FPriceList.SetFocus
            End If
            
        Case "Cost Forecasting":                  If HFApp.UserPermission("SetupVendorPricing") Then Call FCostForecast.Show(vbModal)
        Case "Export pricelists":                 If HFApp.UserPermission("SetupVendorPricing") Then Call FExportPricelists.ShowForm
        Case "Import pricelists":                 If HFApp.UserPermission("SetupVendorPricing") Then Call FImportPricelists.ShowForm
        '----------------------------
        Case "Send Purchase Orders":              Call SendingWizard("PO")
        Case "Send RFQ's":                        Call SendingWizard("RFQ")
        
        Case "Review a published worksheet"
            If HFApp.UserPermission("OpenEstimatingWorkSheets") Then
                s = ""
                s = s & "SELECT Worksheet,Description,SalesEffectiveDate Posted, TStmp Modified, UStmp Author" & vbCrLf
                s = s & "FROM tblSalesSheetMaster" & vbCrLf
                s = s & "WHERE NOT SalesEffectiveDate IS NULL" & vbCrLf
'                Set f = New FRptViewer
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Pricing Worksheet", s, , False, False) Then
                    i = Val(FPickList.SelectedItem("Worksheet"))
                    s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\SalesSheetCosts.rpt")
'                    Call f.ShowReport(s, True, False, "Worksheet", i)
                    Dim c As New ZybUtil.Crystal
                    Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginID, HFApp.LoginPswd)
                    Call c.PrintPreview("Print Preview", False)
                End If

                
            End If
            
        Case "Open a marketing worksheet"
            If HFApp.UserPermission("OpenMarketingWorkSheets") Then
                Set f = New FPricingWorksheet:
                s = ""
                s = s & "Worksheets" & Chr(1)
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE isnull(m.inactive,0)=0 and SalesEffectiveDate IS NULL  AND m.WorksheetType=0 and m.DivisionID = " & HFApp.DivisionID & Chr(0)
                s = s & "Archived Sheets" & Chr(1)
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE isnull(m.inactive,0)=1 and SalesEffectiveDate IS NULL  AND m.WorksheetType=0 and m.DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Pricing Worksheet", s, , False, False) Then
                    i = Val(FPickList.SelectedItem("Worksheet"))
                    s = HFApp.RecordLockedBy("tblSalesSheetMaster", i)
                    If s = "" Or s = HFApp.LoginID Then
                        If Not ShowForm("FPricingWorksheet", i) Then Call f.OpenWorksheet(0, i, "Marketing", "", "")
                    Else
                        MsgBox "Unable to open this worksheet. It is in use by " & s, vbInformation, App.ProductName
                    End If
                End If
            End If
        
        Case "Open an estimating worksheet"
            If HFApp.UserPermission("OpenEstimatingWorkSheets") Then
                Set f = New FPricingWorksheet
                s = ""
                s = s & "Worksheets" & Chr(1)
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",m.SalesEffectiveDate Posted,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE isnull(m.inactive,0)=0 and m.WorksheetType=0 and m.DivisionID = " & DbQuote(Num, HFApp.DivisionID) & Chr(0)
                s = s & "Archived Sheets" & Chr(1)
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",m.SalesEffectiveDate Posted,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE isnull(m.inactive,0)=1 and m.WorksheetType=0 and m.DivisionID = " & DbQuote(Num, HFApp.DivisionID)
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Pricing Worksheet", s, , False, True) Then
                    i = Val(FPickList.SelectedItem("Worksheet"))
                    s = HFApp.RecordLockedBy("tblSalesSheetMaster", i)
                    If s = "" Or s = HFApp.LoginID Then
                        If Not ShowForm("FPricingWorksheet", i) Then Call f.OpenWorksheet(0, i, "Estimating", "", "")
                    Else
                        MsgBox "Unable to open this worksheet. It is in use by " & s, vbInformation, App.ProductName
                    End If
                End If
            End If
        
        Case "Open a design center worksheet"
            If HFApp.UserPermission("OpenEstimatingWorkSheets") Then
                Set f = New FPricingWorksheet
                s = ""
                s = s & "Worksheets" & Chr(1)
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",m.SalesEffectiveDate Posted,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE isnull(m.inactive,0)=0 and m.WorksheetType=1 and m.DivisionID=" & HFApp.DivisionID & Chr(0)
                s = s & "Archived Sheets" & Chr(1)
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",m.SalesEffectiveDate Posted,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE isnull(m.inactive,0)=1 and m.WorksheetType=1 and m.DivisionID=" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Pricing Worksheet", s, , False, True) Then
                    i = Val(FPickList.SelectedItem("Worksheet"))
                    s = HFApp.RecordLockedBy("tblSalesSheetMaster", i)
                    If s = "" Or s = HFApp.LoginID Then
                        If Not ShowForm("FPricingWorksheet", i) Then Call f.OpenWorksheet(1, i, "Estimating", "", "")
                    Else
                        MsgBox "Unable to open this worksheet. It is in use by " & s, vbInformation, App.ProductName
                    End If
                End If
            End If
        
        
        '----------------------------
        Case "Post budgets":                    If HFApp.UserPermission("PostBudgets") Then Call FExportBudgets.Show(vbModal)
        Case "Post purchase orders":            If HFApp.UserPermission("PostCommitments") Then Call FExportPOs.Show
                
        Case "Post AR Estimates":               If HFApp.UserPermission("PostAREstimates") Then Call FExportAREstimates.Show
        
        
        Case Else
            MsgBox TaskName & " unknown", , "FMain.RunTask()"
    End Select
End Sub





Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call IniPut(AppIni, "Options", "Taskbar Displayed", CommandPanel.Visible)
    Call IniPut(AppIni, "Options", "Taskbar Minimized", CommandPanel.Width < 361)
    Call IniPut(AppIni, "Options", "HomeScreen Displayed", IsFormLoaded("FHome2"))

On Error Resume Next



End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    CommandPanel.Move 0, 0, CommandPanel.Width, Me.ScaleHeight
    CommandBar.Move 0, Toolbar.Height, CommandPanel.Width, Me.ScaleHeight - Toolbar.Height + 4 * Screen.TwipsPerPixelY

End Sub




Private Sub mnuAssemblyComponentsSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuAssemblyComponentsSub_Click(Index)
End Sub

Private Sub mnuAssemblyItemsSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuAssemblyItemsSub_Click(Index)
End Sub

Private Sub mnuAssemblyNewSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuAssemblyNewSub_Click(Index)
End Sub

Private Sub mnuDeleteSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuDeleteSub_Click(Index)
End Sub

Private Sub mnuEstimateBidItemsGridSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuEstimateBidItemsGridSub_Click(Index)
End Sub

Private Sub mnuEstimateBidsGridSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuEstimateBidsGridSub_Click(Index)
End Sub

Private Sub mnuEstimateItemsAssemblySub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuEstimateItemsAssemblySub_Click(Index)
End Sub

Private Sub mnuEstimateItemsCustomerSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuEstimateItemsCustomerSub_Click(Index)
End Sub

Private Sub mnuEstimateItemsGridSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuEstimateItemsGridSub_Click(Index)
End Sub

Private Sub mnuEstimateItemsPOSub_Click(Index As Integer)
On Error Resume Next
    
    If Not IsIn(HFApp.Options(AccountingSystem), asTimberline, asIntacct) And mnuEstimateItemsPOSub(Index).Caption = "PO Price Update Wizard..." Then
        MsgBox "This function is not supported for your accounting system", vbInformation + vbOKOnly, "Not available"
    Else
        If HFApp.UserPermission("IssuePOs") Or (HFApp.UserPermission("GeneratePOs") And Index <> 4 And Index <> 6 And Index <> 7) Then
            Call Screen.ActiveForm.mnuEstimateItemsPOSub_Click(Index)
        Else
            MsgBox "You are not authorized to perform this function", vbInformation + vbOKOnly, "Access Denied"
        End If
    End If
End Sub


Private Sub mnuEstimateItemsVendorsSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuEstimateItemsVendorsSub_Click(Index)
End Sub

Private Sub mnuEstimateRFPItemsGridSub_Click(Index As Integer)
On Error Resume Next
'changed activeform
    Call Screen.ActiveForm.mnuEstimateRFPItemsGridSub_Click(Index)
End Sub

Private Sub mnuEstimateRFPsSub_Click(Index As Integer)
On Error Resume Next
'changed activeform
    Call Screen.ActiveForm.mnuEstimateRFPsSub_Click(Index)
End Sub





Private Sub mnuFItemsItemsSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuFItemsItemsSub_Click(Index)
End Sub

Private Sub mnuFItemsPhasesSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuFItemsPhasesSub_Click(Index)
End Sub


Private Sub mnuFItemsPriceGroupSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuFItemsPriceGroupSub_Click(Index)
End Sub





Private Sub mnuGrid2Sub_Click(Index As Integer)
    Call mnuGridSub_Click(Index)
End Sub

Private Sub mnuGridSub_Click(Index As Integer)
On Error Resume Next
    Dim i As Long
    Dim s As String
    
    Select Case Index
        
        Case mcGRID_GROUP
            MouseCtrl.ColData(MouseCol) = IIf(MouseCtrl.ColData(MouseCol) = "GROUPED", "", "GROUPED")
            Call Screen.ActiveForm.GroupGrid
            
        Case mcGRID_EXPANDALL
            MouseCtrl.Outline -1
            
        Case mcGRID_COLLAPSEALL
            MouseCtrl.Outline 6
            MouseCtrl.Outline 5
            MouseCtrl.Outline 4
            MouseCtrl.Outline 3
            MouseCtrl.Outline 2
            MouseCtrl.Outline 1
            MouseCtrl.Outline 0
        
        Case mcGRID_INSERT
                Set FColumns.Grid = MouseCtrl
                FColumns.Show vbModal
                On Error Resume Next
                Call Screen.ActiveForm.GroupGrid
        
        Case mcGRID_RENAME
            If Screen.ActiveForm Is FCostForecast Then
                Call Screen.ActiveForm.RenameForecast
            Else
                MouseCtrl.Col = MouseCol
                s = InputBox(vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Change description from """ & Trim(MouseCtrl.TextMatrix(0, MouseCol)) & """ to:", MouseCtrl.ColKey(MouseCol))
                If s <> "" Then MouseCtrl.TextMatrix(0, MouseCol) = s
            End If
            
        Case mcGRID_HIDE
            MouseCtrl.ColHidden(MouseCol) = True
            On Error Resume Next
            Call Screen.ActiveForm.RefreshGrid
            
        Case mcGRID_PRINT
            MouseCtrl.TopRow = MouseCtrl.FixedRows
            MouseCtrl.LeftCol = MouseCtrl.FixedCols
            Call MouseCtrl.Outline(-1)
            Call MouseCtrl.PrintGrid("", True, , 720, 720)
            
        Case mcGRID_SAVEAS
            If VBGetSaveFileName(s, , , "Excel (*.xls)|*.xls|Text (*.txt)|*.txt|Comma Separated (*.csv)|*.csv", , , , "txt", FMain.hwnd) Then
                Call MouseCtrl.Outline(-1)
                Select Case UCase(FileExt(s))
                    Case "XLS":  Call MouseCtrl.SaveGrid(s, flexFileExcel, True)
                    Case "CSV":  Call MouseCtrl.SaveGrid(s, flexFileCommaText, True)
                    Case Else:   Call MouseCtrl.SaveGrid(s, flexFileTabText, True)
                End Select
            End If
            
        Case mcGRID_ATTACHMENTS
        
            s = "EditAttachments"
            s = s & "|" & MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("AttachmentID")) & "|Customer Option Attachments"
            s = s & "|J~" & Replace(MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("Job_No")), "-", "") & "| Job Attachments"
            s = s & "|C~" & MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("Customer_No")) & "|Customer File Attachments"
            
            mTimerTask = s
            Timer1.Enabled = False
            Timer1.Interval = 10
            Timer1.Enabled = True
            
    End Select
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next

    FreeLibrary m_hMod
    Call SetErrorMode(SEM_NOERRORS)
    Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
    
    Call IniPutForm(Me)
    
    'get rid of any failed report windows
    While Forms.Count > 1
        Unload Forms(1)
    Wend

    Call HFApp.Logout
    
    Set HFApp = Nothing
    Set FMain = Nothing
    
End Sub

Private Sub mnuInvoicesSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuInvoicesSub_Click(Index)
End Sub



Private Sub mnuSnapShotsSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuSnapShotsSub_Click(Index)
End Sub

Private Sub mnuTakeOffJobView_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeOffJobView_Click(Index)
End Sub



Private Sub mnuWindowSub_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case mcVIEW_WORKFLOW
            If IsFormLoaded("FHome2") Then
                Unload FHome2
            Else
                'If IniGet(AppIni, "Options", "HomeScreen Displayed", "True") = "True" Then
                    FHome2.Show
                    FHome2.WindowState = vbMaximized
                'End If
            End If
            
        Case mcVIEW_TASKPANEL
            If CommandPanel.Visible Then
                CommandPanel.Visible = False
            Else
                Call LoadInterface
                Toolbar.Buttons("expand").Visible = False
                Toolbar.Buttons("shrink").Visible = True
                CommandPanel.Move 0, 0, TASKPANELWIDTH, Me.ScaleHeight
                Toolbar.Move 0, 0, 3195, 360
                CommandBar.Visible = True
                CommandPanel.Visible = True
                Call MDIForm_Resize
            End If
    End Select
End Sub

Private Sub mnuFileSub_Click(Index As Integer)
On Error Resume Next
Dim s As String
    Select Case Index
    'changed activeform
        Case mcFILE_SAVE:  Call Screen.ActiveForm.SaveData(False)
        Case mcFILE_CHANGEDIV
            Dim divID As String
            s = "select  d.DivisionCode,d.DivisionID, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & HFApp.LoginID & " or u.userid is null"
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Select Division", s, "", True, False, , "DivisionID") Then
                
                Call HFApp.SetDivision(FPickList.SelectedItem("DivisionID"))
                Call HFApp.Options.ReadData
                If HFApp.DivisionID <> "" Then
                    Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomefront)(0)
                Else
                    Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor
                End If
            End If
        Case mcFILE_CLOSE: Unload Me
    End Select
End Sub



Public Sub ShowColumnMenu(Grid, Optional Hideable As Boolean = True, _
                                Optional Showable As Boolean = True, _
                                Optional Renameable As Boolean = True, _
                                Optional Groupable As Boolean = False, _
                                Optional Attachments As Boolean = True)
On Error GoTo eh

    Dim ParentMenu As Long
    Dim i As Long
    Dim j As Long
    Dim columns() As String

    With PopMenu
        If Groupable Then
            'save this stuff for menu click
            Set MouseCtrl = Grid
            MouseCol = Grid.MouseCol
        
            'set these
            mnuGrid2Sub(mcGRID_GROUP).checked = MouseCtrl.ColData(MouseCol) = "GROUPED"
            mnuGrid2Sub(mcGRID_RENAME).Enabled = Renameable
            mnuGrid2Sub(mcGRID_HIDE).Enabled = MouseCol >= 0 And Hideable
            mnuGrid2Sub(mcGRID_INSERT).Enabled = MouseCol >= 0 And Showable
            
'            'first unload columns
'don't load any columns, just show the "more columns" window
'            Call .ClearSubMenusOfItem(.MenuIndex("mnuGrid2Sub(5)"))
'            Call .ClearSubMenusOfItem(.MenuIndex("mnuGridSub(5)"))
'            ParentMenu = .MenuIndex("mnuGrid2Sub(5)")
'
'            'load/sort column names/keys
'            For i = 0 To MouseCtrl.Cols - 1
'                If MouseCtrl.TextMatrix(0, i) <> "" Then
'                    j = j + 1
'                    ReDim Preserve columns(j)
'                    columns(j) = MouseCtrl.TextMatrix(0, i) & Chr(1) & MouseCtrl.ColKey(i)
'                End If
'            Next
'            Call Sort(columns)
'            For i = 1 To UBound(columns)
'                .AddItem Parse(columns(i), 1, Chr(1)), "ShowColumn" & i, , , ParentMenu, , Not MouseCtrl.ColHidden(MouseCtrl.ColIndex(Parse(columns(i), 2, Chr(1))))
'                .MenuTag("ShowColumn" & i) = Parse(columns(i), 2, Chr(1))
'            Next
'
'
'            If j = 0 Then .AddItem "(none available)", , , , ParentMenu, , , False
'            .AddItem "-", , , , ParentMenu
'            .AddItem "More...", "ShowMoreColumns", , , ParentMenu
            
            
            
            'show menu
            PopupMenu mnuGrid2
        Else
            'save this stuff for menu click
            Set MouseCtrl = Grid
            MouseCol = Grid.MouseCol
        
            'set these
            mnuGridSub(mcGRID_ATTACHMENTS).Visible = Attachments
            mnuGridSub(mcGRID_RENAME).Enabled = Renameable
            mnuGridSub(mcGRID_HIDE).Enabled = MouseCol >= 0 And Hideable
            mnuGridSub(mcGRID_INSERT).Enabled = MouseCol >= 0 And Showable
            
            
'            'first unload columns
'don't load any columns, just show the "more columns" window
'            Call .ClearSubMenusOfItem(.MenuIndex("mnuGrid2Sub(5)"))
'            Call .ClearSubMenusOfItem(.MenuIndex("mnuGridSub(5)"))
'            ParentMenu = .MenuIndex("mnuGridSub(5)")
'
'            'load/sort column names/keys
'            For i = 0 To MouseCtrl.Cols - 1
'                If MouseCtrl.TextMatrix(0, i) <> "" Then
'                    j = j + 1
'                    ReDim Preserve columns(j)
'                    columns(j) = MouseCtrl.TextMatrix(0, i) & Chr(1) & MouseCtrl.ColKey(i)
'                End If
'            Next
'            Call Sort(columns)
'            For i = 1 To UBound(columns)
'                .AddItem Parse(columns(i), 1, Chr(1)), "ShowColumn" & i, , , ParentMenu, , Not MouseCtrl.ColHidden(MouseCtrl.ColIndex(Parse(columns(i), 2, Chr(1))))
'                .MenuTag("ShowColumn" & i) = Parse(columns(i), 2, Chr(1))
'            Next
'
'            If j = 0 Then .AddItem "(none available)", , , , ParentMenu, , , False
'            .AddItem "-", , , , ParentMenu
'            .AddItem "More...", "ShowMoreColumns", , , ParentMenu
            
            'show menu
            PopupMenu mnuGrid
        End If
    End With


    Exit Sub
eh: Call errHandler(SRCFILE & "ShowColumnMenu")
End Sub

Private Sub mnuHelpSub_Click(Index As Integer)
    Select Case Index
        Case 0 'contents
        Case 1 'search
        Case 3 'Help
            Call ShellFile(Me.hwnd, "https://usergroup.hyphensolutions.com/HomeFront", , False)
        
        Case 4 'Support
            Call ShellFile(Me.hwnd, "https://www.hyphensolutions.com/info/contact/", , False)

        Case 6 'about
            Call HFApp.About
    End Select
End Sub

Private Sub mnuSalesSheetSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuSalesSheetSub_Click(Index)
End Sub

Private Sub mnuTakeOffAssemblyView_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeOffAssemblyView_Click(Index)
End Sub

Private Sub mnuTakeOffEstimateView_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeOffEstimateView_Click(Index)
End Sub

Private Sub mnuTakeOffItemView_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeOffItemView_Click(Index)
End Sub

Private Sub mnuTakeOffModelsView_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeOffModelsView_Click(Index)
End Sub

Private Sub mnuTakeOffSub_Click(Index As Integer)
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeOffSub_Click(Index)
End Sub

Private Sub mnuTakeoffCustomOptions_Click()
On Error Resume Next
    Call Screen.ActiveForm.mnuTakeoffCustomOptions_Click
End Sub

Private Sub PopMenu_InitPopupMenu(ParentItemNumber As Long)
    Dim i As Long
    Dim j As Long
    Dim sFile   As String
    Dim FInfo   As New ClsFileInfo
    Dim sFolder As String
    Dim def As String
    Dim c  As New Connection
    Dim rs As Recordset
    Dim s  As String
    Dim b  As Boolean
    
    With PopMenu
        .ImageList = SmallIcons
        
        Select Case True
            
            Case "mnuFile" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                '.AddItem "Import Data File", "import", , , ParentItemNumber, ImageIndex(SmallIcons, "folder")
                .AddItem "Save", "save", , , ParentItemNumber, ImageIndex(SmallIcons, "Save")
                .AddItem "-", , , , ParentItemNumber
                .AddItem "Change Division", "changediv", , , ParentItemNumber
                .AddItem "Close", "close", , , ParentItemNumber
                
            Case "mnuConstruction" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                If HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440 Then
                    .AddItem "Retrieve Jobs from Sales Center", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "sendreceive"), , HFApp.UserPermission("IssueBudgets")
                End If
                If HFApp.Options(SalesSystem) <> SalesSystems.asNone Then
                    .AddItem "Inbox", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "job"), , HFApp.UserPermission("IssueBudgets")
                    .AddItem "Custom Requests", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "Custom Requests"), , HFApp.UserPermission("IssueBudgets")
                End If
                .AddItem "Issue budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes"), , HFApp.UserPermission("IssueBudgets")
                .AddItem "Issue PO's", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "editpos"), , (HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs"))
                .AddItem "Create Manual PO's", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "editpos"), , (HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs"))
                .AddItem "Field PO Requests", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "fieldpos"), , (HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs"))
                
                If HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asTimberline Then
                    .AddItem "Post budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "jobcost"), , HFApp.UserPermission("PostBudgets")
                    .AddItem "Post purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "jobcost"), , HFApp.UserPermission("PostCommitments")
                ElseIf HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asMasterBuilder Then
                    .AddItem "Post budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "MB"), , HFApp.UserPermission("PostBudgets")
                    .AddItem "Post purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "MB"), , HFApp.UserPermission("PostCommitments")
                ElseIf HFApp.Options(AccountingSystem) = asQuickBooks Then
                    .AddItem "Post budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "QB"), , HFApp.UserPermission("PostBudgets")
                    .AddItem "Post purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "QB"), , HFApp.UserPermission("PostCommitments")
                End If
                .AddItem "Send purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "sendpos"), , HFApp.UserPermission("SendPO")
                .AddItem "Post AR Estimates", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "sendpos"), , HFApp.UserPermission("PostAREstimates")
                .AddItem "Edit item prices", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "pricelists"), , HFApp.UserPermission("SetupVendorPricing")
                .AddItem "Cost Forecasting", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "Forecast"), , HFApp.UserPermission("SetupVendorPricing")
                            
            
            Case "mnuSalesPricing" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                If Not HFApp.Options(UseEstimatingWorksheetsOnly) And HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    .AddItem "Open a marketing worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "worksheet"), , HFApp.UserPermission("OpenMarketingWorkSheets")
                End If
                .AddItem "Open an estimating worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "worksheet"), , HFApp.UserPermission("OpenEstimatingWorkSheets")
                If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    If HFApp.UserPermission("OpenEstimatingWorksheets") And SeriesItemsConfigd Then
                        .AddItem "Open a design center worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "worksheet"), , HFApp.UserPermission("OpenEstimatingWorkSheets")
                    End If
                End If
                .AddItem "Review a published worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "preview"), , HFApp.UserPermission("OpenEstimatingWorkSheets")
            
            Case "mnuSetup" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                .AddItem "Job Cost Codes", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes")
                .AddItem "Job Cost Categories", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes")
                
                If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    .AddItem "Rooms", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes"), , HFApp.UserPermission("EditAssemblies")
                    .AddItem "Dimension Categories", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes"), , HFApp.UserPermission("EditAssemblies")
                    .AddItem "Model Dimensions", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes"), , HFApp.UserPermission("EditAssemblies")
                End If
            
                .AddItem "Option Attributes", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes")
                
                If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                    .AddItem "Work Region Setup", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "area"), , HFApp.UserPermission("SetupCommunity")
                Else
                    .AddItem CD_Community & " Setup", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "area"), , HFApp.UserPermission("SetupCommunity")
                End If
                .AddItem CD_Community & " Phases Setup", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "area"), , HFApp.UserPermission("SetupCommunity")
                .AddItem "Project Managers", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "newworksheet"), , HFApp.UserPermission("SetupProjectManager")
                
                .AddItem "Vendor setup", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "vendor"), , HFApp.UserPermission("SetupVendors")
                .AddItem "Customer setup", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ExcelImport"), , HFApp.UserPermission("SetupCustomer")
                .AddItem "Job setup", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "job"), , HFApp.UserPermission("SetupJobs")
                .AddItem "PO Indexes", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "purchaseorder"), , HFApp.UserPermission("SetupPOIndex")
                .AddItem "Default Vendors", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "defaultvendors"), , HFApp.UserPermission("SetupDefaultVendor")
                .AddItem "Edit item database", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ItemDB"), , HFApp.UserPermission("EditItemDb")
                If "True" = HFApp.Options.ValueByName("UseCommunityStandards") Then
                    .AddItem FMain.CD_Community & " Standards", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "communitystandards")
                End If
                If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                    .AddItem "Edit Assemblies", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "EditAssembly"), , HFApp.UserPermission("EditAssemblies")
                Else
                    .AddItem "Edit Models and Options", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "EditAssembly"), , HFApp.UserPermission("EditAssemblies")
                End If
            
            
            Case "mnuInquiries" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                sFolder = PathAppend(HFApp.SystemFolder, "Estimating\Inquiries", Mid(.HierarchyPath(ParentItemNumber, 1, "\"), 10))
                sFile = Dir(PathAppend(sFolder, "*.*"), vbDirectory)
                While sFile <> ""
                    If sFile <> "." And sFile <> ".." And sFile <> "msdir" Then
                        FInfo.FullPathName = PathAppend(sFolder, sFile)
                        If FInfo.hSmlIList > 0 Then PopMenu.ImageList = FInfo.hSmlIList
                        If FInfo.TypeName = "File Folder" Then
                            Call .AddItem("(dummy)", , , , .AddItem(sFile, , , , ParentItemNumber, FInfo.hSmlIcon))
                        Else
                            Call .AddItem(StripExtension(sFile), PathAppend(sFolder, sFile), , , ParentItemNumber, FInfo.hSmlIcon)
                        End If
                    End If
                    sFile = Dir
                Wend
                If .firstChild(ParentItemNumber) = 0 Then
                    .AddItem "(none found)", , , , ParentItemNumber, , , False
                End If
                If .MenuKey(ParentItemNumber) = "mnuInquiries" Then
                    .AddItem "-", , , , ParentItemNumber
                    .AddItem "&Manage", PathAppend(HFApp.SystemFolder, "Estimating\Inquiries"), , , ParentItemNumber
                End If
            
            Case "mnuReports" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                sFolder = PathAppend(HFApp.SystemFolder, "Estimating\Reports", Mid(.HierarchyPath(ParentItemNumber, 1, "\"), 8))
                sFile = Dir(PathAppend(sFolder, "*.*"), vbDirectory)
                While sFile <> ""
                    If sFile <> "." And sFile <> ".." And sFile <> "msdir" Then
                        FInfo.FullPathName = PathAppend(sFolder, sFile)
                        If FInfo.hSmlIList > 0 Then PopMenu.ImageList = FInfo.hSmlIList
                        If FInfo.TypeName = "File Folder" Then
                            Call .AddItem("(dummy)", , , , .AddItem(sFile, , , , ParentItemNumber, FInfo.hSmlIcon))
                        Else
                            Call .AddItem(StripExtension(sFile), PathAppend(sFolder, sFile), , , ParentItemNumber, FInfo.hSmlIcon)
                        End If
                    End If
                    sFile = Dir
                Wend
                If .firstChild(ParentItemNumber) = 0 Then
                    .AddItem "(none found)", , , , ParentItemNumber, , , False
                End If
                If .MenuKey(ParentItemNumber) = "mnuReports" Then
                    .AddItem "-", , , , ParentItemNumber
                    .AddItem "&Manage", PathAppend(HFApp.SystemFolder, "Estimating\Reports"), , , ParentItemNumber
                End If
                
            Case "Sync" = .MenuKey(ParentItemNumber)
                Call .ClearSubMenusOfItem(ParentItemNumber)
                b = False
                
                If HFApp.Options(SalesSystem) = asBuilder1440 Then
                    If b Then .AddItem "-", , , , ParentItemNumber
                    b = True
                    
                    'If HFApp.Options.ValueByName("SalesSimplicity") = "True" Then
                    '    .AddItem "Upload Models && Options to Sales Simplicity", "WriteWeb", , , ParentItemNumber
                    '    .AddItem "Read Customers from Sales Simplicity", "ReadWeb", , , ParentItemNumber
                    'Else
                        .AddItem "Upload Models && Options to Builder 1440", "WriteWebAll", , , ParentItemNumber
                        .AddItem "Upload Models to Builder 1440", "WriteWebModels", , , ParentItemNumber
                        .AddItem "Read Customers from Builder 1440", "ReadWeb", , , ParentItemNumber
                    'End If
                End If
                
                Select Case HFApp.Options(AccountingSystem)
                    Case asIntacct
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", SyncCmd("ReadAllARCustomers"), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", SyncCmd("ReadAllStandardCostCodes"), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", SyncCmd("ReadAllGLAccounts"), , , ParentItemNumber
                        
                    Case asTimberline, asMasterBuilder, asSpectrum
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", SyncCmd("ReadAllARCustomers"), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", SyncCmd("ReadAllStandardCostCodes"), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", SyncCmd("ReadTaxGroups"), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", SyncCmd("ReadAllGLAccounts"), , , ParentItemNumber
                    
                    Case asQuickBooks
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", SyncCmd("ReadAllARCustomers"), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", SyncCmd("ReadAllStandardCostCodes"), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", SyncCmd("ReadTaxGroups"), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", SyncCmd("ReadAllGLAccounts"), , , ParentItemNumber
                        .AddItem "-", , , , ParentItemNumber
                        .AddItem "Import Payroll Costs from Accounting", SyncCmd("ReadPayrollCosts"), , , ParentItemNumber
                    
                    Case asQuickbooksOnline
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        '.AddItem "Read Accounts Receivable from Accounting", SyncCmd("ReadAllARCustomers"), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", SyncCmd("ReadAllStandardCostCodes"), , , ParentItemNumber
                        '.AddItem "Read Tax Settings from Accounting", SyncCmd("ReadTaxGroups"), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", SyncCmd("ReadAllGLAccounts"), , , ParentItemNumber
                        '.AddItem "-", , , , ParentItemNumber
                        '.AddItem "Import Payroll Costs from Accounting", SyncCmd("ReadPayrollCosts"), , , ParentItemNumber
                    
                    Case asSimply
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", SyncCmd("ReadAllARCustomers"), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", SyncCmd("ReadTaxGroups"), , , ParentItemNumber
                    
                    Case asMYOB
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", SyncCmd("ReadTaxGroups"), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", SyncCmd("ReadAllGLAccounts"), , , ParentItemNumber
                    
                    Case asXero
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Vendors from Accounting", SyncCmd("ReadAllVendors"), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", SyncCmd("ReadAllStandardCostCodes"), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", SyncCmd("ReadTaxGroups"), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", SyncCmd("ReadAllGLAccounts"), , , ParentItemNumber
                     
                    Case Else
                End Select
                
                If b Then .AddItem "-", , , , ParentItemNumber
                b = True
                .AddItem "Read New records from Accounting", SyncCmd("Starthfest"), , , ParentItemNumber
                
                
            Case "mnuTools" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                
                If HFApp.Options.ValueByName("EstimatingSystem") = esPipeline Then
                    Call .AddItem("Import BIM Pipeline Models", , , , ParentItemNumber)
                    Call .AddItem("Import BIM Pipeline Global Options", , , , ParentItemNumber)
'old excel import     Call .AddItem("Import BIM Pipeline Assemblies", , , , ParentItemNumber)
                End If
                
                Call .AddItem("Cancel Purchase Order Wizard", , , , ParentItemNumber)
                
                
                If IsIn(HFApp.Options(AccountingSystem), asIntacct, asTimberline) Then
                    Call .AddItem("Purchase Order Price Update Wizard", , , , ParentItemNumber)
                End If
                
                Call .AddItem("Mass Change Wizard", , , , ParentItemNumber)
                Call .AddItem("Duplicate Model Options", , , , ParentItemNumber)
                Call .AddItem("Assign TBD Purchase Orders", , , , ParentItemNumber)
                Call .AddItem("Work Reassignment Wizard", , , , ParentItemNumber)
                Call .AddItem("Reset Default Vendors...", , , , ParentItemNumber)
                
                
                
                Call .AddItem("Division Replication Wizard", , , , ParentItemNumber)
                
                .AddItem "-", , , , ParentItemNumber
                Call .AddItem("Reset grid layouts to default...", , , , ParentItemNumber)
                    
                If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" Then
                    .AddItem "-", , , , ParentItemNumber
                    Call .AddItem("Send POIndexes to BuildPro", , , , ParentItemNumber)
                    Call .AddItem("Send Vendors to BuildPro", , , , ParentItemNumber)
                    Call .AddItem("Send Communities to BuildPro", , , , ParentItemNumber)
                    Call .AddItem("Send Jobs to BuildPro", , , , ParentItemNumber)
                    Call .AddItem("Send POs to BuildPro", , , , ParentItemNumber)
                    Call .AddItem("Send Vendor Payments To BuildPro", , , , ParentItemNumber)
                End If
                
                .AddItem "-", , , , ParentItemNumber
                .AddItem "Import items", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ExcelImport"), , HFApp.UserPermission("EditItemDb")
                If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                    .AddItem "Import Assemblies", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ExcelImport"), , HFApp.UserPermission("EditAssemblies")
                Else
                    .AddItem "Import models and options", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ExcelImport"), , HFApp.UserPermission("EditAssemblies")
                End If
                .AddItem "Import pricelists", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ExcelImport"), , HFApp.UserPermission("SetupVendorPricing")
                .AddItem "Export pricelists", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ExcelExport"), , HFApp.UserPermission("SetupVendorPricing")
                .AddItem "-", , , , ParentItemNumber
                sFolder = PathAppend(HFApp.SystemFolder, "Estimating\Tools", Mid(.HierarchyPath(ParentItemNumber, 1, "\"), 6))
                sFile = Dir(PathAppend(sFolder, "*.*"), vbDirectory)
                While sFile <> ""
                    If sFile <> "." And sFile <> ".." And sFile <> "msdir" Then
                        FInfo.FullPathName = PathAppend(sFolder, sFile)
                        If FInfo.hSmlIList > 0 Then PopMenu.ImageList = FInfo.hSmlIList
                        If FInfo.TypeName = "File Folder" Then
                            Call .AddItem("(dummy)", , , , .AddItem(sFile, , , , ParentItemNumber, FInfo.hSmlIcon))
                        Else
                            Call .AddItem(StripExtension(sFile), sFolder & "\" & sFile, , , ParentItemNumber, FInfo.hSmlIcon)
                        End If
                    End If
                    sFile = Dir
                Wend
                
                If .MenuKey(ParentItemNumber) = "mnuTools" Then
                    .AddItem "User Permissions...", , "task", , ParentItemNumber, , , HFApp.UserPermission("ToolsUsrAdmin")
                    .AddItem "System Settings...", "AppOptions", , , ParentItemNumber, , , HFApp.UserPermission("ToolsSysAdmin")
                    If HFApp.Databases(dbAccounting).State = adStateOpen Or HFApp.Options(SalesSystem) = asBuilder1440 Or HFApp.Options(AccountingSystem) = asQuickBooks Or HFApp.Options(AccountingSystem) = asSimply Or HFApp.Options(AccountingSystem) = asMYOB Or HFApp.Options(AccountingSystem) = asPeachtree Or HFApp.Options(AccountingSystem) = asXero Or HFApp.Options(AccountingSystem) = asIntacct Or HFApp.Options(AccountingSystem) = asQuickbooksOnline Then
                        Call .AddItem("(dummy)", , , , .AddItem("Database Synchronization", "Sync", , , ParentItemNumber))
                    End If
                    .AddItem "-", , , , ParentItemNumber
                    .AddItem "Additional Tools...", PathAppend(HFApp.SystemFolder, "Estimating\Tools"), , , ParentItemNumber
                End If
                
         
                
            Case Else
                
         End Select
        
    End With
End Sub




Private Sub PopMenu_Click(ItemNumber As Long)
On Error GoTo eh
    Dim s As String
    
    Dim Job As String
    Dim PricingCommunity As String
    Dim community As String
    Dim Model As String
    Dim OptionID As String
    Dim Assembly As String
    
    Dim f As Form
    Dim b As Boolean
    
    With PopMenu
        Select Case True
            
            
            Case .Caption(ItemNumber) = "Reset grid layouts to default...":
                If MsgBox("This will reset the columns on all your screens back to their default positions. Please close all screens before clicking OK.", vbQuestion + vbOKCancel, "Restore Defaults?") = vbCancel Then Exit Sub
                HFApp.SqlExec "ZYB_ResetUserGridLayout " & DbQuote(Str, HFApp.LoginID)

            
            Case .Caption(ItemNumber) = "Option Attributes":                  Call HFApp.RunTask("EditAttributeLists")

            Case .Caption(ItemNumber) = "Division Replication Wizard":       Call FAssemblyReplicator.Show(vbModal)
            
            
            Case .Caption(ItemNumber) = "Duplicate Model Options":            Call DuplicateModelOptions
            Case .Caption(ItemNumber) = "Import BIM Pipeline Models":         Call ImportPipelineModelsAndOptions
            Case .Caption(ItemNumber) = "Import BIM Pipeline Global Options": Call ImportPipelineGlobalOptions
                
            Case .Caption(ItemNumber) = "Reset Default Vendors...":        Call ResetDefaultVendors
            Case .Caption(ItemNumber) = "Assign TBD Purchase Orders":      Call FTBDAssignmentWizard.ShowForm
            Case .Caption(ItemNumber) = "Mass Change Wizard":              Call FMassChange.ShowForm
            Case .Caption(ItemNumber) = "Cancel Purchase Order Wizard":    Call FPOMassCancel.ShowForm
            Case .Caption(ItemNumber) = "Purchase Order Price Update Wizard":    Call FPOPriceChangeWiz.ShowForm
            Case .Caption(ItemNumber) = "Work Reassignment Wizard":        Call FVendorChange.ShowForm
            Case .Caption(ItemNumber) = "User Permissions...":             Call HFApp.RunTask("EditSecurity"):       Call LoadInterface

    
            Case .Caption(ItemNumber) = "Send POIndexes to BuildPro":       Call SendBuildProPOIndexes:
            Case .Caption(ItemNumber) = "Send Vendors to BuildPro":         Call SendBuildProVendors:
            Case .Caption(ItemNumber) = "Send Communities to BuildPro":     Call SendBuildProCommunities:
            Case .Caption(ItemNumber) = "Send Jobs to BuildPro":            Call SendBuildProJobs(""):
            Case .Caption(ItemNumber) = "Send POs to BuildPro":             Call SendBuildProPOs("", ""):
            Case .Caption(ItemNumber) = "Send Vendor Payments to BuildPro": Call SendBuildProRemittances:
            
            
            Case .Caption(ItemNumber) = "Change Division"
                Dim divID As String
                s = "select  d.DivisionID,d.DivisionCode, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & DbQuote(Str, HFApp.LoginID) & " or u.userid is null order by 2"
                If FPickList.Choose(HFApp.Databases(dbHomefront), "Select Division", s, HFApp.DivisionID, True, False, , "DivisionID") Then
                    Call HFApp.SetDivision(FPickList.SelectedItem("DivisionID"))
                    Call HFApp.Options.SaveData
                    Call HFApp.Options.ReadData
                    If HFApp.DivisionID <> "" Then
                        Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomefront)(0)
                    Else
                        Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor
                    End If
                    If IsFormLoaded("FHome2") Then Call FHome2.LoadJobs
                End If
            
            Case .Caption(ItemNumber) = "Close":                           Unload Me
            'changed activeform
            Case .Caption(ItemNumber) = "Save":                            On Error Resume Next: Call Screen.ActiveForm.SaveData(False)
            Case .HelpText(ItemNumber) = "task":                           Call RunTask(.Caption(ItemNumber), False)
            
            Case Left(.MenuKey(ItemNumber), 11) = "WriteWebAll":           Call RunTask("WriteToWebAll", False)
            Case Left(.MenuKey(ItemNumber), 14) = "WriteWebModels":        Call RunTask("WriteToWebModels", False)
            
            Case Left(.MenuKey(ItemNumber), 7) = "ReadWeb":                Call RunTask("ReadFromWeb", False)
            Case Left(.MenuKey(ItemNumber), 17) = "EstimateItemViews":     Call Screen.ActiveForm.mnuEstimateItemViewsSub_Click(.ItemData(ItemNumber) + 1)
            Case Left(.MenuKey(ItemNumber), 13) = "PriceListView":         Call Screen.ActiveForm.mnuPriceListViewsSub_Click(.ItemData(ItemNumber))
            
            Case Left(.MenuKey(ItemNumber), 10) = "ShowColumn"
                MouseCtrl.ColHidden(MouseCtrl.ColIndex(.MenuTag(ItemNumber))) = False
                MouseCtrl.ColPosition(MouseCtrl.ColIndex(.MenuTag(ItemNumber))) = Max(MouseCol - 1, 0)
                On Error Resume Next
                Call Screen.ActiveForm.GroupGrid
            
            Case .MenuKey(ItemNumber) = "ShowMoreColumns"
                Set FColumns.Grid = MouseCtrl
                FColumns.Show vbModal
                On Error Resume Next
                Call Screen.ActiveForm.GroupGrid
                
            
            Case "&Manage" = .Caption(ItemNumber) Or "Additional Tools" = .Caption(ItemNumber)
                s = .MenuKey(ItemNumber)
                If Not PathExists(s) Then Call CreatePath("the folder", s)
                Call ShellFile(Me.hwnd, s)
            
                
            Case "AppOptions" = .MenuKey(ItemNumber)
                Call HFApp.RunTask("EditOptions")
                Call LoadInterface
                
                
                If IsFormLoaded("FHome2") Then FHome2.ConfigForm
                On Error Resume Next
                For Each f In Forms
                    Call f.ColorizeItems(-1)
                Next
                
            Case "mnuHelp" = .MenuKey(.UltimateParent(ItemNumber))
            
            Case "mnuTools" = .MenuKey(.UltimateParent(ItemNumber))
                s = .MenuKey(ItemNumber)
                If InStr(1, s, ".exe ", vbTextCompare) Then
                    Call Shell(s, vbNormalFocus)
                Else
                    If Not PathExists(s) Then Call CreatePath("the folder", s)
                    Call ShellFile(Me.hwnd, s)
                End If
            
            Case "mnuInquiries" = .MenuKey(.UltimateParent(ItemNumber))
                Set f = New FInquiry
                Call f.ShowInquiry(.MenuKey(ItemNumber))
                
            
            Case "mnuReports" = .MenuKey(.UltimateParent(ItemNumber))
                s = .MenuKey(ItemNumber)
                
                On Error Resume Next
                Job = Screen.ActiveForm.Job
                If Job = "" Then Job = FMain.CurrentJob
                
                PricingCommunity = Screen.ActiveForm.PricingCommunity
                community = Screen.ActiveForm.community
                Model = Screen.ActiveForm.Model
                OptionID = Screen.ActiveForm.OptionID
                Assembly = Screen.ActiveForm.Assembly
                On Error GoTo eh
                
                'this is stupid but required. Crystal shits the bed sometimes and this is how we prevent it.
                If Job = "" Then Job = ""
                If PricingCommunity = "" Then PricingCommunity = ""
                If community = "" Then community = ""
                If Model = "" Then Model = ""
                If OptionID = "" Then OptionID = ""
                If Assembly = "" Then Assembly = ""
                
                
                If FileExt(s) = "rpt" Then
                    Dim c As New ZybUtil.Crystal
                    Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
                    
                    On Error Resume Next
                    Call c.ParameterValue("DivisionID", HFApp.DivisionID)
                    If Job <> "" Then Call c.ParameterValue("Job", Job)
                    If community <> "" Then Call c.ParameterValue("Community", community)
                    If Assembly <> "" Then Call c.ParameterValue("Assembly", Assembly)
                    If Model <> "" Then Call c.ParameterValue("Model", Model)
                    If OptionID <> "" Then Call c.ParameterValue("OptionID", OptionID)
                    If PricingCommunity <> "" Then Call c.ParameterValue("PricingCommunity", PricingCommunity)
                    On Error GoTo eh
                    
                    Call c.PrintPreview("Print Preview", False)

                Else
                    Call ShellFile(Me.hwnd, s)
                End If
                
                
        End Select
    End With



                    


Exit Sub
eh: Call errHandler(SRCFILE & "PopMenu_Click")
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
    Dim d As Double
    Dim i As Long
    Dim s As String
    Dim s2 As String
    
    
    If Mid(mTimerTask, 1, 15) = "EditAttachments" Then
         Call HFApp.RunTask(mTimerTask)
         mTimerTask = ""
    End If
    
    Timer1.Enabled = False
    Timer1.Interval = 30000
    Timer1.Enabled = True
    i = 0
    
    If HFApp.Options.ValueByName("DisableAutoUpdate") <> "True" Then
        If InboxRow <> 0 Then
            i = 0
            i = i + Val("" & HFApp.SqlExec("select count(*) from unestimatedoptions o  join divisioncommunities d on d.Community=o.Community where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")", dbHomefront)(0))
            i = i + Val("" & HFApp.SqlExec("select count(*) from unestimatedcustomers c  join divisioncommunities d on d.Community=c.Community where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")", dbHomefront)(0))
            CommandBar.TextMatrix(InboxRow, 2) = "Inbox (" & IIf(i = 0, "none", i) & ")"
            CommandBar.Cell(flexcpFontBold, InboxRow, 2) = i <> 0
        End If
        
        If CustomRequestsRow <> 0 Then
            i = Val("" & HFApp.SqlExec("select count(*) from unquotedoptions o left outer join divisioncommunities d on d.Community=o.Community where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")", dbHomefront)(0))
            CommandBar.TextMatrix(CustomRequestsRow, 2) = "Custom Requests (" & IIf(i = 0, "none", i) & ")"
            CommandBar.Cell(flexcpFontBold, CustomRequestsRow, 2) = i <> 0
        End If
        
        If HFApp.Options.ValueByName("BuildProCompanyCode") <> "" And POVendorAssignmentRow <> 0 Then
            i = Val("" & HFApp.SqlExec("select count(*) from pomaster where isnull(BuildPro_AssignedVendor,'')<>'' and divisionid=" & DbQuote(Num, HFApp.DivisionID), dbHomefront)(0))
            CommandBar.TextMatrix(POVendorAssignmentRow, 2) = "TBD Assignments (" & IIf(i = 0, "none", i) & ")"
            CommandBar.Cell(flexcpFontBold, POVendorAssignmentRow, 2) = i <> 0
        End If
    
    End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "close"
            CommandPanel.Visible = False
        Case "shrink"
            Toolbar.Buttons("expand").Visible = True
            Toolbar.Buttons("shrink").Visible = False
            Toolbar.Move 0, 0, 360, TASKPANELWIDTH
            CommandPanel.Move 0, 0, Toolbar.Width, Me.ScaleHeight - Toolbar.Height + 4 * Screen.TwipsPerPixelY
            CommandBar.Visible = False
        Case "expand"
            Toolbar.Buttons("expand").Visible = False
            Toolbar.Buttons("shrink").Visible = True
            Toolbar.Move 0, 0, TASKPANELWIDTH, 360
            CommandPanel.Move 0, 0, TASKPANELWIDTH, Me.ScaleHeight - Toolbar.Height + 4 * Screen.TwipsPerPixelY
            Call MDIForm_Resize
            CommandBar.Visible = True
        Case Else
    End Select
    
End Sub







Private Sub LoadCustomDescriptions()
    
    CD_Community = GetCustomDesc("Community")
    CD_Communities = GetCustomDesc("Community", True)
    CD_SqrFootage = GetCustomDesc("Square Footage")
    CD_Series = GetCustomDesc("Series")
    CD_Color = GetCustomDesc("Color")

End Sub




Private Sub CommandBarAddHeader(Name As String, Optional Collapsed As Boolean)
    With CommandBar
        .Rows = .Rows + 2
        .TextMatrix(.Rows - 1, 1) = Name
        If Collapsed Then .RowData(.Rows - 1) = "collapsed"
    End With
End Sub
Private Function CommandBarAddItem(Name As String, PictureKey As String, Optional Bold As Boolean, Optional ForeColor As OLE_COLOR) As Long
    Dim r As Long
    With CommandBar
        .AddItem ""
        r = .Rows - 1
        .TextMatrix(r, 2) = Name
        .Cell(flexcpPicture, r, 2) = SmallIcons.ListImages(PictureKey).ExtractIcon
        .Cell(flexcpFontBold, r, 2) = Bold
    End With
    CommandBarAddItem = r
End Function

Private Sub CommandBarFormat()
    
    Const TOPHEADING_BACKCOLOR = 9599096
    Const TOPHEADING_FORECOLOR = 16711402
    
    Const HEADING_BACKCOLOR = 16646142
    Const HEADING_FORECOLOR = 12612648
    
    Const ITEM_BACKCOLOR = 15788264
    Const ITEM_FORECOLOR = 12612134
    
    Const MARGIN_BACKCOLOR = 13285558
    
    Dim r As Long
    
    With CommandBar
    
    
        'left margins
        .ColWidth(0) = 15 * 12
        .ColWidth(1) = 15 * 9
        Call .Select(0, 0, .Rows - 1, 1)
        .FillStyle = flexFillRepeat
        .CellBackColor = MARGIN_BACKCOLOR
        
        'right margin
        Call .Select(0, 2, .Rows - 1, 2)
        Call .CellBorder(MARGIN_BACKCOLOR, 0, 0, 12, 0, 0, 0)
        
        'bottom margin
        .BackColorBkg = MARGIN_BACKCOLOR
        .SheetBorder = MARGIN_BACKCOLOR
    
        'format all rows as items
        For r = 0 To .Rows - 1
            .Row = r
            Call .Select(r, 1, r, 2)
            .RowHeight(.Row) = 15 * 21
            .CellFontSize = 9
            .CellForeColor = ITEM_FORECOLOR
            .CellBackColor = ITEM_BACKCOLOR
            .RowOutlineLevel(.Row) = 1
            .IsSubtotal(.Row) = True
        Next

        'reformat header rows
        For r = 1 To .Rows - 1
        If .TextMatrix(r, 1) <> "" Then
            .Row = r - 1
            Call .Select(r - 1, 0, r - 1, 2)
            .CellBackColor = MARGIN_BACKCOLOR
            .RowHeight(.Row) = 15 * 15
            .RowOutlineLevel(.Row) = 0
            
            .Row = r
            Call .Select(r, 1, r, 2)
            .RowHeight(.Row) = 15 * 24
            .CellFontBold = True
            .CellFontSize = 9
            
            If .Row = 1 Then
                .CellForeColor = TOPHEADING_FORECOLOR
                .CellBackColor = TOPHEADING_BACKCOLOR
            Else
                .CellForeColor = HEADING_FORECOLOR
                .CellBackColor = HEADING_BACKCOLOR
            End If
            .RowOutlineLevel(.Row) = 0
        End If
        Next
        
        
        
        For r = 1 To .Rows - 1
        If .TextMatrix(r, 1) <> "" Then
            If .RowData(r) = "collapsed" Then .IsCollapsed(r) = flexOutlineCollapsed
        End If
        Next
        
    End With
End Sub

Private Sub CommandBar_Click()
    Dim r As Long
    With CommandBar
        r = .MouseRow
        If r < 1 Then Exit Sub
        Select Case True
        
            'header
            Case .TextMatrix(r, 1) <> ""
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
                
            'spacer row
            Case .TextMatrix(r, 2) = ""
            
            'trim number of records off these tasks
            Case Left(.TextMatrix(r, 2), 5) = "Inbox"
                Call RunTask("Inbox", False)
            
            Case Left(.TextMatrix(r, 2), 15) = "TBD Assignments"
                Call RunTask("TBD Assignments", False)
                
            Case Left(.TextMatrix(r, 2), 15) = "Custom Requests"
                Call RunTask("Custom Requests", False)
            
            Case Else
                If GetAsyncKeyState(vbKeyControl) Then
                    Call RunTask(.TextMatrix(r, 2), True)
                Else
                    Call RunTask(Replace(.TextMatrix(r, 2), IIf(.TextMatrix(r, 2) <> "Project managers", CD_Community, "ZZZ"), "Community"), False)
                End If
                
                
        End Select
    End With
End Sub


Private Sub ResetDefaultVendors()
    Dim s As String

    If mCurrentJob = "" Then
        Call MsgBox("You must open a job first.", vbInformation, App.ProductName)
        Exit Sub
    End If
    
    s = "Reassign the default community vendor to all the un-generated items on job number " & mCurrentJob & "?"
    If vbCancel = MsgBox(s, vbQuestion + vbOKCancel, App.ProductName) Then Exit Sub
    
    s = ""
    s = s & "update estimateitems" & vbCrLf
    s = s & "set povendor= v.vendor_id" & vbCrLf
    s = s & "from estimateitems e" & vbCrLf
    s = s & "join tbljobs j on e.job=j.job_no" & vbCrLf
    s = s & "join tblvendors v on v.divisionid=e.divisionid and v.vendor_id=dbo.purch_getcommunityvendor(j.community,e.poindex,e.divisionid)" & vbCrLf
    s = s & "where e.pogenbatch=0" & vbCrLf
    s = s & "and e.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and e.job=" & DbQuote(Str, mCurrentJob) & vbCrLf
    s = s & vbCrLf
    s = s & "update estimateitems" & vbCrLf
    s = s & "set budgetvendor = v.vendor_id" & vbCrLf
    s = s & "from estimateitems e" & vbCrLf
    s = s & "join tbljobs j on e.job=j.job_no" & vbCrLf
    s = s & "join tblvendors v on v.divisionid=e.divisionid and v.vendor_id=dbo.purch_getcommunityvendor(j.community,e.poindex,e.divisionid)" & vbCrLf
    s = s & "where e.budgetgenerated=0" & vbCrLf
    s = s & "and e.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "and e.job=" & DbQuote(Str, mCurrentJob) & vbCrLf
    HFApp.SqlExec s
    
    s = "Vendors have been reassigned!" & vbCrLf & vbCrLf & "Costs have NOT been updated."
    Call MsgBox(s, vbExclamation + vbOK, App.ProductName)
    
End Sub
Private Sub DuplicateModelOptions()
    'copy options from selected model to this model
    Dim s As String
    Dim i As Long
    Dim j As Long
    Dim bCopyQuotes As Boolean
    Dim SourceAssemblies As String
    
    If vbCancel = MsgBox("This will copy model options from one model to another. Continue?", vbOKCancel + vbQuestion, App.ProductName) Then Exit Sub
    
    'pick source model
    s = ""
    s = s & "select distinct m.Model,m.Description,0 Source,m.Model ModelID" & vbCrLf
    s = s & "from tbldbassemblymaster o" & vbCrLf
    s = s & "join DistinctModelsByDivision m on o.divisionid=m.divisionid and m.model=o.model" & vbCrLf
    s = s & "where o.assemblytype=2" & vbCrLf
    s = s & "and o.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , "areaid,modelid,optionid,assemblytype,Source", , , , "Choose the options to be copied") Then Exit Sub
    
    'now pick source options
    s = ""
    s = s & "select ct.Description Category,c.area AreaID,c.description " & FMain.CD_Community & ",a.Model,dm.description ModelDescription,a.Series,a.Model modelid,a.OptionID,a.OptionID [Option],a.Assembly,a.Description,a.assemblytype,0 Source,a.assemblyid" & vbCrLf
    s = s & "from tbldbassemblymaster a left outer join tbllocality c on(a.community=c.area) left outer join tblcategories ct on(a.category=ct.category) left outer join distinctmodelsbydivision dm on(a.model=dm.model and a.divisionid=dm.divisionid)" & vbCrLf
    s = s & "where a.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & " and isnull(a.inactive,0)=0" & vbCrLf
    s = s & " and a.assemblytype=2" & vbCrLf
    s = s & " and a.model=" & DbQuote(Str, FPickList.SelectedItem("model")) & vbCrLf
    s = s & " order by 1,2,3,4,5,6,7"
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , "areaid,modelid,optionid,assemblytype,Source,assemblyid", True, , , "Choose the options to be copied") Then Exit Sub
    For i = 1 To FPickList.SelectedItems
        SourceAssemblies = SourceAssemblies & "," & FPickList.SelectedItem("assemblyid", i)
    Next
    SourceAssemblies = Mid(SourceAssemblies, 2)
    If SourceAssemblies = "" Then Exit Sub
        
    'pick destination models
    s = ""
    s = s & "select Model,Description from DistinctModelsByDivision where divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    If Not FPickList.Choose(HFApp.Databases(dbHomefront), IIf(HFApp.Options.ValueByName("BuilderType") = "Commercial", "Assemblies", "Model and Option"), s, , , , , , True, , , "Choose the models the options will be copied to") Then Exit Sub
        
    bCopyQuotes = vbYes = MsgBox("Do you want to copy Vendor pricing as well?", vbQuestion + vbYesNo + vbDefaultButton1, "Copy Vendor Pricing")
    
    s = ""
    For i = 1 To FPickList.SelectedItems
        For j = 1 To Parse(SourceAssemblies)
            s = "exec Purch_DuplicateModelOption " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Num, Parse(SourceAssemblies, j)) & "," & DbQuote(Str, FPickList.SelectedItem("Model", i)) & "," & DbQuote(Bit, bCopyQuotes)
            Call HFApp.SqlExec(s, dbHomefront)
        Next
    Next
    
    MsgBox "Options have been duplicated", , App.ProductName
    
    
End Sub

