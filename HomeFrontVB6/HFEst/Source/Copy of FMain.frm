VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Begin VB.MDIForm FMain 
   BackColor       =   &H8000000C&
   Caption         =   "Precision Builder"
   ClientHeight    =   7545
   ClientLeft      =   14175
   ClientTop       =   1920
   ClientWidth     =   12360
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   4560
   End
   Begin VB.PictureBox CommandPanel 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7545
      Left            =   0
      ScaleHeight     =   7545
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   0
      Width           =   3780
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
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
      Begin vbalExplorerBarLib6.vbalExplorerBarCtl CommandBar 
         Height          =   5025
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   8864
         BackColorEnd    =   0
         BackColorStart  =   0
         Begin cPopMenu6.PopMenu PopMenu 
            Left            =   1680
            Top             =   210
            _ExtentX        =   1058
            _ExtentY        =   1058
            HighlightCheckedItems=   0   'False
            TickIconIndex   =   0
         End
         Begin MSComctlLib.ImageList LargeIcons 
            Left            =   120
            Top             =   930
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   60
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":08CA
                  Key             =   "SaveAssembly"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":11A4
                  Key             =   "Send"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1A7E
                  Key             =   "AssemblyCosts"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2358
                  Key             =   "MassChange"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2C32
                  Key             =   "FieldPOs"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":350C
                  Key             =   "NewRFQ"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3DE6
                  Key             =   "CreateJob"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":46C0
                  Key             =   "takeoffsettings"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":4F9A
                  Key             =   "quote"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":5874
                  Key             =   "DecreaseDecimals"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":614E
                  Key             =   "IncreaseDecimals"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":6A28
                  Key             =   "ExcelImport"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":7302
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":7BDC
                  Key             =   "UpdatePrices"
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":84B6
                  Key             =   "ExcelExport"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":8D90
                  Key             =   "Publish"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":966A
                  Key             =   "Forecast"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":9F44
                  Key             =   "ViewPOs"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":A81E
                  Key             =   "ViewBudgets"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":B0F8
                  Key             =   "Open"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":B9D2
                  Key             =   "Preview"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":C2AC
                  Key             =   "SendPOs"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":CB86
                  Key             =   "SendRFQs"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":D460
                  Key             =   "TakeoffOneTime"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":DD3A
                  Key             =   "Estimate"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":E614
                  Key             =   "TakeoffAssembly"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":EEEE
                  Key             =   "NewAssembly"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":F7C8
                  Key             =   "TakeoffItemChart"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":100A2
                  Key             =   "New"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1097C
                  Key             =   "TakeoffItem"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":11256
                  Key             =   "TakeoffCustom"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":11B30
                  Key             =   "Save"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1240A
                  Key             =   "SaveAs"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":12CE4
                  Key             =   "Delete"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":135BE
                  Key             =   "AddPricelist"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":13E98
                  Key             =   "RePrice"
               EndProperty
               BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":14772
                  Key             =   "Pricebook"
               EndProperty
               BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1504C
                  Key             =   "PricebookEdit"
               EndProperty
               BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":15926
                  Key             =   "PricelistExport"
               EndProperty
               BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":16200
                  Key             =   "PricelistImport"
               EndProperty
               BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":16ADA
                  Key             =   "NewPricelist"
               EndProperty
               BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":173B4
                  Key             =   "View"
               EndProperty
               BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":17C8E
                  Key             =   "Vendor1"
               EndProperty
               BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":18568
                  Key             =   "Vendor"
               EndProperty
               BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":18E42
                  Key             =   "Add"
               EndProperty
               BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1971C
                  Key             =   "Attachments"
               EndProperty
               BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":19FF6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1A8D0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1B1AA
                  Key             =   "Design Center Options"
               EndProperty
               BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1BA84
                  Key             =   "Global Options"
               EndProperty
               BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1C35E
                  Key             =   "Models"
               EndProperty
               BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1CC38
                  Key             =   "Options"
               EndProperty
               BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1D512
                  Key             =   "Generate"
               EndProperty
               BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1DDEC
                  Key             =   "Timberline"
               EndProperty
               BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1E6C6
                  Key             =   "MBImport"
               EndProperty
               BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1EFA0
                  Key             =   "EEEstimating"
               EndProperty
               BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":1F87A
                  Key             =   "EEExport"
               EndProperty
               BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":20154
                  Key             =   "EEImport"
               EndProperty
               BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":20A2E
                  Key             =   "MBExport"
               EndProperty
               BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":21308
                  Key             =   "MasterBuilder"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList SmallIcons 
            Left            =   120
            Top             =   1560
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   77
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":21BE2
                  Key             =   "option"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2217C
                  Key             =   "combo1"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":222D6
                  Key             =   "Custom Requests"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":22BB0
                  Key             =   "MassChange"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2314A
                  Key             =   "RFP"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":236E4
                  Key             =   "quote"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":23C7E
                  Key             =   "sendreceive"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":24218
                  Key             =   "communitystandards"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":247B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2508C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":25966
                  Key             =   "pricelists"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":26240
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":26B1A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":273F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":27CCE
                  Key             =   "MB"
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":28268
                  Key             =   "QB"
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":28802
                  Key             =   "custom"
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":28D9C
                  Key             =   "assembly"
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":29336
                  Key             =   "links"
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":298D0
                  Key             =   "ItemDB"
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":29E6A
                  Key             =   "sendpos"
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2A404
                  Key             =   "HelpSearch"
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2A99E
                  Key             =   "Items"
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2AF38
                  Key             =   "New"
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2B4D2
                  Key             =   "Edit"
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2BA6C
                  Key             =   "HelpContents"
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2C006
                  Key             =   "EditAssembly"
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2C5A0
                  Key             =   "Forecast"
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2CB3A
                  Key             =   "shrink"
               EndProperty
               BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2D0D4
                  Key             =   "preview"
               EndProperty
               BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2D66E
                  Key             =   "customer"
               EndProperty
               BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2DC08
                  Key             =   "close"
               EndProperty
               BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2E1A2
                  Key             =   "expand"
               EndProperty
               BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2E73C
                  Key             =   "SaveAs"
               EndProperty
               BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2ECD6
                  Key             =   "Save"
               EndProperty
               BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2F270
                  Key             =   "RePrice"
               EndProperty
               BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2F80A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":2FDA4
                  Key             =   "estimating"
               EndProperty
               BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3033E
                  Key             =   "jobcost"
               EndProperty
               BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":308D8
                  Key             =   "error"
               EndProperty
               BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":30E72
                  Key             =   "salesworksheet"
               EndProperty
               BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3140C
                  Key             =   "ExcelExport"
               EndProperty
               BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":319A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":31F40
                  Key             =   "newworksheet"
               EndProperty
               BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":324DA
                  Key             =   "worksheet"
               EndProperty
               BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":32A74
                  Key             =   "purchaseorder"
               EndProperty
               BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3300E
                  Key             =   "ExcelImport"
               EndProperty
               BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":335A8
                  Key             =   "groupphase"
               EndProperty
               BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":33B42
                  Key             =   "costcode"
               EndProperty
               BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":340DC
                  Key             =   "job"
               EndProperty
               BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":34676
                  Key             =   "information"
               EndProperty
               BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":34C10
                  Key             =   "item"
               EndProperty
               BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":351AA
                  Key             =   "itemchecked"
               EndProperty
               BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":35744
                  Key             =   "phase"
               EndProperty
               BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":35CDE
                  Key             =   "vendor"
               EndProperty
               BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":36278
                  Key             =   "warning"
               EndProperty
               BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":36812
                  Key             =   "question"
               EndProperty
               BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":36DAC
                  Key             =   "category"
               EndProperty
               BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":37346
                  Key             =   "folder"
               EndProperty
               BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":378E0
                  Key             =   "AddItems"
               EndProperty
               BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":37E7A
                  Key             =   "model"
               EndProperty
               BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":38414
                  Key             =   "area"
               EndProperty
               BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":389AE
                  Key             =   "Delete"
               EndProperty
               BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":38F48
                  Key             =   "Underline"
               EndProperty
               BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":390A2
                  Key             =   "Bold"
               EndProperty
               BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":391FC
                  Key             =   "AlignCenter"
               EndProperty
               BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":39356
                  Key             =   "Italic"
               EndProperty
               BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":394B0
                  Key             =   "AlignLeft"
               EndProperty
               BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3960A
                  Key             =   "Bullet"
               EndProperty
               BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":39764
                  Key             =   "BulletNumber"
               EndProperty
               BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":398BE
                  Key             =   "AlignRight"
               EndProperty
               BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":39A18
                  Key             =   "printer"
               EndProperty
               BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":39FB2
                  Key             =   "defaultprinter"
               EndProperty
               BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3A54C
                  Key             =   "costcodes"
               EndProperty
               BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3AE26
                  Key             =   "defaultvendors"
               EndProperty
               BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3B700
                  Key             =   "cellcomments"
               EndProperty
               BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3BC9A
                  Key             =   "editpos"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList MultiStateIcons 
            Left            =   120
            Top             =   2250
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   16777215
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3C574
                  Key             =   "K00"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3C6CE
                  Key             =   "K01"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3C828
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3C982
                  Key             =   "K02"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3CADC
                  Key             =   "K10"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3CC36
                  Key             =   "K11"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3CD90
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3CEEA
                  Key             =   "K12"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3D044
                  Key             =   "K20"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3D19E
                  Key             =   "K21"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3D2F8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FMain.frx":3D452
                  Key             =   "K22"
               EndProperty
            EndProperty
         End
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
         Caption         =   "Precision Builder Help"
         Index           =   3
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "HomeFront Support"
         Index           =   4
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuHelpSub 
         Caption         =   "About HomeFront..."
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
            Begin VB.Menu mnuTakeOffModelsView 
               Caption         =   "Sales Options"
               Index           =   4
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
            Begin VB.Menu mnuTakeoffCustomOptions 
               Caption         =   "Unprocessed Custom Options"
            End
         End
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "View Timberline Assemblies"
            Index           =   5
            Begin VB.Menu mnuTakeOffAssemblyView 
               Caption         =   "Sort by Assembly"
               Index           =   6
            End
            Begin VB.Menu mnuTakeOffAssemblyView 
               Caption         =   "Sort by Description"
               Index           =   7
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
         Begin VB.Menu mnuTakeOffSub 
            Caption         =   "View Takeoffs"
            Index           =   7
            Begin VB.Menu mnuTakeOffEstimateView 
               Caption         =   "Sort by Job"
               Index           =   12
            End
         End
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "<mnuGrid>"
         Begin VB.Menu mnuGridSub 
            Caption         =   "Sort Ascending"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Sort Descending"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "-"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Hide this column"
            Index           =   4
         End
         Begin VB.Menu mnuGridSub 
            Caption         =   "Insert a column"
            Index           =   5
            Begin VB.Menu mnuColumnsSub 
               Caption         =   "(none available)"
               Enabled         =   0   'False
               Index           =   0
            End
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
            Begin VB.Menu mnuColumns2Sub 
               Caption         =   "(none available)"
               Enabled         =   0   'False
               Index           =   0
            End
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
               Caption         =   "Cancel PO..."
               Index           =   4
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "-"
               Index           =   5
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Edit Vendor..."
               Index           =   6
            End
            Begin VB.Menu mnuEstimateItemsPOSub 
               Caption         =   "Change Vendor..."
               Index           =   7
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
               Caption         =   "-"
               Index           =   4
            End
            Begin VB.Menu mnuEstimateItemsGridSub 
               Caption         =   "Attachments..."
               Index           =   5
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
               Caption         =   "-"
               Index           =   3
            End
            Begin VB.Menu mnuAssemblyItemsSub 
               Caption         =   "Attachments..."
               Index           =   4
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
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" ( _
   ByVal hLibModule As Long) As Long

Private m_hMod As Long


'custom descriptions
Public CD_Community As String
Public CD_Communities As String
Public CD_Series As String
Public CD_SqrFootage As String
Public CD_Color As String
 


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
        CurrentCustomer = "" & HFApp.SqlExec("select ar_customer_deposit from tblcustomers where job_no=" & DbQuote(str, mCurrentJob))(0)
    End If
End Property

Public Property Get CurrentCustomerDesc() As String
On Error Resume Next
    Dim s As String
    If mCurrentJob <> "" Then
        s = ""
        s = s & "select a.description" & vbCrLf
        s = s & "from tblcustomers c join arcustomers a on c.ar_customer_deposit=a.arcustomer" & vbCrLf
        s = s & "where c.job_no=" & DbQuote(str, mCurrentJob)
        CurrentCustomerDesc = "" & HFApp.SqlExec(s)(0)
    End If
End Property

Public Property Get CurrentJobDesc() As String
On Error Resume Next
    CurrentJobDesc = "" & HFApp.SqlExec("select description from tbljobs where DivisionID = " & HFApp.DivisionID & " and job_no=" & DbQuote(str, mCurrentJob))(0)
End Property

Private Sub MDIForm_Initialize()
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControls

End Sub


Private Sub MDIForm_Load()
On Error GoTo eh
Dim s As String

Call dbupgrade
App.HelpFile = "http://www.homefront-software.com/AppHelp/PrecisionBuilder/default.htm"

s = "init"
    mnuHidden.Enabled = InIde
    PopMenu.SubClassMenu Me
    Call IniGetForm(Me)
    Call SetToolbarIcons(Toolbar, SmallIcons)
    Toolbar.Visible = True
    
    
    If HFApp.DivisionID <> "" Then
        Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & ".0." & App.Revision & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomeFront)(0)
    Else
        Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & ".0." & App.Revision
    End If
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
    Dim DBUG As String

    Dim Est As Boolean
    Dim TL As Boolean
    Dim MB As Boolean
    Dim QB As Boolean
    Dim s As String

    
    
    Dim i As Long
    Timer1.Interval = 30000
    Timer1.Enabled = True
DBUG = "get integrations"
    Est = HFApp.Databases(dbEstimating).State = adStateOpen
    TL = HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asTimberline
    MB = HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asMasterBuilder
    QB = HFApp.Options(AccountingSystem) = asquickbooks
    
DBUG = "clear bar"
    Call CommandBar.Bars.Clear
    
        
    With CommandBar
    
        .Redraw = False
        .ImageList = SmallIcons.hImageList
        .BarTitleImageList = LargeIcons.hImageList
        
DBUG = "add purchasing items"
        With .Bars.Add(, , "Purchasing Tasks")
            .IsSpecial = True
            
            If HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440 Then
                Call .items.Add(, , "Retrieve Jobs from Sales Center", ImageIndex(SmallIcons, "sendreceive"))
            End If
            
            If HFApp.Options(SalesSystem) <> SalesSystems.asNone And HFApp.UserPermission("TaskEntQuote") Then
               
                i = 0
                On Error Resume Next
                If HFApp.DivisionID <> "" Then
                    i = Val("" & HFApp.SqlExec("select count(o.*) from unestimatedoptions o left outer join divisioncommunities d on d.Community=o.Community where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")", dbHomeFront)(0)) + Val("" & HFApp.SqlExec("select count(c.*) from unestimatedcustomers c left outer join divisioncommunities d on d.Community=c.Community where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")", dbHomeFront)(0))
                Else
                    i = Val("" & HFApp.SqlExec("select count(*) from unestimatedoptions", dbHomeFront)(0)) + Val("" & HFApp.SqlExec("select count(*) from unestimatedcustomers", dbHomeFront)(0))
                End If
                On Error GoTo eh
                If i = 0 Then
                    .items.Add(, , "Inbox (empty)", ImageIndex(SmallIcons, "job")).Bold = False
                Else
                    With .items.Add(, , "Inbox (" & i & ")", ImageIndex(SmallIcons, "job"))
                        .Bold = True
                        .TextColorOver = 0
                        .TextColor = vbRed
                    End With
                End If
               
               
                i = 0
                On Error Resume Next
                If HFApp.DivisionID <> "" Then
                    i = Val("" & HFApp.SqlExec("select count(o.*) from unquotedoptions o left outer join divisioncommunities d on d.Community=o.Community where (d.DivisionID is null or d.DivisionID = " & HFApp.DivisionID & ")", dbHomeFront)(0))
                Else
                    i = Val("" & HFApp.SqlExec("select count(*) from unquotedoptions", dbHomeFront)(0))
                End If
                On Error GoTo eh
                If i = 0 Then
                    .items.Add(, , "Custom Requests", ImageIndex(SmallIcons, "Custom Requests")).Bold = False
                Else
                    With .items.Add(, , "Custom Requests (" & i & ")", ImageIndex(SmallIcons, "Custom Requests"))
                        .Bold = True
                        .TextColorOver = 0
                        .TextColor = vbRed
                    End With
                End If
               
            End If
            
            If HFApp.UserPermission("IssueBudgets") Then
                Call .items.Add(, , "Prepare Job Quote", ImageIndex(SmallIcons, "quote"))
                
                Call .items.Add(, , "Issue budgets", ImageIndex(SmallIcons, "costcodes"))
            End If
            If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
                Call .items.Add(, , "Issue PO's", ImageIndex(SmallIcons, "editpos"))
                Call .items.Add(, , "Create Manual PO's", ImageIndex(SmallIcons, "editpos"))
            End If
            If HFApp.UserPermission("PostBudgets") Then
                If TL Then
                    Call .items.Add(, , "Post budgets", ImageIndex(SmallIcons, "jobcost"))
                ElseIf MB Then
                    Call .items.Add(, , "Post budgets", ImageIndex(SmallIcons, "MB"))
                ElseIf QB Then
                    Call .items.Add(, , "Post budgets", ImageIndex(SmallIcons, "QB"))
                End If
            End If
            '
            If HFApp.UserPermission("PostCommitments") Then
                If TL Then
                    Call .items.Add(, , "Post purchase orders", ImageIndex(SmallIcons, "jobcost"))
                ElseIf MB Then
                    Call .items.Add(, , "Post purchase orders", ImageIndex(SmallIcons, "MB"))
                ElseIf QB Then
                    Call .items.Add(, , "Post purchase orders", ImageIndex(SmallIcons, "QB"))
                End If
            End If
            If HFApp.UserPermission("SendPO") Then
                Call .items.Add(, , "Send Purchase Orders", ImageIndex(SmallIcons, "sendpos"))
            End If
'            Call .items.Add(, , "Send RFQ's", ImageIndex(SmallIcons, "sendpos"))
            
        End With
        
        
DBUG = "add setup items"
        With .Bars.Add(, , "Setup")
            .State = eBarCollapsed
            If HFApp.UserPermission("SetupProjectManager") Then Call .items.Add(, , "Project Managers", ImageIndex(SmallIcons, "newworksheet"))
            If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                If HFApp.UserPermission("SetupCommunity") Then Call .items.Add(, , CD_Community & " setup", ImageIndex(SmallIcons, "area"))
            Else
                If HFApp.UserPermission("SetupCommunity") Then Call .items.Add(, , "Work Region Setup", ImageIndex(SmallIcons, "area"))
            End If
            If HFApp.UserPermission("SetupJobs") Then Call .items.Add(, , "Job setup", ImageIndex(SmallIcons, "job"))
            If HFApp.UserPermission("SetupVendors") Then Call .items.Add(, , "Vendor setup", ImageIndex(SmallIcons, "vendor"))
            If HFApp.UserPermission("SetupPOIndex") Then Call .items.Add(, , "Purchase Orders", ImageIndex(SmallIcons, "purchaseorder"))
            If HFApp.UserPermission("EditItemDb") Then Call .items.Add(, , "Import items", ImageIndex(SmallIcons, "ExcelImport"))
            If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                If HFApp.UserPermission("EditAssemblies") Then Call .items.Add(, , "Import models and options", ImageIndex(SmallIcons, "ExcelImport"))
            Else
                If HFApp.UserPermission("EditAssemblies") Then Call .items.Add(, , "Import Assemblies", ImageIndex(SmallIcons, "ExcelImport"))
            End If
            If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                If HFApp.UserPermission("EditAssemblies") Then Call .items.Add(, , "Edit models and options", ImageIndex(SmallIcons, "EditAssembly"))
            Else
                If HFApp.UserPermission("EditAssemblies") Then Call .items.Add(, , "Edit Assemblies", ImageIndex(SmallIcons, "EditAssembly"))
            End If
            'Call .items.Add(, , "Library mass change", ImageIndex(SmallIcons, "MassChange"))
            If HFApp.UserPermission("EditItemDb") Then Call .items.Add(, , "Edit item database", ImageIndex(SmallIcons, "ItemDB"))
            If "True" = HFApp.Options.ValueByName("UseCommunityStandards") Then
                If HFApp.UserPermission("EditAssemblies") Then Call .items.Add(, , CD_Community & " Standards", ImageIndex(SmallIcons, "communitystandards"))
            End If
        End With
        
        
        
    DBUG = "add vendor items"
    If HFApp.UserPermission("SetupVendorPricing") Then
            With .Bars.Add(, , "Vendor Pricing")
                .State = eBarCollapsed
                Call .items.Add(, , "Edit item prices", ImageIndex(SmallIcons, "pricelists"))
                Call .items.Add(, , "Cost Forecasting", ImageIndex(SmallIcons, "Forecast"))
                Call .items.Add(, , "Export pricelists", ImageIndex(SmallIcons, "ExcelExport"))
                Call .items.Add(, , "Import pricelists", ImageIndex(SmallIcons, "ExcelImport"))
            End With
    End If
    
    If HFApp.UserPermission("OpenEstimatingWorksheets") Or HFApp.UserPermission("OpenMarketingWorksheets") Then
        DBUG = "add sales prcing items"
            With .Bars.Add(, , "Sales Pricing")
                .State = eBarCollapsed
                If Not HFApp.Options(UseEstimatingWorksheetsOnly) And HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    Call .items.Add(, , "Open a marketing worksheet", ImageIndex(SmallIcons, "worksheet"))
                End If
                If HFApp.UserPermission("OpenEstimatingWorksheets") Then Call .items.Add(, , "Open an estimating worksheet", ImageIndex(SmallIcons, "worksheet"))
                If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    If HFApp.UserPermission("OpenEstimatingWorksheets") Then Call .items.Add(, , "Open a design center worksheet", ImageIndex(SmallIcons, "worksheet"))
                End If
                If HFApp.UserPermission("OpenEstimatingWorksheets") Then Call .items.Add(, , "Review a published worksheet", ImageIndex(SmallIcons, "preview"))
            End With
            

    End If

    DBUG = "done - redraw"
    .Redraw = True
    End With

    
DBUG = "finished"
Exit Sub:
eh: Call errHandler(SRCFILE & "LoadInterface", DBUG)
End Sub






Private Sub CommandBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    If Left(itm.Text, 5) = "Inbox" Then
        Call RunTask("Inbox", False)
    ElseIf Left(itm.Text, 15) = "Custom Requests" Then
        Call RunTask("Custom Requests", False)
    Else
        If GetAsyncKeyState(vbKeyControl) Then
            Call RunTask(itm.Text, True)
        Else
            Call RunTask(Replace(itm.Text, IIf(itm.Text <> "Project managers", CD_Community, "ZZZ"), "Community"), False)
        End If
    End If
End Sub


Public Sub RunTask(TaskName As String, CtrlKey As Boolean)
On Error Resume Next

    Dim s As String
    Dim f As Form
    Dim i As Long
    Dim rs As Recordset
    
    Select Case TaskName
        
        Case "RFI's":                             Call HFApp.RunTask("RFIs|" & FMain.CurrentJob)
        
        Case "Field PO's"
            If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
                FFieldPOs.Show
            End If
        
        Case "Approvals"
            If HFApp.UserPermission("IssuePOs") Or HFApp.UserPermission("GeneratePOs") Then
                FApprovals.Show
            End If
            
        Case "Job Cost Codes":                        Call HFApp.RunTask("EditJCCostCodes")
        Case "Job Cost Categories":                   Call HFApp.RunTask("EditJCCategories")
        Case CD_Community & " Setup", "Community Setup":  If HFApp.UserPermission("SetupCommunity") Then Call HFApp.RunTask("EditCommunities")
        Case "Work Region Setup":                     If HFApp.UserPermission("SetupCommunity") Then Call HFApp.RunTask("EditCommunities")
        Case CD_Community & " Phases Setup":              If HFApp.UserPermission("SetupCommunity") Then Call FDBGrid.ShowForm("Community Phases", "select Community,CommunityPhase,Description from CommunityPhase order by 1,2", "CommunityPhase", "Community,CommunityPhase")
        Case "Default Vendors":                       If HFApp.UserPermission("SetupDefaultVendor") Then Call HFApp.RunTask("EditDefaultVendors")
        Case "Series":                                If HFApp.UserPermission("Setupseries") Then Call FDBGrid.ShowForm("Series List", "select Series,Description from tblSeries where DivisionID = " & HFApp.DivisionID, "tblSeries", "divisionid,series", HFApp.Options(SalesSystem) <> SalesSystems.asHomeFront)
        Case "Option Groups":                         If HFApp.UserPermission("SetupMajorGroup") Then Call FDBGrid.ShowForm("Option Groups", "select Major_Group,Description from tblmajorgroups", "tblmajorgroups", "major_group")
        
        
        Case "Option Categories"
            If HFApp.UserPermission("SetupCategory") Then
                Set rs = HFApp.SqlExec("select major_group from tblmajorgroups order by 1", dbHomeFront)
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
            If HFApp.Options(SalesSystem) = asBuilder1440 Then
                If FWebExport.ReadFrom1440(Not CtrlKey) Then Call FInboxJobs.Show(vbModal, Me)
            Else
                Call FInboxJobs.Show(vbModal, Me)
                'Call LoadInterface
            End If
            
        Case "Inbox"
            If HFApp.UserPermission("IssueBudgets") Then
                Call FInboxJobs.Show(vbModal, Me)
                'Call LoadInterface
            End If

        Case "Custom Requests"
            If HFApp.UserPermission("IssueBudgets") Then
                Call FInboxCustomQuote.Show(vbModal, Me)
                'Call LoadInterface
            End If
        
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
        Case "Purchase Orders":                   If HFApp.UserPermission("SetupPOIndex") Then Call HFApp.RunTask("EditPOIndexes")
        
        
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
        Case "Edit item database":                If HFApp.UserPermission("EditItemDb") Then Call FItems.Show(vbModal)
        '----------------------------
        Case "Edit item prices":
            If HFApp.UserPermission("SetupVendorPricing") Then
                Call FPriceList.Show
                Call FPriceList.SetFocus
            End If
            
        Case "Cost Forecasting":                  If HFApp.UserPermission("SetupVendorPricing") Then Call FCostForecast.Show(vbModal)
        Case "Export pricelists":                 If HFApp.UserPermission("SetupVendorPricing") Then Call FExportPricelists.Show(vbModal)
        Case "Import pricelists":                 If HFApp.UserPermission("SetupVendorPricing") Then Call FImportPricelists.ShowForm
        '----------------------------
        Case "Send Purchase Orders":              If HFApp.UserPermission("SendPO") Then Call FSendingWizard.ShowForm("PO")
        Case "Send RFQ's":                        If HFApp.UserPermission("SendPO") Then Call FSendingWizard.ShowForm("RFQ")
        
        Case "Review a published worksheet"
            If HFApp.UserPermission("OpenEstimatingWorkSheets") Then
                s = ""
                s = s & "SELECT Worksheet,Description,SalesEffectiveDate Posted, TStmp Modified, UStmp Author" & vbCrLf
                s = s & "FROM tblSalesSheetMaster" & vbCrLf
                s = s & "WHERE NOT SalesEffectiveDate IS NULL" & vbCrLf
                Set f = New FRptViewer
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Pricing Worksheet", s, , False, False) Then
                    i = Val(FPickList.SelectedItem("Worksheet"))
                    s = PathAppend(HFApp.SystemFolder, "System\Reports\Estimating\SalesSheetCosts.rpt")
                    Call f.ShowReport(s, True, False, "Worksheet", i)
                End If
            End If
            
        Case "Open a marketing worksheet"
            If HFApp.UserPermission("OpenMarketingWorkSheets") Then
                Set f = New FPricingWorksheet:
                s = ""
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective""" & vbCrLf
                s = s & "      ,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "     LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE SalesEffectiveDate IS NULL" & vbCrLf
                s = s & "  AND m.WorksheetType=0"
                s = s & " and m.DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Pricing Worksheet", s, , False, False) Then
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
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",m.SalesEffectiveDate Posted" & vbCrLf
                s = s & "      ,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "     LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE m.WorksheetType=0 and m.DivisionID = " & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Pricing Worksheet", s, , False, True) Then
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
                s = s & "SELECT m.Worksheet,m.Description,m.CostsEffectiveDate ""Costs Effective"",m.SalesEffectiveDate Posted" & vbCrLf
                s = s & "      ,l.User_ID + ' on ' + l.Workstation_name ""Locked by""" & vbCrLf
                s = s & "FROM tblSalesSheetMaster m" & vbCrLf
                s = s & "     LEFT OUTER JOIN TableLog l ON(l.TableName='tblSalesSheetMaster' AND RecordLocked=1 AND l.KeyRecord=CAST(m.Worksheet AS VARCHAR))" & vbCrLf
                s = s & "WHERE m.WorksheetType=1 and m.DivisionID=" & HFApp.DivisionID
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Pricing Worksheet", s, , False, True) Then
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
                
        Case Else
            MsgBox TaskName & " unknown"
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
    If HFApp.UserPermission("IssuePOs") Or (HFApp.UserPermission("GeneratePOs") And Index <> 4 And Index <> 6 And Index <> 7) Then
        Call Screen.ActiveForm.mnuEstimateItemsPOSub_Click(Index)
    Else
        MsgBox "You are not authorized to perform this function", vbInformation + vbOKOnly, "Access Denied"
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
            MouseCtrl.ColPosition(MouseCol) = 0
            Call Screen.ActiveForm.GroupGrid
            
        Case mcGRID_PRINT
            MouseCtrl.TopRow = MouseCtrl.FixedRows
            MouseCtrl.LeftCol = MouseCtrl.FixedCols
            Call MouseCtrl.Outline(-1)
            Call MouseCtrl.PrintGrid("", True, , 720, 720)
            
        Case mcGRID_SAVEAS
            If VBGetSaveFileName(s, , , "Excel (*.xls)|*.xls|Text (*.txt)|*.txt|Comma Separated (*.csv)|*.csv", , , , "txt", FMain.hWnd) Then
                Select Case UCase(FileExt(s))
                    Case "XLS":  Call MouseCtrl.SaveGrid(s, flexFileExcel, True)
                    Case "CSV":  Call MouseCtrl.SaveGrid(s, flexFileCommaText, True)
                    Case Else:   Call MouseCtrl.SaveGrid(s, flexFileTabText, True)
                End Select
            End If

        
    End Select
End Sub

Private Sub mnuGridSub_Click(Index As Integer)
On Error Resume Next
    Dim i As Long
    Dim s As String
    
    Select Case Index
        
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
            MouseCtrl.ColPosition(MouseCol) = 0
            On Error Resume Next
            Call Screen.ActiveForm.RefreshGrid
            
        Case mcGRID_PRINT
            MouseCtrl.TopRow = MouseCtrl.FixedRows
            MouseCtrl.LeftCol = MouseCtrl.FixedCols
            Call MouseCtrl.Outline(-1)
            Call MouseCtrl.PrintGrid("", True, , 720, 720)
            
        Case mcGRID_SAVEAS
            If VBGetSaveFileName(s, , , "Excel (*.xls)|*.xls|Text (*.txt)|*.txt|Comma Separated (*.csv)|*.csv", , , , "txt", FMain.hWnd) Then
                Call MouseCtrl.Outline(-1)
                Select Case UCase(FileExt(s))
                    Case "XLS":  Call MouseCtrl.SaveGrid(s, flexFileExcel, True)
                    Case "CSV":  Call MouseCtrl.SaveGrid(s, flexFileCommaText, True)
                    Case Else:   Call MouseCtrl.SaveGrid(s, flexFileTabText, True)
                End Select
            End If
        Case 10 'Attachments
            mTimerTask = "EditAttachments|" & MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("AttachmentID")) & "|Customer Option Attachments" & "|J~" & Replace(MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("Job_No")), "-", "") & "| Job Attachments"
            mTimerTask = mTimerTask & "|C~" & MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("Customer_No")) & "|Customer File Attachments"
            '"EditAttachments|" & MouseCtrl.TextMatrix(MouseCtrl.Row, MouseCtrl.ColIndex("AttachmentID")) & "|Customer Option Attachments"
            Timer1.Interval = 10
            Timer1.Enabled = True
            
    End Select
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    FreeLibrary m_hMod
    Call SetErrorMode(SEM_NOERRORS)
    Call SetErrorMode(SEM_NOGPFAULTERRORBOX)
    Call IniPutForm(Me)
    'get rid of any failed report windows
    On Error Resume Next
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
            If FPickList.Choose(HFApp.Databases(dbHomeFront), "Select Division", s, "", True, False, , "DivisionID") Then
                
                Call HFApp.SetDivision(FPickList.SelectedItem("DivisionID"))
                Call HFApp.Options.ReadData
                If HFApp.DivisionID <> "" Then
                    Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & ".0." & App.Revision & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomeFront)(0)
                Else
                    Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & ".0." & App.Revision
                End If
            End If
        Case mcFILE_CLOSE: Unload Me
    End Select
End Sub


Public Sub ShowColumnMenu(Grid, Optional Hideable As Boolean = True, _
                                Optional Showable As Boolean = True, _
                                Optional Renameable As Boolean = True, _
                                Optional Groupable As Boolean = False)
On Error GoTo eh

    Dim ParentMenu As Long
    Dim i As Long
    Dim j As Long

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
            
            'first unload columns
            Call .ClearSubMenusOfItem(.MenuIndex("mnuGrid2Sub(5)"))
            Call .ClearSubMenusOfItem(.MenuIndex("mnuGridSub(5)"))
            ParentMenu = .MenuIndex("mnuGrid2Sub(5)")
                
            'now load column names
            For i = 0 To MouseCtrl.cols - 1
                If MouseCtrl.TextMatrix(0, i) <> "" And MouseCtrl.ColHidden(i) Then
                    j = j + 1
                    .AddItem MouseCtrl.TextMatrix(0, i), "ShowColumn" & j, , , ParentMenu
                    .MenuTag("ShowColumn" & j) = MouseCtrl.ColKey(i)
                End If
            Next
            
            If j = 0 Then .AddItem "(none available)", , , , ParentMenu, , , False
            .AddItem "-", , , , ParentMenu
            .AddItem "More...", "ShowMoreColumns", , , ParentMenu
            'show menu
            PopupMenu mnuGrid2
        Else
            'save this stuff for menu click
            Set MouseCtrl = Grid
            MouseCol = Grid.MouseCol
        
            'set these
            mnuGridSub(mcGRID_RENAME).Enabled = Renameable
            mnuGridSub(mcGRID_HIDE).Enabled = MouseCol >= 0 And Hideable
            mnuGridSub(mcGRID_INSERT).Enabled = MouseCol >= 0 And Showable
            
            
            'first unload columns
            Call .ClearSubMenusOfItem(.MenuIndex("mnuGrid2Sub(5)"))
            Call .ClearSubMenusOfItem(.MenuIndex("mnuGridSub(5)"))
            ParentMenu = .MenuIndex("mnuGridSub(5)")
                
            'now load column names
            j = 0
            For i = 0 To MouseCtrl.cols - 1
                If MouseCtrl.TextMatrix(0, i) <> "" And MouseCtrl.ColHidden(i) Then
                    j = j + 1
                    .AddItem MouseCtrl.TextMatrix(0, i), "ShowColumn" & j, , , ParentMenu
                    .MenuTag("ShowColumn" & j) = MouseCtrl.ColKey(i)
                End If
            Next
            
            If j = 0 Then .AddItem "(none available)", , , , ParentMenu, , , False
            .AddItem "-", , , , ParentMenu
            .AddItem "More...", "ShowMoreColumns", , , ParentMenu
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
        Case 3 'HomeFront on the web
            Call ShellFile(Me.hWnd, "http://www.homefront-software.com/AppHelp/PrecisionBuilder/default.htm", , False)
        
        Case 4 'HomeFront Support
            Call ShellFile(Me.hWnd, "http://s3.parature.com/ics/support/default.asp?deptID=5534", , False)

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
                If HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asTimberline Then
                    .AddItem "Post budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "jobcost"), , HFApp.UserPermission("PostBudgets")
                    .AddItem "Post purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "jobcost"), , HFApp.UserPermission("PostCommitments")
                ElseIf HFApp.Databases(dbAccounting).State = adStateOpen And HFApp.Options(AccountingSystem) = asMasterBuilder Then
                    .AddItem "Post budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "MB"), , HFApp.UserPermission("PostBudgets")
                    .AddItem "Post purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "MB"), , HFApp.UserPermission("PostCommitments")
                ElseIf HFApp.Options(AccountingSystem) = asquickbooks Then
                    .AddItem "Post budgets", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "QB"), , HFApp.UserPermission("PostBudgets")
                    .AddItem "Post purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "QB"), , HFApp.UserPermission("PostCommitments")
                End If
                .AddItem "Send purchase orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "sendpos"), , HFApp.UserPermission("SendPO")
                .AddItem "Edit item prices", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "pricelists"), , HFApp.UserPermission("SetupVendorPricing")
                .AddItem "Cost Forecasting", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "Forecast"), , HFApp.UserPermission("SetupVendorPricing")
                            
            
            Case "mnuSalesPricing" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                If Not HFApp.Options(UseEstimatingWorksheetsOnly) And HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    .AddItem "Open a marketing worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "worksheet"), , HFApp.UserPermission("OpenMarketingWorkSheets")
                End If
                .AddItem "Open an estimating worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "worksheet"), , HFApp.UserPermission("OpenEstimatingWorkSheets")
                If HFApp.Options.ValueByName("BuilderType") <> "Commercial" Then
                    .AddItem "Open a design center worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "worksheet"), , HFApp.UserPermission("OpenEstimatingWorkSheets")
                End If
                .AddItem "Review a published worksheet", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "preview"), , HFApp.UserPermission("OpenEstimatingWorkSheets")
            
            Case "mnuSetup" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                .AddItem "Job Cost Codes", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes")
                .AddItem "Job Cost Categories", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "costcodes")
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
                .AddItem "Purchase Orders", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "purchaseorder"), , HFApp.UserPermission("SetupPOIndex")
                .AddItem "Default Vendors", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "defaultvendors"), , HFApp.UserPermission("SetupDefaultVendor")
                .AddItem "Edit item database", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "ItemDB"), , HFApp.UserPermission("EditItemDb")
                 If "True" = HFApp.Options.ValueByName("UseCommunityStandards") Then
                    .AddItem FMain.CD_Community & " Standards", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "communitystandards")
                End If
                If HFApp.Options.ValueByName("BuilderType") = "Commercial" Then
                    .AddItem "Edit Assemblies", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "EditAssembly"), , HFApp.UserPermission("EditAssemblies")
                Else
                    .AddItem "Edit models and options", , "task", , ParentItemNumber, ImageIndex(SmallIcons, "EditAssembly"), , HFApp.UserPermission("EditAssemblies")
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
                    Case asTimberline, asMasterBuilder
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", App.Path & "\HFSync.exe ReadARCustomers" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", App.Path & "\HFSync.exe ReadAllVendors" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", App.Path & "\HFSync.exe ReadAllStandardCostCodes" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", App.Path & "\HFSync.exe ReadTaxGroups" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", App.Path & "\HFSync.exe ReadGLAccounts" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                    Case asquickbooks
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", App.Path & "\HFSync.exe ReadARCustomers" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", App.Path & "\HFSync.exe ReadAllVendors" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", App.Path & "\HFSync.exe ReadAllStandardCostCodes" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", App.Path & "\HFSync.exe ReadTaxGroups" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "-", , , , ParentItemNumber
                        .AddItem "Import Payroll Costs from Accounting", App.Path & "\HFSync.exe ReadPayrollCosts" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                    Case asSimply
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        .AddItem "Read Accounts Receivable from Accounting", App.Path & "\HFSync.exe ReadARCustomers" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", App.Path & "\HFSync.exe ReadAllVendors" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", App.Path & "\HFSync.exe ReadTaxGroups" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                    Case asPeachtree
                       If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        '.AddItem "Read Accounts Receivable from Accounting", App.Path & "\HFSync.exe ReadARCustomers|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", App.Path & "\HFSync.exe ReadAllVendors" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", App.Path & "\HFSync.exe ReadTaxGroups" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        '.AddItem "Read GL Accounts from Accounting", App.Path & "\HFSync.exe ReadGLAccounts|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Cost Codes from Accounting", App.Path & "\HFSync.exe ReadAllStandardCostCodes" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                    Case asMYOB
                        If b Then .AddItem "-", , , , ParentItemNumber
                        b = True
                        '.AddItem "Read Accounts Receivable from Accounting", App.Path & "\HFSync.exe ReadARCustomers|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Vendors from Accounting", App.Path & "\HFSync.exe ReadAllVendors" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read Tax Settings from Accounting", App.Path & "\HFSync.exe ReadTaxGroups" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                        .AddItem "Read GL Accounts from Accounting", App.Path & "\HFSync.exe ReadGLAccounts" & ":" & HFApp.DivisionID & ":" & HFApp.LoginID & ":" & HFApp.LoginPswd & "|" & HFApp.ConnectionString(dbHomeFront), , , ParentItemNumber
                     
                    Case Else
                End Select
                
                
            Case "mnuTools" = .MenuKey(.UltimateParent(ParentItemNumber))
                Call .ClearSubMenusOfItem(ParentItemNumber)
                
                Call .AddItem("Mass Change Wizard", , , , ParentItemNumber)
                Call .AddItem("Work Reassignment Wizard", , , , ParentItemNumber)
                
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
                    .AddItem "User Manager", , "task", , ParentItemNumber, , , HFApp.UserPermission("AdminUser")
                    If HFApp.Databases(dbAccounting).State = adStateOpen Or HFApp.Options(SalesSystem) = asBuilder1440 Or HFApp.Options(AccountingSystem) = asquickbooks Or HFApp.Options(AccountingSystem) = asSimply Or HFApp.Options(AccountingSystem) = asMYOB Or HFApp.Options(AccountingSystem) = asPeachtree Then
                        Call .AddItem("(dummy)", , , , .AddItem("Database Synchronization", "Sync", , , ParentItemNumber))
                    End If
                    .AddItem "&Options", "AppOptions", , , ParentItemNumber, , , (HFApp.UserPermission("AdminUser") Or HFApp.UserPermission("PrecisionBldrOptions"))
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

            Case .MenuKey(ItemNumber) = "import":                          Call ReadPipelineXML("")
            
            
            
            Case .Caption(ItemNumber) = "Mass Change Wizard":              Call FMassChange.ShowForm
            Case .Caption(ItemNumber) = "Work Reassignment Wizard":        Call FVendorChange.ShowForm
            Case .Caption(ItemNumber) = "User Manager":                    Call frmuser.ShowForm
            Case .Caption(ItemNumber) = "Change Division"
                Dim divID As String
                s = "select  d.DivisionID,d.DivisionCode, d.DivisionName from Divisions d left outer join divisionusers u on u.DivisionID = d.DivisionID where u.userID = " & DbQuote(str, HFApp.LoginID) & " or u.userid is null order by 2"
                If FPickList.Choose(HFApp.Databases(dbHomeFront), "Select Division", s, HFApp.DivisionID, True, False, , "DivisionID") Then
                    Call HFApp.SetDivision(FPickList.SelectedItem("DivisionID"))
                    Call HFApp.Options.ReadData
                    If HFApp.DivisionID <> "" Then
                        Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & ".0." & App.Revision & " Division: " & "" & HFApp.SqlExec("Select DivisionCode from Divisions where DivisionID = " & HFApp.DivisionID, dbHomeFront)(0)
                    Else
                        Me.Caption = "Precision Builder (" & Trim(HFApp.LoginDSN) & ", Login User: " & Trim(HFApp.LoginID) & ") - version " & App.Major & "." & App.Minor & ".0." & App.Revision
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
                Call ShellFile(Me.hWnd, s)
            
                
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
                    Call ShellFile(Me.hWnd, s)
                End If
            
            Case "mnuInquiries" = .MenuKey(.UltimateParent(ItemNumber))
                Set f = New FInquiry
                Call f.ShowInquiry(.MenuKey(ItemNumber))
                
            
            Case "mnuReports" = .MenuKey(.UltimateParent(ItemNumber))
                s = .MenuKey(ItemNumber)
                
                On Error Resume Next
                Job = Screen.ActiveForm.Job
                PricingCommunity = Screen.ActiveForm.PricingCommunity
                community = Screen.ActiveForm.community
                Model = Screen.ActiveForm.Model
                OptionID = Screen.ActiveForm.OptionID
                Assembly = Screen.ActiveForm.Assembly
                On Error GoTo eh
                
                If FileExt(s) = "rpt" Then
                    Set f = New FRptViewer
                    Call f.ShowReport(s, True, False, "Job", Job, "Model", Model, "Assembly", Assembly, "OptionID", OptionID, "Community", community, "PricingCommunity", PricingCommunity)
                Else
                    Call ShellFile(Me.hWnd, s)
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
    
    Timer1.Interval = 30000
    i = 0
    If HFApp.Options.ValueByName("DisableAutoUpdate") <> "True" Then
        With CommandBar.Bars
        On Error Resume Next
        s = "select count(*) from unestimatedoptions o"
        If HFApp.DivisionID <> "" Then
                 s = s & " join divisioncommunities d on d.Community=o.Community" & vbCrLf
                 s = s & " where d.divisionid = " & HFApp.DivisionID & vbCrLf
        End If
        
        s2 = "select count(*) from unestimatedcustomers o"
        If HFApp.DivisionID <> "" Then
                 s2 = s2 & " join divisioncommunities d on d.Community=o.Community" & vbCrLf
                 s2 = s2 & " where d.divisionid = " & HFApp.DivisionID & vbCrLf
        End If
        
        i = Val("" & HFApp.SqlExec(s, dbHomeFront)(0)) + Val("" & HFApp.SqlExec(s2, dbHomeFront)(0))
            
        If HFApp.Options(SalesSystem) = SalesSystems.asBuilder1440 Then
            If i = 0 Then
                CommandBar.Bars.Item(1).items(2).Text = "Inbox (empty)"
            Else
                CommandBar.Bars.Item(1).items(2).Text = "Inbox (" & i & ")"
            End If
            i = 0
            On Error Resume Next
            s = "select count(*) from unquotedoptions o"
            If HFApp.DivisionID <> "" Then
                 s = s & " join divisioncommunities d on d.Community=o.Community" & vbCrLf
                 s = s & " where d.divisionid = " & HFApp.DivisionID & vbCrLf
            End If
            i = Val("" & HFApp.SqlExec(s, dbHomeFront)(0))
            If i = 0 Then
                CommandBar.Bars.Item(1).items(3).Text = "Custom Requests"
            Else
               CommandBar.Bars.Item(1).items(3).Text = "Custom Requests (" & i & ")"
            End If
        
        Else
            If i = 0 Then
                CommandBar.Bars.Item(1).items(1).Text = "Inbox (empty)"
            Else
                CommandBar.Bars.Item(1).items(1).Text = "Inbox (" & i & ")"
            End If
            i = 0
            On Error Resume Next
            s = "select count(*) from unquotedoptions o"
            If HFApp.DivisionID <> "" Then
                 s = s & " join divisioncommunities d on d.Community=o.Community" & vbCrLf
                 s = s & " where d.divisionid = " & HFApp.DivisionID & vbCrLf
            End If
            i = Val("" & HFApp.SqlExec(s, dbHomeFront)(0))
            If i = 0 Then
                CommandBar.Bars.Item(1).items(2).Text = "Custom Requests"
            Else
                CommandBar.Bars.Item(1).items(2).Text = "Custom Requests (" & i & ")"
            End If
        End If
        End With
    End If
    If Mid(mTimerTask, 1, 15) = "EditAttachments" Then
         Call HFApp.RunTask(mTimerTask)
         mTimerTask = ""
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



Private Sub dbupgrade()
On Error GoTo eh
'    Call HFApp.SqlExec("alter table poitems add LineNumber int", dbHomeFront)
'    Call HFApp.SqlExec("alter table poitems add LineDescription varchar(200)", dbHomeFront)
'    Call HFApp.SqlExec("alter table poitems drop column Summarized", dbHomeFront)
'
'    Call HFApp.SqlExec("alter table pomaster add Summarized bit", dbHomeFront)
'
'    Call HFApp.SqlExec("alter table EstimateItems add SalesQty float", dbHomeFront)
'
'    Call HFApp.SqlExec("alter view dbo.DistinctModelSeries as select Model,Series, min(description) Description ,cast(min(cast(inactive as int)) as bit) Inactive from tblmodels where description <> 'Unknown Model' group by model,series", dbHomeFront)
Exit Sub
eh:
If Err.Number <> 0 Then
    MsgBox Err.Description
    Resume Next
End If

End Sub




Private Sub LoadCustomDescriptions()
    
    CD_Community = GetCustomDesc("Community")
    CD_Communities = GetCustomDesc("Community", True)
    CD_SqrFootage = GetCustomDesc("Square Footage")
    CD_Series = GetCustomDesc("Series")
    CD_Color = GetCustomDesc("Color")

End Sub

