VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FJob 
   Caption         =   "Job Setup"
   ClientHeight    =   11070
   ClientLeft      =   6525
   ClientTop       =   2055
   ClientWidth     =   20430
   Icon            =   "FJob.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   20430
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   20430
      _ExtentX        =   36036
      _ExtentY        =   1058
      ButtonWidth     =   2487
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Attachments"
            Key             =   "Attachments"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Correspondence"
            Key             =   "JobDocs"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList SmallIcons 
         Left            =   7380
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   73
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":000C
               Key             =   "option"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":05A6
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":0B40
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":10DA
               Key             =   "sendreceive"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1674
               Key             =   "communitystandards"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1C0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":24E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2DC2
               Key             =   "pricelists"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":369C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":3F76
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":4850
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":512A
               Key             =   "MB"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":56C4
               Key             =   "QB"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":5C5E
               Key             =   "custom"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":61F8
               Key             =   "assembly"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":6792
               Key             =   "links"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":6D2C
               Key             =   "ItemDB"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":72C6
               Key             =   "sendpos"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":7860
               Key             =   "HelpSearch"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":7DFA
               Key             =   "Items"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":8394
               Key             =   "New"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":892E
               Key             =   "Edit"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":8EC8
               Key             =   "HelpContents"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":9462
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":99FC
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":9F96
               Key             =   "shrink"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":A530
               Key             =   "preview"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":AACA
               Key             =   "customer"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":B064
               Key             =   "close"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":B5FE
               Key             =   "expand"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":BB98
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":C132
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":C6CC
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":CC66
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":D200
               Key             =   "estimating"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":D79A
               Key             =   "jobcost"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":DD34
               Key             =   "error"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":E2CE
               Key             =   "salesworksheet"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":E868
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":EE02
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":F39C
               Key             =   "newworksheet"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":F936
               Key             =   "worksheet"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":FED0
               Key             =   "purchaseorder"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1046A
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":10A04
               Key             =   "groupphase"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":10F9E
               Key             =   "costcode"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":11538
               Key             =   "job"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":11AD2
               Key             =   "information"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1206C
               Key             =   "item"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":12606
               Key             =   "itemchecked"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":12BA0
               Key             =   "phase"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1313A
               Key             =   "vendor"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":136D4
               Key             =   "warning"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":13C6E
               Key             =   "question"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":14208
               Key             =   "category"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":147A2
               Key             =   "folder"
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":14D3C
               Key             =   "AddItems"
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":152D6
               Key             =   "model"
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":15870
               Key             =   "area"
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":15E0A
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":163A4
               Key             =   "Underline"
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":164FE
               Key             =   "Bold"
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":16658
               Key             =   "AlignCenter"
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":167B2
               Key             =   "Italic"
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1690C
               Key             =   "AlignLeft"
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":16A66
               Key             =   "Bullet"
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":16BC0
               Key             =   "BulletNumber"
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":16D1A
               Key             =   "AlignRight"
            EndProperty
            BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":16E74
               Key             =   "printer"
            EndProperty
            BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1740E
               Key             =   "defaultprinter"
            EndProperty
            BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":179A8
               Key             =   "costcodes"
            EndProperty
            BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":18282
               Key             =   "defaultvendors"
            EndProperty
            BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":18B5C
               Key             =   "editpos"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList LargeIcons 
         Left            =   6780
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   56
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":19436
               Key             =   "EditAssembly"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":19D10
               Key             =   "RFP"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1A5EA
               Key             =   "CreateJob"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1AEC4
               Key             =   "takeoffsettings"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1B79E
               Key             =   "quote"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1C078
               Key             =   "JobDocs"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1C952
               Key             =   "DecreaseDecimals"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1D22C
               Key             =   "IncreaseDecimals"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1DB06
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1E3E0
               Key             =   "ExcelImport"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1ECBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1F594
               Key             =   "UpdatePrices"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":1FE6E
               Key             =   "ExcelExport"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":20748
               Key             =   "Publish"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":21022
               Key             =   "Forecast"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":218FC
               Key             =   "ViewPOs"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":221D6
               Key             =   "ViewBudgets"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":22AB0
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2338A
               Key             =   "Preview"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":23C64
               Key             =   "Send"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2453E
               Key             =   "TakeoffOneTime"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":24E18
               Key             =   "Estimate"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":256F2
               Key             =   "TakeoffAssembly"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":25FCC
               Key             =   "NewAssembly"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":268A6
               Key             =   "New"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":27180
               Key             =   "TakeoffItem"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":27A5A
               Key             =   "TakeoffCustom"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":28334
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":28C0E
               Key             =   "SaveAs"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":294E8
               Key             =   "Delete"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":29DC2
               Key             =   "AddPricelist"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2A69C
               Key             =   "RePrice"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2AF76
               Key             =   "Pricebook"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2B850
               Key             =   "PricebookEdit"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2C12A
               Key             =   "PricelistExport"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2CA04
               Key             =   "PricelistImport"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2D2DE
               Key             =   "NewPricelist"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2DBB8
               Key             =   "View"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2E492
               Key             =   "Vendor1"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2ED6C
               Key             =   "Vendor"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2F646
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":2FF20
               Key             =   "Attachments"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":307FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":310D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":319AE
               Key             =   "Design Center Options"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":32288
               Key             =   "Global Options"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":32B62
               Key             =   "Models"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":3343C
               Key             =   "Options"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":33D16
               Key             =   "Generate"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":345F0
               Key             =   "Timberline"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":34ECA
               Key             =   "MBImport"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":357A4
               Key             =   "EEEstimating"
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":3607E
               Key             =   "EEExport"
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":36958
               Key             =   "EEImport"
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":37232
               Key             =   "MBExport"
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FJob.frx":37B0C
               Key             =   "MasterBuilder"
            EndProperty
         EndProperty
      End
   End
   Begin HFSystem.VBCombo cboProvince 
      Height          =   240
      Left            =   3435
      TabIndex        =   14
      Top             =   4305
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   423
      Text            =   "Combo1"
   End
   Begin VB.TextBox txtExternalID 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1380
      Width           =   2865
   End
   Begin VB.TextBox txtARCustomer 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1875
      Width           =   2640
   End
   Begin VB.TextBox txtNotes 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4320
      MaxLength       =   8000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "FJob.frx":387E6
      Top             =   1170
      Width           =   6975
   End
   Begin VB.TextBox txtLotPlan 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3000
      Width           =   945
   End
   Begin VB.TextBox txtBlock 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3000
      Width           =   945
   End
   Begin VB.TextBox txtLot 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3000
      Width           =   945
   End
   Begin HFSystem.VBCombo cboStatus 
      Height          =   240
      Left            =   1320
      TabIndex        =   11
      Top             =   3240
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   423
      Text            =   "Combo1"
   End
   Begin VSFlex8Ctl.VSFlexGrid gProperties 
      Height          =   3135
      Left            =   180
      TabIndex        =   18
      Top             =   5580
      Width           =   4035
      _cx             =   7117
      _cy             =   5530
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   15
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FJob.frx":387E8
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   1
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
      OleDropMode     =   1
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
   Begin VB.TextBox txtPostal 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   15
      Top             =   4560
      Width           =   1395
   End
   Begin VB.TextBox txtFax 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   17
      Top             =   5040
      Width           =   1755
   End
   Begin VB.TextBox txtPhone 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   16
      Top             =   4800
      Width           =   1755
   End
   Begin VB.TextBox txtCity 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   13
      Top             =   4305
      Width           =   2100
   End
   Begin VB.TextBox txtAddress 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   420
      Left            =   1335
      MaxLength       =   75
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3870
      Width           =   2865
   End
   Begin VB.TextBox txtDescription 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   230
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1140
      Width           =   2865
   End
   Begin HFSystem.VBCombo cboCommunity 
      Height          =   240
      Left            =   1320
      TabIndex        =   5
      Top             =   2235
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   423
   End
   Begin HFSystem.VBCombo cboModel 
      Height          =   240
      Left            =   1320
      TabIndex        =   7
      Top             =   2745
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   423
   End
   Begin HFSystem.VBCombo cboGLPrefix 
      Height          =   240
      Left            =   1320
      TabIndex        =   3
      Top             =   1620
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   423
   End
   Begin VSFlex8Ctl.VSFlexGrid gPOs 
      Height          =   2985
      Left            =   4320
      TabIndex        =   22
      Top             =   7530
      Width           =   16995
      _cx             =   1986360601
      _cy             =   1986335889
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FJob.frx":388CA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
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
      OleDropMode     =   1
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
   Begin HFSystem.VBCombo cboPhase 
      Height          =   240
      Left            =   1320
      TabIndex        =   6
      Top             =   2490
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   423
   End
   Begin VSFlex8Ctl.VSFlexGrid gJobContacts 
      Height          =   2625
      Left            =   4350
      TabIndex        =   21
      Top             =   4560
      Width           =   11295
      _cx             =   1986350547
      _cy             =   1986335254
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FJob.frx":38B27
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
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid gContacts 
      Height          =   1515
      Left            =   4350
      TabIndex        =   20
      Top             =   2670
      Width           =   6975
      _cx             =   1986342927
      _cy             =   1986333296
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FJob.frx":38C7C
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
      OutlineBar      =   5
      OutlineCol      =   1
      Ellipsis        =   0
      ExplorerBar     =   7
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
      OleDropMode     =   1
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
   Begin VB.TextBox txtJob 
      BorderStyle     =   0  'None
      Height          =   230
      Left            =   1320
      TabIndex        =   0
      Top             =   900
      Width           =   2640
   End
   Begin VB.TextBox txtGLPrefix 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   240
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   23
      Text            =   "0"
      Top             =   1620
      Width           =   2865
   End
   Begin VB.Label lblExternalID 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ExternalID"
      Height          =   195
      Left            =   465
      TabIndex        =   44
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label lblLink 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AR Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   43
      Top             =   1890
      Width           =   930
   End
   Begin VB.Image cmdBrowse 
      Height          =   240
      Index           =   10
      Left            =   3975
      Picture         =   "FJob.frx":38D8B
      Top             =   1875
      Width           =   240
   End
   Begin VB.Label lblContacts 
      AutoSize        =   -1  'True
      Caption         =   "Personnel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4320
      TabIndex        =   42
      Top             =   2430
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4350
      TabIndex        =   41
      Top             =   4260
      Width           =   765
   End
   Begin VB.Image cmdJob 
      Height          =   240
      Left            =   3975
      Picture         =   "FJob.frx":38ED5
      Top             =   900
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4320
      TabIndex        =   40
      Top             =   930
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phase"
      Height          =   195
      Left            =   765
      TabIndex        =   39
      Top             =   2505
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Purchase Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4350
      TabIndex        =   38
      Top             =   7260
      Width           =   1425
   End
   Begin VB.Label Label1120 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Model"
      Height          =   195
      Left            =   765
      TabIndex        =   37
      Top             =   2790
      Width           =   435
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lot/Block/Plan"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   36
      Top             =   3030
      Width           =   1095
   End
   Begin VB.Label lblCommunity 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Community"
      Height          =   195
      Left            =   450
      TabIndex        =   35
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   765
      TabIndex        =   34
      Top             =   3270
      Width           =   450
   End
   Begin VB.Label lblGLPrefix 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "GL Prefix"
      Height          =   195
      Left            =   555
      TabIndex        =   32
      Top             =   1650
      Width           =   645
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Postal"
      Height          =   195
      Left            =   765
      TabIndex        =   31
      Top             =   4545
      Width           =   435
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   510
      TabIndex        =   30
      Top             =   3870
      Width           =   690
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      Height          =   195
      Left            =   945
      TabIndex        =   29
      Top             =   5025
      Width           =   255
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Phone"
      Height          =   195
      Left            =   735
      TabIndex        =   28
      Top             =   4785
      Width           =   465
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City/Prov"
      Height          =   195
      Left            =   540
      TabIndex        =   27
      Top             =   4320
      Width           =   660
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Job Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   26
      Top             =   930
      Width           =   1020
   End
   Begin VB.Label lblProperties 
      AutoSize        =   -1  'True
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   180
      TabIndex        =   25
      Top             =   5310
      Width           =   870
   End
   Begin VB.Label Label112 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   405
      TabIndex        =   33
      Top             =   1170
      Width           =   795
   End
End
Attribute VB_Name = "FJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public EventTraps As Collection
Const SRCFILE = "FJob::"

Private mDirty    As Boolean
Private mJob      As String
Private mIsCanadian As Boolean
Private mIsAustralia As Boolean
Private mIsMultifamily As Boolean

'property rows
Private mprop_PermitNumber As Integer
Private mprop_PermitDate As Integer
Private mprop_ShellTemplate As Integer
Private mprop_PreconTemplate As Integer
Private mprop_UnitTemplate As Integer
Private mprop_PreconStart As Integer
Private mprop_ConstStart As Integer
Private mprop_IntacctDepartment As Integer
Private mprop_TarionBuilderNumber  As Integer
Private mprop_TarionEnrollmentNumber As Integer
Private mprop_LabTax As Integer
Private mprop_MatTax As Integer
Private mprop_SubTax As Integer
Private mprop_EquTax As Integer
Private mprop_OvrTax As Integer
Private mprop_OthTax As Integer
Private mprop_ARTax As Integer
Private mprop_UserFlds As Integer



Private Property Get Dirty() As Boolean
    Dirty = mDirty
End Property
Private Property Let Dirty(RHS As Boolean)
    mDirty = RHS
    Toolbar.Buttons("Save").Enabled = mDirty
End Property

Public Property Get Job() As String
    Job = mJob
End Property

Private Sub cboCommunity_Change()
    Dirty = True
End Sub

Private Sub cboCommunity_Click()
    On Error Resume Next
    Dirty = True
    
    
    Dim s As String
    
    s = ""
    s = s & "select distinct communityphase ,'',0 " & vbCrLf
    s = s & "from tbljobs " & vbCrLf
    s = s & "where community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "and isnull(communityphase,'')<>'' " & vbCrLf
    s = s & "union " & vbCrLf
    s = s & "select communityphase,'',0" & vbCrLf
    s = s & "from communityphase" & vbCrLf
    s = s & "where community=" & DbQuote(Str, GetComboBoxListKey(cboCommunity)) & vbCrLf
    s = s & "ORDER BY 1" & vbCrLf
    
    Call LoadComboBox(cboPhase, HFApp.Databases(dbHomefront), s)


End Sub

Private Sub cboGLPrefix_Click()
    Dirty = True
End Sub

Private Sub cboModel_Change()
    Dirty = True
End Sub

Private Sub cboModel_Click()
    Dirty = True

End Sub

Private Sub cboPhase_Change()
    Dirty = True
End Sub

Private Sub cboPhase_Click()
    Dirty = True
End Sub

Private Sub cboProvince_Change()
    Dirty = True
End Sub

Private Sub cboProvince_Click()
    Dirty = True
End Sub

Private Sub cboStatus_Click()
    Dirty = True
End Sub


Private Sub cmdBrowse_Click(Index As Integer)
    Dim s As String
    If Not cmdBrowse(Index).Enabled Then Exit Sub
    Select Case Index
        Case 10:
        
'            s = ""
'            s = s & "Active Customers" & Chr(1) & "select arcustomer Customer,Description from arcustomers where divisionid=" & HFApp.DivisionID & " and isnull(inactive,0)=0" & Chr(0)
'            s = s & "Inactive Customers" & Chr(1) & "select arcustomer Customer,Description from arcustomers where divisionid=" & HFApp.DivisionID & " and isnull(inactive,0)=1"
'            If FPickList.Choose(HFApp.Databases(dbHomefront), "Customer", s, txtARCustomer.Text, , True) Then
'
'                txtARCustomer.Text = FPickList.SelectedItem("Customer")
'                If txtARCustomer.Text = "" Then Call HFApp.RunTask("EditCustomer|")
'                Dirty = True
'            End If
'
'
    
            s = "select ARCustomer,Description,ExternalID from arcustomers where divisionid=" & HFApp.DivisionID
        
            'only allow new customer on single family dbs
            If Not FPickList.Choose(HFApp.Databases(dbHomefront), "Customer", s, , , HFApp.Options(MultiFamily) = "False", , "ExternalID") Then Exit Sub
            
            If HFApp.Options(MultiFamily) = "False" And FPickList.SelectedItem("arcustomer") = "" Then
            
                s = ""
                s = s & "insert into arcustomers(" & vbCrLf
                s = s & "ARCustomer, Description, Inactive, Contact1, Phone1, Fax1, Email1, Contact2, Phone2, Fax2, Email2" & vbCrLf
                s = s & ", BillAddr1, BillAddr2, BillCity, BillProvince, BillPostalCode" & vbCrLf
                s = s & ", Comments, ExternalID,  DivisionID)" & vbCrLf
                s = s & "select" & vbCrLf
                s = s & " c.Customer_No, c.Description, 0," & vbCrLf
                s = s & "c.customer_lname + ', ' + c.customer_name Contact1 , c.phone,c.fax,c.email," & vbCrLf
                s = s & "c.cobuyer_lname + ', ' + c.cobuyer_name Contact2 , c.c_phone,c.c_fax,c.c_email," & vbCrLf
                s = s & "c.address1,c.address2,c.city,c.province,c.zip," & vbCrLf
                s = s & "c.comments,c.customer_no,c.divisionid" & vbCrLf
                s = s & "from tblCustomers c" & vbCrLf
                s = s & "left join arcustomers a on c.customer_no=a.arcustomer and c.divisionid=a.divisionid" & vbCrLf
                s = s & "where a.arcustomer is null" & vbCrLf
                s = s & "and c.divisionid=" & HFApp.DivisionID & vbCrLf
                s = s & "and c.job_no=" & DbQuote(Str, mJob) & vbCrLf
                s = s & "" & vbCrLf
                s = s & "update tblcustomers set ar_customer_deposit=customer_no" & vbCrLf
                s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
                s = s & "and job_no=" & DbQuote(Str, mJob) & vbCrLf
                HFApp.SqlExec s
                
                s = "select customer_no from tblcustomers" & vbCrLf
                s = s & "where divisionid=" & HFApp.DivisionID & vbCrLf
                s = s & "and job_no=" & DbQuote(Str, mJob) & vbCrLf
                txtARCustomer.Text = HFApp.SqlExec(s)(0)
            Else
        
                txtARCustomer.Text = FPickList.SelectedItem("ExternalID")
                Dirty = True
            End If
    
    End Select
End Sub

Private Sub cmdJob_Click()
On Error GoTo eh
    Dim i As Long
    Dim s As String
    Dim SimpleQB As Boolean
    Dim Intacct As IntacctWrapper.IntacctWrapper
    
    If cmdJob.Visible = False Then Exit Sub
    
    SetCtrlFocus txtJob
    
    Select Case HFApp.Options(AccountingSystem)
    

        Case asQuickbooksOnline
            s = SubmitQBOXml("GetCustomers", "")
            s = "exec QBO_Jobs " & DbQuote(Str, s) & " -- no order by"  '<-- no order by is required by fpicklist

            SimpleQB = HFApp.Options.ValueByName("QuickbooksJobStyle") = "Simple"

            If FPickList.Choose(HFApp.Databases(dbHomefront), "Job", s) Then
                mDirty = True
                Dirty = True
                If txtJob.Text = "" Then
                    txtJob.Text = FPickList.SelectedItem("JobID")
                End If
                txtExternalID.Text = FPickList.SelectedItem("JobID")
                txtARCustomer.Text = FPickList.SelectedItem("CustomerID")
                txtDescription.Text = FPickList.SelectedItem("Description")

            End If
        
        
        Case asSimply
            Call HFApp.OpenSimplyODBC
            s = "select lid Job,sname Description from tproject"
            If FPickList.Choose(HFApp.Databases(dbAccounting), "Job Number", s) Then
                mDirty = True
                If txtJob.Text = "" Then txtJob.Text = FPickList.SelectedItem("Job")
                txtExternalID.Text = FPickList.SelectedItem("Job")
                txtDescription.Text = FPickList.SelectedItem("Description")
            End If
            Call HFApp.CloseSimplyODBC
    
        Case asSpectrum
            s = ""
            s = s & "select rtrim(ltrim(job_number)) Job, job_description Description " & vbCrLf
            s = s & "from jc_job_master_mc" & vbCrLf
            s = s & "where company_code=" & DbQuote(Str, HFApp.Options.ValueByName("SpectrumCompany"))
            If FPickList.Choose(HFApp.Databases(dbAccounting), "Job Number", s) Then
                mDirty = True
                txtJob.Text = FPickList.SelectedItem("Job")
                txtDescription.Text = FPickList.SelectedItem("Description")
            End If
            
        Case asTimberline
            s = "select Job Job,jdesc Description from master_jcm_record_1_1"
            If FPickList.Choose(HFApp.Databases(dbAccountingDictionary), "Job Number", s) Then
                mDirty = True
                txtJob.Text = FPickList.SelectedItem("Job")
                txtDescription.Text = FPickList.SelectedItem("Description")
            End If
            
        Case asMasterBuilder
            s = "select recnum Job,jobnme Description from actrec"
            If FPickList.Choose(HFApp.Databases(dbAccountingDictionary), "Job Number", s) Then
                mDirty = True
                txtJob.Text = FPickList.SelectedItem("Job")
                txtDescription.Text = FPickList.SelectedItem("Description")
            End If
        
            
        Case asIntacct
            Set Intacct = New IntacctWrapper.IntacctWrapper
            Call Intacct.OpenMessage(HFApp.Options.ValueByName("IntacctCompanyID"), HFApp.Options.ValueByName("IntacctUID"), HFApp.Options.ValueByName("IntacctPWD"), HFApp.Options.ValueByName("IntacctEntity"))
            Call Intacct.GetJobs("")
            Call Intacct.CloseMessage
            Call WriteLogFile("intacct.getjobs.req.xml", Intacct.xml())
            s = Intacct.PostMessage(True)
            Call WriteLogFile("intacct.getjobs.res.xml", s)
            s = "exec Intacct_Jobs " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, s) & " -- no order by" '<-- no order by is required by fpicklist
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Job Number", s) Then
                mDirty = True
                txtJob.Text = FPickList.SelectedItem("Job")
                txtDescription.Text = FPickList.SelectedItem("Name")
            End If
        
        
        
        
        
        Case asQuickBooks
            s = HFApp.XmlQBStart
            s = s & "<CustomerQueryRq metaData=""MetaDataAndResponseData"">" & vbCrLf
            s = s & "</CustomerQueryRq>" & vbCrLf
            s = s & HFApp.XmlQBEnd
            s = HFApp.XmlQBSubmit(s)
            s = "exec QB_Jobs " & DbQuote(Num, HFApp.DivisionID) & "," & DbQuote(Str, s) & " -- no order by" '<-- no order by is required by fpicklist
            
            SimpleQB = HFApp.Options.ValueByName("QuickbooksJobStyle") = "Simple"
            
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Job Number", s, , , , , "CustomerNumber,JobComments" & IIf(SimpleQB, ",Customer", "")) Then
                mDirty = True
                Dirty = True
                If txtJob.Text = "" Then
                    txtJob.Text = FPickList.SelectedItem("JobNumber")
                End If
                txtExternalID.Text = FPickList.SelectedItem("JobNumber")
                txtARCustomer.Text = FPickList.SelectedItem("CustomerNumber")
                If txtDescription.Text = "" Then
                    txtDescription.Text = FPickList.SelectedItem("Job")
                End If
                If txtNotes.Text = "" Then
                    txtNotes.Text = FPickList.SelectedItem("JobComments")
                End If
                
            End If
    End Select
    
    
Exit Sub
eh: Call errHandler(SRCFILE & "cmdJob_Click")
End Sub


Private Sub Form_Click()
Dim i As Long
For i = 0 To 12: gPOs.ColHidden(i) = False: Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case True
        Case Shift = vbCtrlMask And KeyCode = vbKeyO:  Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        Case Shift = vbCtrlMask And KeyCode = vbKeyS:  Call Toolbar_ButtonClick(Toolbar.Buttons("Save"))
    End Select
End Sub

Private Sub LoadProvinces()

    Const provinces = "AB,BC,MB,NB,NL,NS,NT,NU,ON,PE,QC,SK,YT,AK,AL,AR,AZ,CA,CO,CT,DC,DE,FL,GA,GU,HI,IA,ID,IL,IN,KS,KY,LA,MA,MD,ME,MI,MN,MO,MS,MT,NC,ND,NE,NH,NJ,NM,NV,NY,OH,OK,OR,PA,PR,RI,SC,SD,TN,TX,UT,VA,VI,VT,WA,WI,WV,WY"
    Const states = "AK,AL,AR,AZ,CA,CO,CT,DC,DE,FL,GA,GU,HI,IA,ID,IL,IN,KS,KY,LA,MA,MD,ME,MI,MN,MO,MS,MT,NC,ND,NE,NH,NJ,NM,NV,NY,OH,OK,OR,PA,PR,RI,SC,SD,TN,TX,UT,VA,VI,VT,WA,WI,WV,WY,AB,BC,MB,NB,NL,NS,NT,NU,ON,PE,QC,SK,YT"
    Const AUstates = "ACT,NSW,NT,Qld,SA,Tas,Vi,WA"

    Dim s As String
    Dim i As Long

    cboProvince.Clear
    s = IIf(mIsCanadian, provinces, states)
    If mIsAustralia Then s = AUstates
    For i = 1 To Parse(s)
        cboProvince.AddItem Parse(s, i)
    Next
    
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim s As String
    
    Call SetToolbarIcons(Toolbar, LargeIcons)
    
    Call IniGetForm(Me)
    Call IniGetGrid(Me, gContacts)
    Call IniGetGrid(Me, gJobContacts)
    Call IniGetGrid(Me, gPOs)
    cboStatus.Clear
    cboStatus.AddItem "In progress"
    cboStatus.AddItem "Closed"
    
    lblCommunity.Caption = GetCustomDesc("Community")
    
    mIsCanadian = HFApp.Options(Country) = "CA"
    mIsAustralia = HFApp.Options(Country) = "AU"
    
    Select Case HFApp.Options(AccountingSystem)
        
        Case asQuickBooks
            mIsCanadian = HFApp.Options.ValueByName("AccountingVersion") = "CA"
            lblExternalID.Caption = "Quickbooks ID"
            
        Case asSimply:         lblExternalID.Caption = "Simply ID"
        Case asMYOB:           lblExternalID.Caption = "MYOB ID"
        Case asPeachtree:      lblExternalID.Caption = "Peachtree ID"
        Case asSpectrum:       lblExternalID.Caption = "Spectrum ID"
        
        Case asMasterBuilder
            mIsCanadian = HFApp.Options.ValueByName("MasterBuilderEdition") = "CA"
            cboProvince.Style = vbComboDropdownList 'Sage100 enforces this list so limit picklist

            lblExternalID.Visible = False
            txtExternalID.Visible = False
            'arcustomer
            cmdBrowse(10).Visible = False
            txtARCustomer.Visible = False
            lblLink(0).Visible = False
        
        Case Else
            lblExternalID.Visible = False
            txtExternalID.Visible = False
            
    End Select
    
    Call LoadProvinces
    Call LoadCommunities
    
'changed feb 2016 MK
'the model box was loaded using the ItemKey mode.
'    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct isnull(model,'') + isnull(' - ' + description,''),isnull(model,'') ,0 from tbldbassemblymaster where assemblytype=0 order by 1")
'then the loaddata routine calls this, but it was very slow on very large datasets (3000+ items).
'    Call SetComboBoxListIndex(cboModel, , "" & rs("Model"))
'both procs changed to use ItemDesc mode
    s = ""
    s = s & "select isnull(model,'') + isnull(' - ' + description,''),'' ,0" & vbCrLf
    s = s & "from tbldbassemblymaster" & vbCrLf
    s = s & "where assemblytype=0" & vbCrLf
    s = s & "and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
    s = s & "order by 1" & vbCrLf
    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), s)
    
    
    Call LoadGLPrefixes
    
    If HFApp.Options(AccountingSystem) = asTimberline Then
        txtJob.MaxLength = 10 + Len(Replace(HFApp.Options(Job_Mask), "&", ""))
    Else
        txtJob.MaxLength = 12
    End If
    
    If mJob = "" Then
        mDirty = False
        Call Toolbar_ButtonClick(Toolbar.Buttons("Open"))
        If mJob = "" Then Unload Me
    Else
    
        Call LoadData
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Const margin = 120
    txtNotes.Width = Me.ScaleWidth - margin - gContacts.left
    gContacts.Width = Me.ScaleWidth - margin - gContacts.left
    gJobContacts.Width = Me.ScaleWidth - margin - gContacts.left
    gProperties.Height = Me.ScaleHeight - margin - gProperties.Top
    gPOs.Move gPOs.left, gPOs.Top, Me.ScaleWidth - gPOs.left - margin, Me.ScaleHeight - gPOs.Top - margin
End Sub

Private Sub gContacts_DblClick()
    Call HFApp.EditProjectManagerList
End Sub

Private Sub gContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    With gContacts
        If KeyCode = vbKeyDelete And .Row > 0 And IsIn(.ColKey(.Col), "PM", "PMName") Then
        
            .TextMatrix(.Row, .ColIndex("PM")) = ""
            .TextMatrix(.Row, .ColIndex("PMName")) = ""
            .TextMatrix(.Row, .ColIndex("Phone")) = ""
            .TextMatrix(.Row, .ColIndex("Cell")) = ""
            .TextMatrix(.Row, .ColIndex("Fax")) = ""
            .TextMatrix(.Row, .ColIndex("Email")) = ""
            .RowData(.Row) = "DIRTY"
            Dirty = True
        End If
    End With
End Sub

Private Sub gPOs_DblClick()
    Dim s As String
    Dim rs As Recordset
    'Dim v As New HFPrinter.ReportViewer
    Dim PONumber As String
    
If IsIn(gPOs.RowData(gPOs.Row), "Pending") Then Exit Sub
    
    
    PONumber = gPOs.TextMatrix(gPOs.Row, gPOs.ColIndex("PONumber"))
    
    s = ""
    s = s & "SELECT rpt = CASE " & vbCrLf
    s = s & "               WHEN ISNULL(v.poformat,'')<>'' THEN v.poformat" & vbCrLf
    s = s & "               WHEN ISNULL(i.poformat,'')<>'' THEN i.poformat" & vbCrLf
    s = s & "               ELSE " & DbQuote(Str, HFApp.Options(POFormat)) & vbCrLf
    s = s & "             END" & vbCrLf
    s = s & "  FROM POMaster p" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblPOIndex i ON(p.DivisionID = i.DivisionID and p.POIndex=i.POIndex)" & vbCrLf
    s = s & "       LEFT OUTER JOIN tblVendors v ON(p.Vendor=v.Vendor_id and v.Divisionid = " & HFApp.DivisionID & ")" & vbCrLf
    s = s & " WHERE p.DivisionID = " & HFApp.DivisionID & " and p.PONumber=" & DbQuote(Str, PONumber) & vbCrLf
    Set rs = HFApp.SqlExec(s)
    
    Screen.MousePointer = vbHourglass
    s = PathAppend(HFApp.SystemFolder, "Estimating\PO Formats", "" & rs(0) & ".rpt")
    'Call v.ShowReport(HFApp.ConnectionString(dbHomefront), s, rvPreview, "", "", "PONumber", PONumber, "DivisionID", HFApp.DivisionID)
    Dim c As New ZybUtil.Crystal
    Call c.LoadODBCReport(s, HFApp.LoginDSN, HFApp.LoginDBUID, HFApp.LoginDBPWD)
    On Error Resume Next
    Call c.ParameterValue("DivisionID", HFApp.DivisionID)
    Call c.ParameterValue("PONumber", PONumber)
    On Error GoTo 0
    Call c.PrintPreview("Print Preview")
    
    Screen.MousePointer = vbDefault
    
End Sub






Private Sub gProperties_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    FinishEdit = True
End Sub

Private Sub lblLink_Click(Index As Integer)
    If Not lblLink(Index).Enabled Then Exit Sub
    Select Case Index
        Case 0:       If txtARCustomer.Text <> "" Then HFApp.RunTask ("EditCustomer|" & txtARCustomer.Text)
    End Select
End Sub



Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Dim f As Form
    Dim s As String
    
    Select Case Button.Key
    
        Case "JobDocs"
            If mJob <> "" Then
                Call FJobCorrespondence.ShowForm(mJob)
            End If
            
        Case "Attachments"
            'Call FAtt.ShowForm("File Attachments - " & txtJob.Text & " - " & txtDescription.Text, "J~" & mJob, "Job Documents")
            Call HFApp.RunTask("EditAttachments|J~" & mJob & "|Job Documents")
            
        Case "Open"
            If Not SaveData(True) Then Exit Sub
            
            s = "SELECT j.Job_No Job,j.Description,c.Description Community,j.CommunityPhase Phase,j.Municipal_Address Address,j.PM,j.Purchaser,j.Estimator FROM tblJobs j LEFT OUTER JOIN tblLocality c on(j.community=c.area) "
            If HFApp.DivisionID <> "" Then
                 's = s & " left outer join DivisionCommunities d on d.Community = j.Community"
                 s = s & " where j.DivisionID  = " & HFApp.DivisionID
            Else
                s = s & " where j.DivisionID = " & HFApp.DivisionID & " and j.isquote=0"
            End If
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Job", s, mJob, , True) Then
            
                mJob = FPickList.SelectedItem("Job")
                If mJob = "" Then mJob = Chr(1) 'new job
                Set f = FindForm("FJob", mJob)
                
                If f Is Nothing Then
                    Call LoadData
                Else
                    If f.WindowState = vbMinimized Then f.WindowState = vbNormal
                    f.SetFocus
                    If Not Toolbar.Buttons(3).Enabled Then
                        'close empty screen
                        Unload Me
                    End If
                End If
            End If
            
            
         Case "Save"
            Call SaveData(False)
            
            
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetCtrlFocus txtNotes
    If Not SaveData(True) Then
        Cancel = True
        Exit Sub
    End If
    Call IniPutForm(Me)
    Call IniPutGrid(Me, gContacts)
    Call IniPutGrid(Me, gJobContacts)
    Call IniPutGrid(Me, gPOs)
End Sub

Public Function SaveData(prompt As Boolean) As Boolean
On Error GoTo eh
    
    Dim GLPrefix As String
    Dim s As String
    Dim Community As String
    Dim CommunityPhase As String
    Dim model As String
    Dim r As Long
    Dim rs As Recordset
    Dim i As Long
    
    If Not Dirty Then
        SaveData = True
        Exit Function
    End If
    If prompt Then
        Select Case MsgBox(Me.Caption & " has changed." & vbCrLf & vbCrLf & "Do you want to save these changes?" & vbCrLf, vbExclamation + vbYesNoCancel, App.ProductName)
            Case vbNo
                SaveData = True
                Exit Function
            Case vbCancel
                SaveData = False
                Exit Function
        End Select
    End If
    
    txtJob.Text = Trim(txtJob.Text)
    If Trim(txtJob.Text) = "" Then
        MsgBox "A job number is required.", vbExclamation, App.ProductName
        Call SetCtrlFocus(txtJob)
        Exit Function
    End If
    
    
    
    
    
    Screen.MousePointer = vbHourglass
    
    

    Community = GetComboBoxListKey(cboCommunity)
    CommunityPhase = cboPhase.Text
    GLPrefix = IIf(cboGLPrefix.Visible, GetComboBoxListKey(cboGLPrefix), txtGLPrefix.Text)
    
'changed feb 2016 MK
'the model box was loaded using the ItemKey mode.
'    Call LoadComboBox(cboModel, HFApp.Databases(dbHomefront), "select distinct isnull(model,'') + isnull(' - ' + description,''),isnull(model,'') ,0 from tbldbassemblymaster where assemblytype=0 order by 1")
'then the loaddata routine calls this, but it was very slow on very large datasets (3000+ items).
'    Call SetComboBoxListIndex(cboModel, , "" & rs("Model"))
'both procs changed to use ItemDesc mode
'save used to use get combolistkey
'    model = GetComboBoxListKey(cboModel)
'now changed to parse model from description
    model = Parse(cboModel.Text, 1, " - ")
    
    
    
    'save job
    With gProperties
        If mJob = Chr(1) Then
            If Not ValidateJobNumber() Then
                Screen.MousePointer = vbDefault
                Call SetCtrlFocus(txtJob)
                Exit Function
            End If
            
            s = ""
            s = s & "INSERT INTO tblcustomers (home_selection,presale_selection,DivisionID,Customer_No,Job_no,ar_customer_deposit,Description,bal_sheet_Prefix,community,phase,address1,city,province,zip,phone,fax,purchased,approved,contract_assigned,cancelled,inactive,lot,block,lotplan,model,sale_posted)" & vbCrLf
            s = s & "VALUES('Spec Home','Spec Home'," & HFApp.DivisionID & vbCrLf
            s = s & "      ," & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            s = s & "      ," & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            s = s & "      ," & DbQuote(Str, CleanJob(txtARCustomer.Text)) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, GLPrefix) & vbCrLf
            s = s & "      ," & DbQuote(Str, Community) & vbCrLf
            s = s & "      ," & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtAddress.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtCity.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, cboProvince.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPostal.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtPhone.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtFax.Text) & ",1,1,1,0,0"
            s = s & "      ," & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "      ," & DbQuote(Str, model) & vbCrLf
            s = s & "      ,1)"
            HFApp.SqlExec s
            'tblcustomer trigger will create the tbljob record
            
            
            s = ""
            s = s & "update tblJobs" & vbCrLf
            s = s & "set isquote=0" & vbCrLf
            s = s & "   ,ExternalJobID=" & DbQuote(Str, txtExternalID.Text) & vbCrLf
            s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
            s = s & "   ,GL_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
            s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
            s = s & "   ,CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "   ,municipal_address=" & DbQuote(Str, txtAddress.Text) & vbCrLf
            s = s & "   ,city=" & DbQuote(Str, txtCity.Text) & vbCrLf
            s = s & "   ,province=" & DbQuote(Str, cboProvince.Text) & vbCrLf
            s = s & "   ,zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
            s = s & "   ,sitephone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
            s = s & "   ,sitefax=" & DbQuote(Str, txtFax.Text) & vbCrLf
            s = s & "   ,Inactive=" & DbQuote(Bit, cboStatus.ListIndex = 1) & vbCrLf
            s = s & "   ,arcustomer=" & DbQuote(Str, CleanJob(txtARCustomer.Text)) & vbCrLf
            s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "   ,Model=" & DbQuote(Str, model) & vbCrLf
            
            'properties
            If HFApp.Options.ValueByName("EnableTarionFields") = "true" Then
                s = s & "   ,TarionEnrollmentNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionEnrollmentNumber, .ColIndex("value"))) & vbCrLf
                s = s & "   ,TarionBuilderNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionBuilderNumber, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,PermitNumber=" & DbQuote(Str, .TextMatrix(mprop_PermitNumber, .ColIndex("value"))) & vbCrLf
            s = s & "   ,PermitReceivedDate=" & DbQuote(Date, .TextMatrix(mprop_PermitDate, .ColIndex("value"))) & vbCrLf
            
            If mIsMultifamily Then
                s = s & "   ,ShellScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_ShellTemplate, .ColIndex("value"))) & vbCrLf
            End If
            
            If HFApp.Options.ValueByName("EnablePreconJobs") = "true" Then
                s = s & "   ,preconScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_PreconTemplate, .ColIndex("value"))) & vbCrLf
                s = s & "   ,PreconStartDate=" & DbQuote(Date, .TextMatrix(mprop_PreconStart, .ColIndex("value"))) & vbCrLf
            End If
            
            s = s & "   ,ScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_UnitTemplate, .ColIndex("value"))) & vbCrLf
            s = s & "   ,Start_Date=" & DbQuote(Date, .TextMatrix(mprop_ConstStart, .ColIndex("value"))) & vbCrLf
            If HFApp.Options(AccountingSystem) = asIntacct Then
                s = s & "   ,IntacctDepartment=" & DbQuote(Str, .TextMatrix(mprop_IntacctDepartment, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_LabTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_MatTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_SubTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_EquTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OvrTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OthTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,ARTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_ARTax, .ColIndex("value"))) & vbCrLf
            
            s = s & "WHERE DivisionID=" & HFApp.DivisionID & vbCrLf
            s = s & "  AND job_no=" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            HFApp.SqlExec s
            mJob = CleanJob(txtJob.Text)
            Me.tag = mJob
            txtJob.Enabled = False
            cmdJob.Visible = False
            txtJob.Width = 2865
        
        
            
            Call HFApp.WriteJobToAccounting(CleanJob(txtJob.Text))
            s = ""
            s = s & "select externaljobid" & vbCrLf
            s = s & "from tbljobs" & vbCrLf
            s = s & "where job_no=" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            s = s & "and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            Set rs = HFApp.SqlExec(s)
            If Not rs.EOF Then
                txtExternalID.Text = "" & rs(0)
            End If
            
            s = ""
            s = s & "INSERT INTO EstimateAssemblies(job,Customer_no,EstimateIndex,HFDescription,AssemblyType,OptionType,DivisionID,salesqty)" & vbCrLf
            s = s & "VALUES(" & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            s = s & "      ," & DbQuote(Str, CleanJob(txtJob.Text)) & vbCrLf
            s = s & "      ,0,'Manual Estimates',-1,-1," & HFApp.DivisionID & ",1)" & vbCrLf
            HFApp.SqlExec s
        
        Else
        
            s = ""
            s = s & "UPDATE tblJobs" & vbCrLf
            s = s & "SET Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
            s = s & "   ,Notes=" & DbQuote(Str, txtNotes.Text) & vbCrLf
            s = s & "   ,GL_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
            s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
            s = s & "   ,CommunityPhase=" & DbQuote(Str, CommunityPhase) & vbCrLf
            s = s & "   ,Municipal_Address=" & DbQuote(Str, txtAddress.Text) & vbCrLf
            s = s & "   ,City=" & DbQuote(Str, txtCity.Text) & vbCrLf
            s = s & "   ,Province=" & DbQuote(Str, cboProvince.Text) & vbCrLf
            s = s & "   ,Zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
            s = s & "   ,SitePhone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
            s = s & "   ,SiteFax=" & DbQuote(Str, txtFax.Text) & vbCrLf
            s = s & "   ,Inactive=" & DbQuote(Bit, cboStatus.ListIndex = 1) & vbCrLf
            s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
            s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
            s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
            s = s & "   ,Model=" & DbQuote(Str, model) & vbCrLf
            
            'properties
            If HFApp.Options.ValueByName("EnableTarionFields") = "true" Then
                s = s & "   ,TarionEnrollmentNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionEnrollmentNumber, .ColIndex("value"))) & vbCrLf
                s = s & "   ,TarionBuilderNumber=" & DbQuote(Str, .TextMatrix(mprop_TarionBuilderNumber, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,PermitNumber=" & DbQuote(Str, .TextMatrix(mprop_PermitNumber, .ColIndex("value"))) & vbCrLf
            s = s & "   ,PermitReceivedDate=" & DbQuote(Date, .TextMatrix(mprop_PermitDate, .ColIndex("value"))) & vbCrLf
            If mIsMultifamily Then
                s = s & "   ,ShellScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_ShellTemplate, .ColIndex("value"))) & vbCrLf
            End If
            
            If HFApp.Options.ValueByName("EnablePreconJobs") = "true" Then
                s = s & "   ,preconScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_PreconTemplate, .ColIndex("value"))) & vbCrLf
                s = s & "   ,PreconStartDate=" & DbQuote(Date, .TextMatrix(mprop_PreconStart, .ColIndex("value"))) & vbCrLf
            End If
            
            s = s & "   ,ScheduleTemplate=" & DbQuote(Str, .TextMatrix(mprop_UnitTemplate, .ColIndex("value"))) & vbCrLf
            s = s & "   ,Start_Date=" & DbQuote(Date, .TextMatrix(mprop_ConstStart, .ColIndex("value"))) & vbCrLf
            If HFApp.Options(AccountingSystem) = asIntacct Then
                s = s & "   ,IntacctDepartment=" & DbQuote(Str, .TextMatrix(mprop_IntacctDepartment, .ColIndex("value"))) & vbCrLf
            End If
            s = s & "   ,LabourTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_LabTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,MaterialTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_MatTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,SubContractTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_SubTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,EquipmentTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_EquTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OverheadTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OvrTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,OtherTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_OthTax, .ColIndex("value"))) & vbCrLf
            s = s & "   ,ARTaxGroup=" & DbQuote(Str, .TextMatrix(mprop_ARTax, .ColIndex("value"))) & vbCrLf
            
            If Not HFApp.Options(MultiFamily) Then
                s = s & "   ,arcustomer=" & DbQuote(Str, CleanJob(txtARCustomer.Text)) & vbCrLf
            End If
            
            'Added for Quickbooks users.
            If ("" & txtARCustomer.Text <> "" Or "" & txtExternalID.Text <> "") And IsIn(HFApp.Options(AccountingSystem), asQuickBooks, asQuickbooksOnline) Then
                If HFApp.Options.ValueByName("QuickBooksJobStyle") = "Simple" Then
                    If txtARCustomer.Text <> "" Then
                        s = s & "   ,ExternalJobID=" & DbQuote(Str, txtARCustomer.Text) & vbCrLf
                    End If
                Else
                    If txtExternalID.Text <> "" Then
                        s = s & "   ,ExternalJobID=" & DbQuote(Str, txtExternalID.Text) & vbCrLf
                    End If
                End If
            End If
            
            
            If "" & txtExternalID.Text <> "" And (HFApp.Options(AccountingSystem) = asSimply Or HFApp.Options(AccountingSystem) = asMYOB) Then
                s = s & "   ,ExternalJobID=" & DbQuote(Str, txtExternalID.Text) & vbCrLf
            End If
            
            s = s & "WHERE Job_No=" & DbQuote(Str, CleanJob(txtJob.Text)) & " and DivisionID = " & HFApp.DivisionID & vbCrLf
            HFApp.SqlExec s
            
            
            If Not HFApp.Options(MultiFamily) Then
                s = ""
                s = s & "UPDATE tblcustomers" & vbCrLf
                s = s & "SET ar_customer_deposit=" & DbQuote(Str, txtARCustomer.Text) & vbCrLf
                s = s & "WHERE Job_No=" & DbQuote(Str, CleanJob(txtJob.Text)) & " and DivisionID = " & HFApp.DivisionID
                HFApp.SqlExec s
            End If
            
            If HFApp.Options(SalesSystem) = SalesSystems.asNone Then
            
                s = ""
                s = s & "UPDATE tblcustomers" & vbCrLf
                s = s & "SET ar_customer_deposit=" & DbQuote(Str, txtARCustomer.Text) & vbCrLf
                s = s & "   ,Description=" & DbQuote(Str, txtDescription.Text) & vbCrLf
                s = s & "   ,bal_sheet_Prefix=" & DbQuote(Str, GLPrefix) & vbCrLf
                s = s & "   ,Community=" & DbQuote(Str, Community) & vbCrLf
                s = s & "   ,Phase=" & DbQuote(Str, CommunityPhase) & vbCrLf
                s = s & "   ,Address1=" & DbQuote(Str, txtAddress.Text) & vbCrLf
                s = s & "   ,City=" & DbQuote(Str, txtCity.Text) & vbCrLf
                s = s & "   ,Province=" & DbQuote(Str, cboProvince.Text) & vbCrLf
                s = s & "   ,Zip=" & DbQuote(Str, txtPostal.Text) & vbCrLf
                s = s & "   ,Phone=" & DbQuote(Str, txtPhone.Text) & vbCrLf
                s = s & "   ,Fax=" & DbQuote(Str, txtFax.Text) & vbCrLf
                s = s & "   ,Lot=" & DbQuote(Str, txtLot.Text) & vbCrLf
                s = s & "   ,Block=" & DbQuote(Str, txtBlock.Text) & vbCrLf
                s = s & "   ,LotPlan=" & DbQuote(Str, txtLotPlan.Text) & vbCrLf
                s = s & "   ,Model=" & DbQuote(Str, model) & vbCrLf
                s = s & "WHERE Customer_No=" & DbQuote(Str, CleanJob(txtJob.Text)) & " and DivisionID = " & HFApp.DivisionID
                HFApp.SqlExec s
                

            End If
            
            Call HFApp.WriteJobToAccounting(mJob)
            
        End If
        
    
        'save properties
        s = ""
        For r = mprop_UserFlds To .Rows - 1
            If .TextMatrix(r, .ColIndex("DataType")) <> "" Then
                s = s & "," & vbQuote & .TextMatrix(r, .ColIndex("Name")) & vbQuote & "=" & DbQuote(.ValueMatrix(r, .ColIndex("DataType")), .TextMatrix(r, .ColIndex("Value")))
            End If
        Next
        If s <> "" Then
            On Error Resume Next
            Call HFApp.SqlExec("INSERT INTO JobCustomFields(Job_No,divisionid) VALUES(" & DbQuote(Str, CleanJob(mJob)) & "," & DbQuote(Num, HFApp.DivisionID) & ")", dbHomefront)
            On Error GoTo eh
            
            s = "UPDATE JobCustomFields SET " & Mid(s, 2) & " WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and Job_No=" & DbQuote(Str, CleanJob(mJob))
            Call HFApp.SqlExec(s, dbHomefront)
        End If
    End With
    
    
    'save contacts
    With gContacts
        For r = 1 To .Rows - 1
            s = "UPDATE tblJobs" & vbCrLf & _
                "SET " & .TextMatrix(r, .ColIndex("ContactType")) & "=" & DbQuote(Str, .TextMatrix(r, .ColIndex("PM"))) & vbCrLf & _
                "WHERE divisionid=" & DbQuote(Num, HFApp.DivisionID) & " and Job_No=" & DbQuote(Str, CleanJob(mJob))

            Call HFApp.SqlExec(s, dbHomefront)
        Next
    End With
    
    
    Call SaveJobContacts
        
    SaveData = True
    Dirty = False
    Screen.MousePointer = vbDefault
    
Exit Function
eh: If InStr(1, Err.Description, "duplicate", vbTextCompare) Then
        Screen.MousePointer = vbDefault
        MsgBox "Unable to save this job. " & vbQuote & txtJob.Text & vbQuote & " has already been used." & vbCrLf & vbCrLf & "Please enter a different job number.", vbExclamation, App.ProductName
    Else
        Call errHandler(SRCFILE & "SaveData", s)
    End If
End Function


Private Sub LoadData()
On Error GoTo eh
    Dim s As String
    Dim rs As adodb.Recordset
    Dim r As Long
    Dim b As Boolean
    Dim ra As Long
    Dim model As String
    
    Me.tag = mJob
    
    
    'disable arcustomer if you are multifamility
    b = Not HFApp.Options(MultiFamily)
    If Not b Then
        'may already be hidden at startup
        cmdBrowse(10).Visible = b
        txtARCustomer.Visible = b
        lblLink(0).Visible = b
    End If
    
    If mJob = Chr(1) Then 'newjob
        txtJob.Enabled = True
        b = True
    Else
        txtJob.Enabled = False
        s = ""
        s = s & "select j.*" & vbCrLf
        s = s & "      ,c.ar_customer_deposit" & vbCrLf
        s = s & "  from tbljobs j" & vbCrLf
        s = s & "  left outer join tblcustomers c on(j.job_no=c.job_no)" & vbCrLf
        s = s & " where j.DivisionID = " & HFApp.DivisionID & " and j.job_no=" & DbQuote(Str, mJob) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        
        If rs.EOF Then
            b = False
        Else
            b = True
            s = ""
            s = s & "select isnull(model,'') + isnull(' - ' + description,''),'' ,0" & vbCrLf
            s = s & "from tbldbassemblymaster" & vbCrLf
            s = s & "where assemblytype=0" & vbCrLf
            s = s & "and divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
            s = s & "and model=" & DbQuote(Str, "" & rs("model"))
            On Error Resume Next
            model = HFApp.SqlExec(s)(0)
            On Error GoTo eh
        End If
    End If
    
    'if job number is editable and we have an accounting system then show picklist button
    Select Case HFApp.Options(AccountingSystem)
        Case AccountingSystems.asNone
            cmdJob.Visible = False
            txtJob.Width = 2865
        
        Case asQuickBooks, asSimply, asMYOB, asSpectrum, asIntacct, asQuickbooksOnline
            cmdJob.Visible = True
            txtJob.Width = 2640
        
        Case asTimberline, asMasterBuilder
            cmdJob.Visible = txtJob.Enabled
            txtJob.Width = IIf(txtJob.Enabled, 2640, 2865)
    End Select
    
    
    
    
    txtDescription.Enabled = b
    txtNotes.Enabled = b
    cboGLPrefix.Enabled = b
    txtGLPrefix.Enabled = b
    txtAddress.Enabled = b
    txtCity.Enabled = b
    cboProvince.Enabled = b
    txtPostal.Enabled = b
    txtPhone.Enabled = b
    txtFax.Enabled = b
    cboStatus.Enabled = b And HFApp.UserPermission("ChangeJobStatuses")
    cboModel.Enabled = b
    txtLot.Enabled = b
    txtBlock.Enabled = b
    txtLotPlan.Enabled = b
    gContacts.Enabled = b
    gProperties.Enabled = b
    
    If mJob = Chr(1) Or Not b Then
        SetCtrlFocus txtJob
        Me.Caption = "Job Setup"
        txtJob.Text = ""
        txtARCustomer.Text = ""
        txtExternalID.Text = ""
        txtDescription.Text = ""
        txtNotes.Text = ""
        txtGLPrefix.Text = ""
        cboGLPrefix.ListIndex = -1
        txtAddress.Text = ""
        txtCity.Text = ""
        cboProvince.ListIndex = -1
        txtPostal.Text = ""
        
        txtPhone.Text = ""
        txtFax.Text = ""
        cboModel.Text = ""
        txtLot.Text = ""
        txtBlock.Text = ""
        txtLotPlan.Text = ""
        cboStatus.ListIndex = 0
        cboCommunity.ListIndex = -1
    Else
        SetCtrlFocus txtDescription
        txtJob.Text = HFApp.FormatJob("" & rs("Job_No"))
        txtExternalID.Text = "" & rs("ExternalJobID")
        Me.Caption = "Job Setup - " & txtJob.Text
        txtDescription.Text = "" & rs("Description")
        txtNotes.Text = "" & rs("Notes")
        
        If HFApp.Options(MultiFamily) = False Then
            txtARCustomer.Text = "" & rs("ar_customer_deposit")
        Else
            txtARCustomer.Text = "" & rs("ARCustomer")
        End If
        
        'cboGLPrefix.Clear
        Call SetComboBoxListIndex(cboGLPrefix, , "" & rs("gl_prefix"))
        txtGLPrefix.Text = "" & rs("gl_prefix")
        txtAddress.Text = "" & rs("municipal_address")
        txtCity.Text = "" & rs("city")
        Call SetComboBoxListIndex(cboProvince, "" & rs("province"))
        On Error Resume Next
        If cboProvince.ListIndex = -1 Then cboProvince.Text = "" & rs("province")
        On Error GoTo 0
        
        txtPostal.Text = "" & rs("zip")
        txtPhone.Text = FormatPhone("" & rs("sitephone"))
        txtFax.Text = FormatPhone("" & rs("sitefax"))
        Call SetComboBoxListIndex(cboModel, model)
        txtLot.Text = "" & rs("Lot")
        txtBlock.Text = "" & rs("Block")
        txtLotPlan.Text = "" & rs("LotPlan")
        
        cboStatus.ListIndex = IIf("" & rs("inactive") = "true", 1, 0)
        Call SetComboBoxListIndex(cboCommunity, "", "" & rs("Community"))
        cboPhase.Clear
        cboPhase.Text = "" & rs("CommunityPhase")

    End If

    Call LoadProperties
    Call LoadPOs
    'load contacts
    s = ""
    s = s & "SELECT 'PM' ContactType,p.*" & vbCrLf
    s = s & "  FROM system_setup s" & vbCrLf
    s = s & "       left outer join tblJobs j on(Job_No=" & DbQuote(Str, mJob) & " and j.divisionid=s.id)" & vbCrLf
    s = s & "       left outer JOIN tblProjectManager p ON(j.PM=p.PM)" & vbCrLf
    s = s & "       where s.id=" & HFApp.DivisionID & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'Estimator',p.*" & vbCrLf
    s = s & "  FROM system_setup s" & vbCrLf
    s = s & "       left outer join tblJobs j on(Job_No=" & DbQuote(Str, mJob) & " and j.divisionid=s.id)" & vbCrLf
    s = s & "       LEFT JOIN tblProjectManager p ON(j.Estimator=p.PM)" & vbCrLf
    s = s & "       where s.id=" & HFApp.DivisionID & vbCrLf
    s = s & "UNION ALL" & vbCrLf
    s = s & "SELECT 'Purchaser',p.*" & vbCrLf
    s = s & "  FROM system_setup s" & vbCrLf
    s = s & "       left outer join tblJobs j on(Job_No=" & DbQuote(Str, mJob) & " and j.divisionid=s.id)" & vbCrLf
    s = s & "       LEFT JOIN tblProjectManager p ON(j.Purchaser=p.PM)" & vbCrLf
    s = s & "       where s.id=" & HFApp.DivisionID & vbCrLf
    With gContacts
        r = 0
        .Rows = 1
        Set rs = HFApp.SqlExec(s, dbHomefront, ra)
        While Not rs.EOF
            r = r + 1
            .AddItem ""
            .TextMatrix(r, .ColIndex("ContactType")) = "" & rs("ContactType")
            .Cell(flexcpPicture, r, .ColIndex("ContactType")) = SmallIcons.ListImages("vendor").Picture
            
            
            .TextMatrix(r, .ColIndex("PM")) = "" & rs("PM")
            .TextMatrix(r, .ColIndex("PMName")) = "" & rs("PMName")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("Phone")
            .TextMatrix(r, .ColIndex("Cell")) = "" & rs("Cell")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            rs.MoveNext
        Wend
    End With
    
    Call LoadJobContacts
    
    Dirty = False
    Exit Sub
eh: Call errHandler(SRCFILE & "LoadData", s)
End Sub



Private Sub txtBlock_Change()
    Dirty = True
End Sub

Private Sub txtBlock_GotFocus()
    SelectAll txtBlock
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
    txtFax.Text = FormatPhone(txtFax.Text)
End Sub

Private Sub txtGLPrefix_Change()
    Dirty = True
End Sub

Private Sub txtGLPrefix_GotFocus()
    SelectAll txtGLPrefix
End Sub


Private Sub txtJob_Change()
    Dirty = True
End Sub
Private Sub txtDescription_Change()
    Dirty = True
End Sub
Private Sub cboGLPrefix_Change()
    Dirty = True
End Sub
Private Sub txtAddress_Change()
    Dirty = True
End Sub
Private Sub txtCity_Change()
    Dirty = True
End Sub

Private Sub txtJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyF4 Then Call cmdJob_Click
End Sub

Private Sub txtJob_Validate(Cancel As Boolean)
    txtJob.Text = HFApp.FormatJob(txtJob.Text)
End Sub

Private Sub txtLot_Change()
    Dirty = True
End Sub

Private Sub txtLot_GotFocus()
    SelectAll txtLot
End Sub

Private Sub txtLotPlan_Change()
    Dirty = True
End Sub

Private Sub txtLotPlan_GotFocus()
    SelectAll txtLotPlan
End Sub

Private Sub txtNotes_Change()
    Dirty = True
End Sub

Private Sub txtNotes_GotFocus()
    SelectAll txtNotes
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    txtPhone.Text = FormatPhone(txtPhone.Text)
End Sub

Private Sub txtPostal_Change()
    Dirty = True
End Sub
Private Sub txtPhone_Change()
    Dirty = True
End Sub
Private Sub txtFax_Change()
    Dirty = True
End Sub


Private Sub txtJob_GotFocus()
    SelectAll txtJob
End Sub
Private Sub txtDescription_GotFocus()
    SelectAll txtDescription
End Sub
Private Sub txtAddress_GotFocus()
    SelectAll txtAddress
End Sub
Private Sub txtCity_GotFocus()
    SelectAll txtCity
End Sub
Private Sub cboProvince_GotFocus()
    SelectAll cboProvince
End Sub
Private Sub txtPostal_GotFocus()
    SelectAll txtPostal
End Sub
Private Sub txtPhone_GotFocus()
    SelectAll txtPhone
End Sub
Private Sub txtFax_GotFocus()
    SelectAll txtFax
End Sub

Private Sub LoadCommunities()
    Call LoadComboBox(cboCommunity, HFApp.Databases(dbHomefront), "SELECT c.Description,c.Area,0 FROM tblLocality c join divisioncommunities d on c.Area = d.Community where d.DivisionID = " & HFApp.DivisionID & " ORDER BY c.Description")
End Sub

Private Sub LoadPOs()
    Dim s As String
    Dim rs As Recordset
    
    s = ""
    s = s & "select p.PONumber,p.PODate,p.POIndex,p.description" & vbCrLf
    s = s & "      ,p.Vendor,v.vendor_name" & vbCrLf
    s = s & "      ,sum(d.pretax) Pretax" & vbCrLf
    s = s & "      ,sum(d.jctax + d.njctax) Tax" & vbCrLf
    s = s & "      ,p.status,p.approvedby,p.approveddate" & vbCrLf
    s = s & "      ,p.postingbatch" & vbCrLf
    s = s & "      ,p.cancelled" & vbCrLf
    s = s & "      ,p.cancelledBy" & vbCrLf
    s = s & "      ,p.cancelleddate" & vbCrLf
    s = s & "      ,p.cancelledNotes" & vbCrLf
    s = s & "      ,c.postingbatch CancelPostingBatch" & vbCrLf
    s = s & "  from pomaster p" & vbCrLf
    s = s & "       left outer join tblvendors v on(p.vendor=v.vendor_id and v.DivisionID = p.DivisionID)" & vbCrLf
    s = s & "       left outer join poitems d on(p.ponumber=d.ponumber and p.DivisionID = d.DivisionID)" & vbCrLf
    s = s & "       left outer join pochangeorders c on(p.ponumber=c.ponumber and p.DivisionID = d.DivisionID)" & vbCrLf
    s = s & " where p.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "   and p.job = " & DbQuote(Str, mJob) & vbCrLf
    s = s & "group by p.PODate,p.PONumber,p.POIndex,p.description,p.Vendor,v.vendor_name,p.postingbatch" & vbCrLf
    s = s & "      ,p.cancelled,p.cancelledBy,c.postingbatch,p.cancelleddate,p.cancellednotes,p.status,p.approvedby,p.approveddate" & vbCrLf
    s = s & "order by p.ponumber" & vbCrLf
    Set rs = HFApp.SqlExec(s, dbHomefront)
    With gPOs
        .Rows = 1
        While Not rs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, .ColIndex("PODate")) = "" & rs("PODate")
            .TextMatrix(.Rows - 1, .ColIndex("PONumber")) = "" & rs("PONumber")
            .TextMatrix(.Rows - 1, .ColIndex("POIndex")) = "" & rs("POIndex")
            .TextMatrix(.Rows - 1, .ColIndex("Description")) = "" & rs("Description")
            .TextMatrix(.Rows - 1, .ColIndex("Vendor")) = "" & rs("Vendor")
            .TextMatrix(.Rows - 1, .ColIndex("Vendor_name")) = "" & rs("Vendor_name")
            .TextMatrix(.Rows - 1, .ColIndex("PostingBatch")) = "" & rs("PostingBatch")
            .TextMatrix(.Rows - 1, .ColIndex("CancelPostingBatch")) = "" & rs("CancelPostingBatch")
            .TextMatrix(.Rows - 1, .ColIndex("CancelledBy")) = "" & rs("CancelledBy")
            .TextMatrix(.Rows - 1, .ColIndex("CancelledDate")) = "" & rs("CancelledDate")
            .TextMatrix(.Rows - 1, .ColIndex("CancelledNotes")) = "" & rs("CancelledNotes")
            .TextMatrix(.Rows - 1, .ColIndex("Pretax")) = Format(Val("" & rs("Pretax")), "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("Tax")) = Format(Val("" & rs("Tax")), "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("Total")) = Format(Val("" & rs("Pretax")) + Val("" & rs("Tax")), "0.00")
            

            If "" & rs("Cancelled") = "True" Then
                .RowData(.Rows - 1) = "Cancelled"
                .Cell(flexcpFontStrikethru, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbGrayText
            Else
                .RowData(.Rows - 1) = "" & rs("Status")
                If "" & rs("Status") = "Pending" Then
                    .TextMatrix(.Rows - 1, .ColIndex("Approval")) = "Approval Required"
                    '.Cell(flexcpFontBold, .Rows - 1, .ColIndex("Approval"), .Rows - 1, .ColIndex("Approval")) = True
                    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
                    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &H80&
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("Approval")) = "" & rs("approvedby") & ": " & Format("" & rs("approveddate"), "mmm d, yyyy")
                End If
            End If

            rs.MoveNext
        Wend
    End With

End Sub


Private Sub LoadProperties()
On Error GoTo eh
    Dim s As String
    Dim rs As adodb.Recordset
    Dim r As Long
    Dim Category As String
    Dim FieldName As String
    Dim fields As Recordset
    Dim Data   As Recordset
    Dim TaxGroups As String
    Dim IntacctDepartments As String
    Dim ScheduleTemplates As String
    
   
    mIsMultifamily = HFApp.Options(MultiFamily) = "true"
    
On Error Resume Next
    TaxGroups = ""
    Set rs = HFApp.SqlExec("select taxgroup,description from TaxGroups where DivisionID = " & HFApp.DivisionID & " order by 1")
    While Not rs.EOF
        TaxGroups = TaxGroups & "|" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    TaxGroups = Mid(TaxGroups, 2)
    
    IntacctDepartments = ""
    Set rs = HFApp.SqlExec("select id,name from IntacctDepartments where DivisionID = " & HFApp.DivisionID & " order by 1")
    While Not rs.EOF
        IntacctDepartments = IntacctDepartments & "|" & rs(0) & vbTab & rs(1)
        rs.MoveNext
    Wend
    IntacctDepartments = Mid(IntacctDepartments, 2)
    

    ScheduleTemplates = HFApp.Options.ValueByName("ScheduleTemplates")
    
    s = ""
    s = s & "select" & vbCrLf
    s = s & "  j.shellscheduletemplate" & vbCrLf
    s = s & " ,j.scheduletemplate" & vbCrLf
    s = s & " ,j.permitnumber" & vbCrLf
    s = s & " ,j.permitreceiveddate" & vbCrLf
    s = s & " ,j.preconStartDate" & vbCrLf
    s = s & " ,j.preconScheduleTemplate" & vbCrLf
    s = s & " ,j.start_date" & vbCrLf
    s = s & " ,j.LabourTaxGroup" & vbCrLf
    s = s & " ,j.MaterialTaxGroup" & vbCrLf
    s = s & " ,j.SubcontractTaxGroup" & vbCrLf
    s = s & " ,j.EquipmentTaxGroup" & vbCrLf
    s = s & " ,j.OverheadTaxGroup" & vbCrLf
    s = s & " ,j.OtherTaxGroup" & vbCrLf
    s = s & " ,j.ARTaxGroup" & vbCrLf
    s = s & " ,j.IntacctDepartment" & vbCrLf
    s = s & " ,j.TarionEnrollmentNumber" & vbCrLf
    s = s & " ,isnull(isnull(p.TarionBuilderNumber,c.TarionBuilderNumber),j.TarionBuilderNumber) TarionBuilderNumber" & vbCrLf
    s = s & " ,case when isnull(p.TarionBuilderNumber,'')<>'' then 1" & vbCrLf
    s = s & "       when isnull(c.TarionBuilderNumber,'')<>'' then 1" & vbCrLf
    s = s & "       else 0 end TarionBuilderReadOnly" & vbCrLf
    s = s & "from tblJobs j" & vbCrLf
    s = s & "left join tbllocality c on c.area=j.community" & vbCrLf
    s = s & "left join CommunityPhase p on p.community=j.community and p.communityphase=j.communityphase" & vbCrLf
    s = s & "where j.DivisionID = " & HFApp.DivisionID & vbCrLf
    s = s & "and j.job_no=" & DbQuote(Str, mJob)
    Set rs = HFApp.SqlExec(s)
    With gProperties
        r = -1
        .Rows = 0
        
        
        'add built in properties first
        ' -- permit, permitdate, scheduletemplate, startdate
        Category = "Scheduling"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_PermitNumber = r
        .TextMatrix(r, .ColIndex("name")) = "Permit Number"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("PermitNumber")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_PermitDate = r
        .TextMatrix(r, .ColIndex("name")) = "Permit Received Date"
        .TextMatrix(r, .ColIndex("value")) = Format("" & rs("PermitReceivedDate"), "medium date")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = DateTime
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        
If HFApp.Options.ValueByName("EnablePreconJobs") = "true" Then
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_PreconTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Precon Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("PreconScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_PreconStart = r
        .TextMatrix(r, .ColIndex("name")) = "Precon Start Date"
        .TextMatrix(r, .ColIndex("value")) = Format("" & rs("PreconStartDate"), "medium date")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = DateTime
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
End If
        
If mIsMultifamily Then
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_ShellTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Shell Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ShellScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_UnitTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Unit Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
Else
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_UnitTemplate = r
        .TextMatrix(r, .ColIndex("name")) = "Const Template"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ScheduleTemplate")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = ScheduleTemplates
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
End If
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_ConstStart = r
        .TextMatrix(r, .ColIndex("name")) = "Const Start Date"
        .TextMatrix(r, .ColIndex("value")) = Format("" & rs("Start_Date"), "medium date")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = DateTime
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

        
        
'add built in properties first
' -- Tarion
If HFApp.Options.ValueByName("EnableTarionFields") = "true" Then
        Category = "Tarion"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True

        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_TarionBuilderNumber = r
        .TextMatrix(r, .ColIndex("name")) = "Builder No"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("TarionBuilderNumber")
        .TextMatrix(r, .ColIndex("category")) = Category
        
        If "" & rs("TarionBuilderReadOnly") = 1 Then
            .TextMatrix(r, .ColIndex("DataType")) = "READONLY"
            .Cell(flexcpForeColor, r, .ColIndex("name")) = vbGrayText
            .Cell(flexcpForeColor, r, .ColIndex("value")) = vbGrayText
        Else
            .TextMatrix(r, .ColIndex("DataType")) = Str
            .Cell(flexcpForeColor, r, .ColIndex("name")) = vbWindowText
            .Cell(flexcpForeColor, r, .ColIndex("value")) = vbWindowText
        End If
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_TarionEnrollmentNumber = r
        .TextMatrix(r, .ColIndex("name")) = "Job Enrollment No"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("TarionEnrollmentNumber")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

End If

        
If HFApp.Options(AccountingSystem) = asIntacct Then
        

        Category = "Intacct"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_IntacctDepartment = r
        .TextMatrix(r, .ColIndex("name")) = "Department"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("IntacctDepartment")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .TextMatrix(r, .ColIndex("PickList")) = IntacctDepartments
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
End If
        
        ' -- tax groups
        Category = "Tax Groups"
        r = r + 1
        .AddItem ""
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("name")) = Category
        Set .Cell(flexcpPicture, r, .ColIndex("name")) = SmallIcons.ListImages("category").Picture
        .RowOutlineLevel(r) = 0
        .IsSubtotal(r) = True
        
        
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_LabTax = r
        .TextMatrix(r, .ColIndex("name")) = "Labour"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("LabourTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_MatTax = r
        .TextMatrix(r, .ColIndex("name")) = "Material"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("MaterialTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
                    
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_SubTax = r
        .TextMatrix(r, .ColIndex("name")) = "Subcontract"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("SubContractTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_EquTax = r
        .TextMatrix(r, .ColIndex("name")) = "Equipment"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("EquipmentTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_OvrTax = r
        .TextMatrix(r, .ColIndex("name")) = "Overhead"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("OverheadTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
            
        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_OthTax = r
        .TextMatrix(r, .ColIndex("name")) = "Other"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("OtherTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True

        r = r + 1
        .AddItem ""
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
        mprop_ARTax = r
        .TextMatrix(r, .ColIndex("name")) = "Accounts Receivable"
        .TextMatrix(r, .ColIndex("value")) = "" & rs("ARTaxGroup")
        .TextMatrix(r, .ColIndex("category")) = Category
        .TextMatrix(r, .ColIndex("PickList")) = TaxGroups
        .TextMatrix(r, .ColIndex("DataType")) = Str
        .RowOutlineLevel(r) = 1
        .IsSubtotal(r) = True
On Error GoTo eh
            
            
        'now load user defined fields
        mprop_UserFlds = r + 1
        Set fields = HFApp.SqlExec("SELECT * FROM JobCustomFieldDefs ORDER BY Category,Name")
        Set Data = HFApp.SqlExec("SELECT * FROM JobCustomFields WHERE divisionid=" & HFApp.DivisionID & " and Job_No=" & DbQuote(Str, mJob))
        
        While Not fields.EOF
            
            If Category <> "" & fields("category") Then
                Category = "" & fields("category")
                r = r + 1
                .AddItem ""
                .TextMatrix(r, .ColIndex("category")) = Category
                .TextMatrix(r, .ColIndex("name")) = IIf(Category = "", "unclassified", Category)
                Set .Cell(flexcpPicture, r, .ColIndex("name")) = SmallIcons.ListImages("category").Picture
                .RowOutlineLevel(r) = 0
                .IsSubtotal(r) = True
            End If
            
            r = r + 1
            .AddItem ""
            FieldName = "" & fields("Name")
            .TextMatrix(r, .ColIndex("category")) = Category
            .TextMatrix(r, .ColIndex("name")) = FieldName
            .TextMatrix(r, .ColIndex("PickList")) = "" & fields("PickList")
            Select Case Data(FieldName).Type
                Case adInteger, adTinyInt, adSmallInt, adBigInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adUnsignedBigInt: .TextMatrix(r, .ColIndex("DataType")) = NumInt
                Case adDouble, adSingle, adDecimal, adNumeric:                                                                           .TextMatrix(r, .ColIndex("DataType")) = Num
                Case adCurrency:                                                                                                         .TextMatrix(r, .ColIndex("DataType")) = Cur
                Case adBoolean:                                                                                                          .TextMatrix(r, .ColIndex("DataType")) = Bit
                Case adDate, adDBDate, adDBTime, adDBTimeStamp:                                                                          .TextMatrix(r, .ColIndex("DataType")) = DateTime
                Case adVarChar, adBSTR, adChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adVariant:                           .TextMatrix(r, .ColIndex("DataType")) = Str
            End Select
            .TextMatrix(r, .ColIndex("length")) = Data(FieldName).DefinedSize
            On Error Resume Next
            .TextMatrix(r, .ColIndex("value")) = Data(FieldName).Value
            On Error GoTo eh
            
            .RowOutlineLevel(r) = 1
            .IsSubtotal(r) = True
            
            fields.MoveNext
        Wend
        
        
        
        
        Call .AutoSize(0, .Cols - 1)
        Call .Outline(0)
        
        

        
    End With
    
Exit Sub
eh: Call errHandler(SRCFILE & "LoadProperties")
End Sub



Private Sub gPOs_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gPOs.MouseRow = 0 Then
            Cancel = True
'            Call ShowColumnMenu(gPOs)
        End If
    End If
End Sub

Public Sub EditJob(Job As String)
On Error Resume Next
    mJob = Job
    Me.Show vbModal
End Sub

Private Sub cboCommunity_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete, vbKeyBack
            cboCommunity.ListIndex = -1
    End Select
End Sub

Private Function CleanJob(FormatedJob As String) As String
    If HFApp.Options(AccountingSystem) = asTimberline Then
        CleanJob = Trim(Replace(Replace(Replace(Replace(Replace(FormatedJob, "\", ""), ",", ""), "/", ""), ".", ""), "-", ""))
    Else
        CleanJob = FormatedJob
    End If
End Function

Private Function LoadGLPrefixes()
    Dim s As String
    Dim rs As Recordset
    
    Select Case True
        
        Case HFApp.Options(AccountingSystem) = asIntacct
            lblGLPrefix.Caption = "Location"
            cboGLPrefix.Visible = True
            txtGLPrefix.Visible = False
            s = "select description,gl_prefix,'' from gl_prefix where divisionid=" & DbQuote(Num, HFApp.DivisionID)
            Call LoadComboBox(cboGLPrefix, HFApp.Databases(dbHomefront), s)
        
        Case HFApp.Options(AccountingSystem) = asTimberline And HFApp.Databases(dbAccounting).State = adStateOpen
            
            cboGLPrefix.Visible = True
            txtGLPrefix.Visible = False
            
            'which prefix table?
            s = "select account_prefix_a_length,account_prefix_ab_length,account_prefix_abc_length from glm_master__account_format"
            Set rs = HFApp.SqlExec(s, dbAccounting)
            Select Case True
                Case Val("" & rs(2)) <> 0: s = "select account_prefix_abc_description ,account_prefix_abc, 0   from glm_master__account_prefix_abc_1"
                Case Val("" & rs(1)) <> 0: s = "select account_prefix_ab_description ,account_prefix_ab, 0   from glm_master__account_prefix_ab_1"
                Case Val("" & rs(0)) <> 0: s = "select account_prefix_a_description ,account_prefix_a, 0   from glm_master__account_prefix_a_1"
            End Select
            Call LoadComboBox(cboGLPrefix, HFApp.Databases(dbAccounting), s)
        
        Case Else
            
            cboGLPrefix.Visible = False
            txtGLPrefix.Visible = False
            lblGLPrefix.Visible = False
    
    End Select
    
End Function

Private Function ValidateJobNumber() As Boolean
    Dim b As Boolean
    Select Case HFApp.Options(AccountingSystem)
    
        Case asTimberline
            If IsBetween(Len(CleanJob(txtJob.Text)), Val(HFApp.Options(Job_Section2)) + Val(HFApp.Options(Job_Section3)), Val(HFApp.Options(Job_Section1)) + Val(HFApp.Options(Job_Section2)) + Val(HFApp.Options(Job_Section3))) Then
                ValidateJobNumber = True
            Else
                MsgBox "Incorrect job format." & vbCrLf & "Correct format is " & Replace(HFApp.Options(Job_Mask), "&", "x") & vbCrLf & "Reenter job.", vbExclamation, App.ProductName
            End If
            
        Case asMasterBuilder
            If IsNumeric(txtJob.Text) And Val(txtJob.Text) - Int(Val(txtJob.Text)) = 0 Then
                ValidateJobNumber = True
                txtJob.Text = Int(Val(txtJob.Text))
            Else
                MsgBox "Job number must be an integer greater than zero." & vbCrLf & "Reenter job.", vbExclamation, App.ProductName
            End If
            
        Case Else
            ValidateJobNumber = True
            
    End Select
End Function
Private Sub gProperties_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh

    With gProperties
        .ComboList = ""

        If Col <> .ColIndex("Value") Then
            Cancel = True
            Exit Sub
        End If

        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt
            Case Num
            Case Cur
            Case Bit:      .ComboList = "Yes|No"
            Case Date:     .ComboList = "|..."
            Case DateTime: .ComboList = "|..."
            Case Str:      .EditMaxLength = .ValueMatrix(Row, .ColIndex("Length"))
                           .ComboList = .TextMatrix(Row, .ColIndex("PickList"))
            Case Else:     Cancel = True
        End Select
        
    End With

Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_BeforeEdit")
End Sub

Private Sub gProperties_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo eh
    With gProperties
        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt:
            Case Num:
            Case Cur:
            Case Bit:
            Case DateTime, Date
                Call DCalendar.Popup(gProperties, .RowPos(.Row) + .RowHeight(.Row), .colPos(.Col))
                Dirty = True
            Case Str:
            Case Else:
        End Select
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_CellButtonClick")
End Sub

Private Sub gProperties_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    With gProperties
        If .Row < 0 Then Exit Sub
        Select Case KeyCode

            Case vbKeySpace
                .IsCollapsed(.Row) = IIf(.IsCollapsed(.Row) = flexOutlineCollapsed, flexOutlineExpanded, flexOutlineCollapsed)
                
            Case vbKeyLeft
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 And .IsCollapsed(.Row) <> flexOutlineCollapsed Then
                    .IsCollapsed(.Row) = flexOutlineCollapsed
                Else
                    If .GetNodeRow(.Row, flexNTParent) <> -1 Then .Row = .GetNodeRow(.Row, flexNTParent)
                End If
                
            Case vbKeyDelete
                If .Col = .ColIndex("value") Then
                    .Text = ""
                    Call gProperties_ValidateEdit(.Row, .Col, False)
                    Dirty = True
                    
                    Select Case .ValueMatrix(.Row, .ColIndex("DataType"))
                        Case Cur, Num, NumInt:  .Text = "0"
                    End Select
                End If
            
            
            
            Case vbKeyRight
                If .GetNodeRow(.Row, flexNTFirstChild) <> -1 Then
                    If .IsCollapsed(.Row) = flexOutlineCollapsed Then
                        .IsCollapsed(.Row) = flexOutlineExpanded
                    Else
                        .Row = .GetNodeRow(.Row, flexNTFirstChild)
                    End If
                End If
                
        End Select
    End With
End Sub

Private Sub gProperties_RowColChange()
On Error GoTo eh

    With gProperties
        If .Row >= 0 Then
            If .TextMatrix(.Row, .ColIndex("DataType")) = "" Then
                .Col = .ColIndex("Name")
            Else
                .Col = .ColIndex("Value")
            End If
        End If
    End With

Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_ValidateEdit")
End Sub

Private Sub gProperties_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo eh
    
    With gProperties
        Select Case .ValueMatrix(Row, .ColIndex("DataType"))
            Case NumInt:    .EditText = Int(Val(.EditText))
            Case Num:       .EditText = Val(.EditText)
            Case Cur:       .EditText = Format(Round(Val("" & Replace(Replace(Replace(.EditText, "%", ""), ",", ""), "$", "")), 2), "Currency")
            Case Bit:
            Case DateTime:
                If IsDate(.EditText) Or .EditText = "" Then
                    .EditText = Format(.EditText, "medium date")
                Else
                    Cancel = True
                End If
            Case Str:
            Case Else:       Cancel = True
        End Select
    End With
    If Not Cancel Then Dirty = True
    
Exit Sub
eh: Call errHandler(SRCFILE & "gProperties_ValidateEdit")
End Sub

Private Sub gContacts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gContacts
        Select Case .ColKey(Col)
            Case "PM":     .ComboList = "..."
            Case "PMName": .ComboList = "..."
            Case Else:     Cancel = True
        End Select
    End With
End Sub

Private Sub gContacts_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    If Button = vbRightButton Then
        If gContacts.MouseRow = 0 Then
            Cancel = True
'            Call FMain.ShowColumnMenu(gContacts)
        End If
    End If
End Sub

Private Sub gContacts_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    With gContacts
    
        If .TextMatrix(Row, .ColIndex("ContactType")) = "PM" Then
            s = "SELECT PM,PMName Name,Phone,fax,email,cell FROM tblProjectManager WHERE Inactive<>1 AND 1=PrjMgr"
        Else
            s = "SELECT PM,PMName Name,Phone,fax,email,cell FROM tblProjectManager WHERE Inactive<>1 AND 1=" & .TextMatrix(Row, .ColIndex("ContactType"))
        End If
        
        If FPickList.Choose(HFApp.Databases(dbHomefront), Choose(Row, "Project Manager", "Estimator", "Purchaser"), s, .TextMatrix(Row, .ColIndex("PM")), , , , "PM,Phone,fax,email,cell") Then
            .TextMatrix(Row, .ColIndex("PM")) = FPickList.SelectedItem("PM")
            .TextMatrix(Row, .ColIndex("PMName")) = FPickList.SelectedItem("Name")
            .TextMatrix(Row, .ColIndex("Phone")) = FPickList.SelectedItem("Phone")
            .TextMatrix(Row, .ColIndex("Cell")) = FPickList.SelectedItem("cell")
            .TextMatrix(Row, .ColIndex("Fax")) = FPickList.SelectedItem("Fax")
            .TextMatrix(Row, .ColIndex("Email")) = FPickList.SelectedItem("Email")
            .RowData(Row) = "DIRTY"
            Dirty = True
        End If
    End With
End Sub



Private Sub gJobContacts_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim s As String
    Dim rs As Recordset
    Dim r As Long
    
    With gJobContacts
    Select Case .ColKey(Col)
        Case "Name"
            s = ""
            s = s & "Customer Contacts" & Chr(1) & vbCrLf
            s = s & "   select Role,firstname Name,ContactID" & vbCrLf
            s = s & "   from contacts " & vbCrLf
            s = s & "   where contacttypeid=99 and isnull(arcustomer,'')<>'' and isnull(arcustomer,'')=" & DbQuote(Str, txtARCustomer.Text) & Chr(0)
            s = s & "Vendor Contacts" & Chr(1) & vbCrLf
            s = s & "   select v.vendor_name Company,c.Role,c.firstname Name,ContactID" & vbCrLf
            s = s & "   from contacts c" & vbCrLf
            s = s & "   join tblvendors v on(c.vendorcode=v.vendor_id and v.DivisionID = " & HFApp.DivisionID & ")" & vbCrLf
            s = s & "   where contacttypeid=99 and vendorcode<>'' and c.divisionid=" & HFApp.DivisionID & vbCrLf
            If FPickList.Choose(HFApp.Databases(dbHomefront), "Contact", s, , , , , "ContactID") Then
                Set rs = HFApp.SqlExec("select * from contacts where contactid=" & DbQuote(Num, FPickList.SelectedItem("ContactID")))
                r = Row
                .TextMatrix(r, .ColIndex("ContactID")) = "" & rs("ContactID")
                .TextMatrix(r, .ColIndex("Company")) = "" & rs("CompanyName")
                .TextMatrix(r, .ColIndex("Name")) = "" & rs("FirstName")
                .TextMatrix(r, .ColIndex("Phone")) = "" & rs("WorkPhone")
                .TextMatrix(r, .ColIndex("Cell")) = "" & rs("CellPhone")
                .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
                .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
                .TextMatrix(r, .ColIndex("SendVia")) = Val("" & rs("CommModeID"))
                .TextMatrix(r, .ColIndex("SmsAddress")) = "" & rs("SmsAddress")
                Dirty = True
                .RowData(r) = "DIRTY"
            End If
            
    End Select
    End With
End Sub

Private Sub gJobContacts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim r As Long
    If KeyCode = vbKeyDelete And Shift <> 0 And gJobContacts.Rows > 1 And gJobContacts.Row <> gJobContacts.Rows - 1 Then
    With gJobContacts
        r = .Row
        .TextMatrix(r, .ColIndex("ContactID")) = ""
        .TextMatrix(r, .ColIndex("Name")) = ""
        .TextMatrix(r, .ColIndex("Phone")) = ""
        .TextMatrix(r, .ColIndex("Cell")) = ""
        .TextMatrix(r, .ColIndex("Fax")) = ""
        .TextMatrix(r, .ColIndex("Email")) = ""
        .TextMatrix(r, .ColIndex("SendVia")) = ""
        .TextMatrix(r, .ColIndex("SmsAddress")) = ""
        .RowData(r) = Dirty
        Dirty = True
    End With
    End If
End Sub

Private Sub gJobContacts_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim r As Long
    
    With gJobContacts
        Select Case .ColKey(Col)
            Case "Cell":    .EditText = FormatPhone(.EditText)
            Case "Phone":   .EditText = FormatPhone(.EditText)
            Case "fax":     .EditText = FormatPhone(.EditText)
        End Select
        .RowData(Row) = "DIRTY"
        Dirty = True
    End With
End Sub

Private Sub gJobContacts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With gJobContacts
        .EditMaxLength = 0
        .ComboList = ""
        If .TextMatrix(Row, .ColIndex("ContactID")) = "" Then
            Cancel = True
        End If
        Select Case .ColKey(Col)
            Case "Role":         Cancel = True
            Case "Company":      Cancel = True
            Case "Phone":        .EditMaxLength = 35
            Case "Cell":         .EditMaxLength = 35
            Case "Fax":          .EditMaxLength = 35
            Case "Email":        .EditMaxLength = 500
            Case "Name":         .ComboList = "...": Cancel = False
            Case "Role":         .EditMaxLength = 25
            Case "SmsAddress":   .EditMaxLength = 50
        End Select
    End With
End Sub

Private Sub LoadJobContacts()
On Error GoTo eh
    Dim roles As String
    Dim s As String
    Dim r As Long
    Dim rs As Recordset
    
    
    roles = HFApp.Options.ValueByName("JobContactRoles")
    With gJobContacts
        
        'add roles
        .Rows = 1
        For r = 1 To Parse(roles, , "|")
            s = Parse(roles, r, "|")
            If s <> "" Then
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("Role")) = s
            End If
        Next
    
        'now look up contacts assigned to those roles
        s = ""
        s = s & "select j.Role JobRole,c.*" & vbCrLf
        s = s & "  from jobContacts j" & vbCrLf
        s = s & "  join contacts c on(j.contactid=c.contactid)" & vbCrLf
        s = s & " where j.divisionid=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
        s = s & "   and j.job=" & DbQuote(Str, mJob) & vbCrLf
        Set rs = HFApp.SqlExec(s)
        While Not rs.EOF
            r = FindRow("" & rs("JobRole"))
            If r = -1 Then
                .AddItem ""
                r = .Rows - 1
            End If
            .TextMatrix(r, .ColIndex("ContactID")) = "" & rs("ContactID")
            .TextMatrix(r, .ColIndex("Role")) = "" & rs("JobRole")
            .TextMatrix(r, .ColIndex("Company")) = "" & rs("CompanyName")
            .TextMatrix(r, .ColIndex("Name")) = "" & rs("FirstName")
            .TextMatrix(r, .ColIndex("Phone")) = "" & rs("WorkPhone")
            .TextMatrix(r, .ColIndex("Cell")) = "" & rs("CellPhone")
            .TextMatrix(r, .ColIndex("Fax")) = "" & rs("Fax")
            .TextMatrix(r, .ColIndex("Email")) = "" & rs("Email")
            .TextMatrix(r, .ColIndex("SendVia")) = Val("" & rs("CommModeID"))
            .TextMatrix(r, .ColIndex("SmsAddress")) = "" & rs("SmsAddress")
            rs.MoveNext
        Wend
    End With

Exit Sub
eh: errHandler ("LoadJobContacts")
End Sub

Private Function FindRow(Role As String) As Long
    Dim r As Long
    With gJobContacts
    FindRow = -1
    For r = 1 To .Rows - 1
        If .TextMatrix(r, .ColIndex("Role")) = Role Then
            FindRow = r
            Exit Function
        End If
    Next
    End With
End Function

Private Sub SaveJobContacts()
On Error GoTo eh
    Dim s As String
    Dim r As Long
    
    With gJobContacts
        For r = 1 To .Rows - 1
            If .RowData(r) = "DIRTY" Then
                
                
                s = ""
                s = s & "INSERT INTO JobContacts(DivisionID,ContactID,Role,Job)" & vbCrLf
                s = s & "VALUES(" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "      ," & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                s = s & "      ," & DbQuote(Str, mJob) & ")"
                On Error Resume Next
                Call HFApp.SqlExec(s)
                On Error GoTo eh
                
                s = ""
                s = s & "UPDATE JobContacts" & vbCrLf
                s = s & "SET ContactID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                s = s & "WHERE DivisionID=" & DbQuote(Num, HFApp.DivisionID) & vbCrLf
                s = s & "  AND Role=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                s = s & "  AND Job=" & DbQuote(Str, mJob) & vbCrLf
                Call HFApp.SqlExec(s)
            
            
                If .TextMatrix(r, .ColIndex("ContactID")) <> "" Then
                    s = ""
                    s = s & "UPDATE Contacts" & vbCrLf
                    s = s & "SET Role=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Role"))) & vbCrLf
                    s = s & "   ,FirstName=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Name"))) & vbCrLf
                    s = s & "   ,WorkPhone=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Phone"))) & vbCrLf
                    s = s & "   ,CellPhone=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Cell"))) & vbCrLf
                    s = s & "   ,Fax=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Fax"))) & vbCrLf
                    s = s & "   ,Email=" & DbQuote(Str, .TextMatrix(r, .ColIndex("Email"))) & vbCrLf
                    s = s & "   ,CommModeID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("SendVia"))) & vbCrLf
                    s = s & "   ,SmsAddress=" & DbQuote(Str, .TextMatrix(r, .ColIndex("SmsAddress"))) & vbCrLf
                    s = s & "WHERE ContactID=" & DbQuote(Num, .TextMatrix(r, .ColIndex("ContactID"))) & vbCrLf
                    Call HFApp.SqlExec(s)
                End If
                .RowData(r) = ""
            End If
        Next
    End With
Exit Sub
eh: Call errHandler(SRCFILE & "SaveJobContacts", s)
End Sub



Public Function GetCustomDesc(ItemName As String, Optional Plural As Boolean = False) As String
On Error Resume Next
    Dim s As String
    s = ItemName
    s = HFApp.SqlExec("select isnull(nullif(custom_description,''),item) from customdescriptions where item=" & DbQuote(Str, ItemName), dbHomefront)(0)
    
    If Plural Then
        Select Case Right(s, 1)
            Case "s"
            Case "y":   s = left(s, Len(s) - 1) & "ies"
            Case Else:  s = s & "s"
        End Select
    End If
    
    GetCustomDesc = s
End Function

