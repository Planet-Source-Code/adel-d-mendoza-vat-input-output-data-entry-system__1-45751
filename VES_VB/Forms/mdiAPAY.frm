VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiAPAY 
   BackColor       =   &H80000001&
   Caption         =   "VAT Input/Output Data Entry System (version 3.00)"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10605
   Icon            =   "mdiAPAY.frx":0000
   LinkTopic       =   "mdiAPAY"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5955
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "12/05/2003"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:01 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   4  'Align Right
      Height          =   5955
      Left            =   9615
      TabIndex        =   6
      Top             =   0
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   10504
      ButtonWidth     =   820
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   120
         Top             =   7200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":0E42
               Key             =   "cust"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":1C94
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":20E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":2538
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":298A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":2DDC
               Key             =   "sec"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":3C2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":4080
               Key             =   "report"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":4D5A
               Key             =   "party"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":5BAC
               Key             =   "profit"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":6E2E
               Key             =   "reminder"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":7708
               Key             =   "img7"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiAPAY.frx":7FE2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin LVbuttons.LaVolpeButton btnReports 
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   5040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "mdiAPAY.frx":8434
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "4"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnDebitAcct 
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   3120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "mdiAPAY.frx":8450
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "5"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnBranch 
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   2160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "mdiAPAY.frx":846C
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "3"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnVendor 
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "mdiAPAY.frx":8488
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "9"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnTaxCodes 
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   4080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "mdiAPAY.frx":84A4
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "2"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnCustomer 
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         BCOL            =   14215660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "mdiAPAY.frx":84C0
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Reports"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   17
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "TAX Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   16
         Top             =   4680
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Debit Acct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   14
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Branch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   13
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Vendor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Logged In Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   8
         Top             =   6600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "12:00 AM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   7
         Top             =   7080
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":84DC
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":91B6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":A008
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":AE5A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":B734
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":C00E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":C8E8
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":D2B2
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":DB8C
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":DEA6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":E780
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":F05A
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":F934
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":FC4E
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":10528
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":10E02
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":116DC
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":11FB6
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":12890
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":1316A
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":13A44
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":1431E
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":14BF8
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":154D2
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":15DAC
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":16686
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":16F60
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":1783A
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":18114
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":189EE
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":192A4
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":19B7E
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":19FD0
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":1A422
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiAPAY.frx":1CBD4
            Key             =   "IMG35"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "&Data Entry"
      Begin VB.Menu mnuSbar1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Data Entry|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuVatInput 
         Caption         =   "{IMG:4}VAT Input Transactions"
      End
      Begin VB.Menu mnuVatOutput 
         Caption         =   "{IMG:4}VAT Output Transactions"
      End
      Begin VB.Menu mnuVATexempt 
         Caption         =   "{IMG:4}VAT Output (Exempt)"
      End
      Begin VB.Menu mnuTax 
         Caption         =   "{IMG:4}Witholding Tax Transactions"
      End
   End
   Begin VB.Menu mnuFileUpdate 
      Caption         =   "&File Update"
      Begin VB.Menu mnuSbar2 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:File Update|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuBranchCode 
         Caption         =   "{IMG:16}Branch Code"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUploadCustomer 
         Caption         =   "{IMG:16}Upload Customer"
      End
      Begin VB.Menu mnuUploadVendor 
         Caption         =   "{IMG:16}Upload Vendor"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "{IMG:16}Change Password"
      End
   End
   Begin VB.Menu mnuFileMaintenance 
      Caption         =   "File &Maintenance"
      Begin VB.Menu mnuSbar3 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:File Maintenance|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuVendorMaintenance 
         Caption         =   "{IMG:4}Vendor Maintenance"
      End
      Begin VB.Menu mnuCustomerMaintenance 
         Caption         =   "{IMG:4}Customer Maintenance"
      End
      Begin VB.Menu mnuBranchMaintenance 
         Caption         =   "{IMG:4}Branch Maintenance"
      End
      Begin VB.Menu mnuDebitAcctMaintenance 
         Caption         =   "{IMG:4}Debit Accounts Maintenance"
      End
      Begin VB.Menu mnuTaxCodeMaintenance 
         Caption         =   "{IMG:4}TAX Codes Maintenance"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsersMaintenance 
         Caption         =   "{IMG:4}User's Maintenance"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuSBar4 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Utilities|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuVATinputDisk 
         Caption         =   "{IMG:6}Create VAT Input Diskette"
      End
      Begin VB.Menu mnuVAToutputDisk 
         Caption         =   "{IMG:6}Create VAT Output Diskette"
      End
      Begin VB.Menu mnuTaxDisk 
         Caption         =   "{IMG:6}Create TAX Diskette"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackUpFile 
         Caption         =   "{IMG:9}Backup File"
      End
      Begin VB.Menu mnuRestoreFile 
         Caption         =   "{IMG:13}Restore File"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuSbar5 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Tools|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuViewLogDetails 
         Caption         =   "{IMG:34}View Log Details"
      End
      Begin VB.Menu mnuClearLogDetails 
         Caption         =   "{IMG:11}Clear Log Details"
      End
      Begin VB.Menu sep_microsoft 
         Caption         =   "-Microsoft Tools"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "{IMG:10}Calculator"
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "{IMG:8}Notepad"
      End
      Begin VB.Menu mnuWindowsExplorer 
         Caption         =   "{IMG:6}Windows Explorer"
      End
      Begin VB.Menu mnuOnScreenKeyboard 
         Caption         =   "{IMG:24}On-Screen Keyboard"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuSbar6 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Reports|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuVatInputEditlist 
         Caption         =   "{IMG:1}VAT Input Editlist"
      End
      Begin VB.Menu mnuVatOutputEditlist 
         Caption         =   "{IMG:33}VAT Output Editlist"
      End
      Begin VB.Menu mnuVatOutputExemptEditlist 
         Caption         =   "{IMG:33}VAT Output (Exempt) Editlist"
      End
      Begin VB.Menu mnuWitholdingTaxEditlist 
         Caption         =   "{IMG:1}Witholding TAX Editlist"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&System"
      Begin VB.Menu mnuSbar7 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:System|Font:Arial|BOLD|Fsize:10|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "{IMG:12}Log Off"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "{IMG:14}Exit"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "{IMG:32}About"
      End
   End
End
Attribute VB_Name = "mdiAPAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#          Coded by Adel D. Mendoza          #
'#        Designed by Ronald S. Abian         #
'#  VES - VAT Input/Output Data Entry System  #
'#           for SWIFT FOODS, INC.            #
'#                                            #
'#           area :  mdiAPAY                  #
'#    description :  Code File MAIN MENU      #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Public cur_user As String

Private Sub MDIForm_Activate()
  If cur_user = "OPERATOR" Then
  End If
End Sub

Private Sub MDIForm_Load()
  SetMenus hwnd, SmallImages
  mdiAPAY.Toolbar2.Width = 990
  Set rsSETF = db.OpenRecordset("SELECT * FROM SETFLE")
  pubBRANCH = rsSETF.Fields!Branch
  pubBRNAME = rsSETF.Fields!Brname
  Me.StatusBar1.Panels(1).Text = "  Branch:  " + pubBRANCH + " - " + pubBRNAME
  frmSplash.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  ReleaseMenus hwnd
  Dim rsLog As DAO.Recordset

  On Error Resume Next
  Set rsLog = db.OpenRecordset("SELECT * FROM LOGIN WHERE USER = '" & pubUserName & "' AND LOGOUT = 'CURRENT'")
  rsLog.Edit
  rsLog.Fields!Logout = Now
  rsLog.Update
  rsLog.Close
  On Error Resume Next
  db.Close
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  Unload Me
  End
End Sub

Private Sub mnuLogOff_Click()
  Unload Me
  openDB
  Dim rsUSR As DAO.Recordset
  Set rsUSR = db.OpenRecordset("SELECT * FROM USERS")
  frmUserPass.Combo1.Clear
  Do While Not rsUSR.EOF
    frmUserPass.Combo1.addItem rsUSR.Fields!UserName
    rsUSR.MoveNext
  Loop
  rsUSR.Close
  frmUserPass.Show
End Sub

'******************************
'******************************
'******************************
Private Sub mnuVatInput_Click()
  mnuPointer = 1
  frmAddInfo.Show
End Sub

Private Sub mnuVatOutput_Click()
  mnuPointer = 2
  frmAddInfo.Show
End Sub

Private Sub mnuVATexempt_Click()
  mnuPointer = 3
  frmAddInfo.Show
End Sub

Private Sub mnuTax_Click()
  mnuPointer = 4
  frmAddInfo.Show
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuBackUpFile_Click()
  frmBackupDba.Show
End Sub

Private Sub mnuBranchCode_Click()
  frmBranch.Show
End Sub

Private Sub mnuBranchMaintenance_Click()
  frmVESBRN.Show
End Sub

Private Sub mnuCalculator_Click()
  On Error Resume Next
  Shell ("calc"), vbMinimizedFocus
  Exit Sub
End Sub

Private Sub mnuChangePassword_Click()
  frmChangePassword.Show
End Sub

Private Sub mnuDebitAcctMaintenance_Click()
   frmVESDBS.Show
End Sub

Private Sub mnuNotepad_Click()
  On Error Resume Next
  Shell ("notepad"), vbMaximizedFocus
  Exit Sub
End Sub

Private Sub mnuOnScreenKeyboard_Click()
  On Error Resume Next
  Shell "osk", vbNormalFocus
  Exit Sub
End Sub

Private Sub mnuRestoreFile_Click()
  frmRestoreDba.Show
End Sub

Private Sub mnuTaxCodeMaintenance_Click()
  frmVESCOD.Show
End Sub

Private Sub mnuUploadCustomer_Click()
  frmUpdateCustomer.Show
End Sub

Private Sub mnuUploadVendor_Click()
  frmUpdateVendor.Show
End Sub
Private Sub mnuVATinputDisk_Click()
  cCode = "I"
  frmCreateText.Show
End Sub

Private Sub mnuVAToutputDisk_Click()
  cCode = "O"
  frmCreateText.Show
End Sub

Private Sub mnuTaxDisk_Click()
  cCode = "T"
  frmCreateText.Show
End Sub
Private Sub mnuVatInputEditlist_Click()
  rptPointer = 1
  frmReport.Show
End Sub

Private Sub mnuVatOutputEditlist_Click()
   rptPointer = 2
   frmReport.Show
End Sub

Private Sub mnuVatOutputExemptEditlist_Click()
   rptPointer = 3
   frmReport.Show
End Sub

Private Sub mnuViewLogDetails_Click()
  frmLogDetails.Show
End Sub

Private Sub mnuWitholdingTaxEditlist_Click()
  rptPointer = 4
  frmReport.Show
End Sub

Private Sub mnuWindowsExplorer_Click()
  On Error Resume Next
  Shell ("explorer"), vbMaximizedFocus
  Exit Sub
End Sub

Private Sub mnuVendorMaintenance_Click()
  frmVESVEN.Show
End Sub

Private Sub mnuCustomerMaintenance_Click()
  frmVESCUS.Show
End Sub

Private Sub mnuUsersMaintenance_Click()
  frmUsers.Show
End Sub

Private Sub btnBranch_Click()
  frmVESBRN.Show
End Sub

Private Sub btnCustomer_Click()
  frmVESCUS.Show
End Sub

Private Sub btnDebitAcct_Click()
  frmVESDBS.Show
End Sub

Private Sub btnTaxCodes_Click()
  frmVESCOD.Show
End Sub

Private Sub btnVendor_Click()
  frmVESVEN.Show
End Sub

Private Sub btnReports_Click()
  PopupMenu mnuReports
End Sub

Private Sub mnuClearLogDetails_Click()
  Dim Sagot
  Dim LD As DAO.Recordset
  Sagot = MsgBox("Are you sure you want to Clear Log Details?", vbYesNo Or vbQuestion, "Want to Clear Log Details ...")
  If Sagot = vbYes Then
     Set LD = db.OpenRecordset("SELECT * FROM LOGIN WHERE LOGOUT <> 'CURRENT'")
     If LD.RecordCount <> 0 Then
        Do While Not LD.EOF
           LD.Delete
           LD.MoveNext
        Loop
        MsgBox "Log Details Cleared ...", vbInformation, "Log File Cleared ..."
     End If
  End If
End Sub

