VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About "
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5160
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Agri Division"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "© 2003 - All Rights Reserved"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Licensed to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "VAT Input / Output Data Entry System"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "SAP / R3 Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "SFI-AGRI-IT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "SWIFT FOODS, INC."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1560
         Left            =   240
         Picture         =   "frmAbout.frx":0000
         Top             =   240
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmAbout"
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
'#           area :  frmAbout                 #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  Me.Top = 2100
  Me.Left = 3200
End Sub

