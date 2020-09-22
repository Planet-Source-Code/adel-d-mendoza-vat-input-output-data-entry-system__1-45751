VERSION 5.00
Begin VB.Form frmBackupDba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Backup Database"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   Icon            =   "frmBackupDba.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup Database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdDestination 
      Caption         =   "..."
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtDestination 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Backup Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblDbaSize 
      Alignment       =   2  'Center
      Caption         =   "Current Database Size is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmBackupDba"
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
'#           area :  frmBackupDba             #
'#    description :  Code File                #
'#                :  Backup Database File     #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim dbasize As Long
Dim PathName As String

Private Sub cmdBackup_Click()
   If txtDestination <> "" Then
      DoBackup PathName, txtDestination
   ElseIf txtDestination = "" Then
      MsgBox "You must specify a destination for the backup", vbCritical
   End If
End Sub

Private Sub cmdClose_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdDestination_Click()
   Dim strTemp As String
   strTemp = fBrowseForFolder(Me.hWnd, "Select backup path")
   If strTemp <> "" Then
      txtDestination = strTemp
   End If
End Sub

Private Sub Form_Activate()
   lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB."
End Sub

Private Sub Form_Load()
   Me.Top = 2000
   Me.Left = 3000
   Me.WindowState = 0
   PathName = App.Path & "\VES.MDB"
   dbasize = FileLen(PathName)
   
   Call Disable_System_Menu
End Sub


