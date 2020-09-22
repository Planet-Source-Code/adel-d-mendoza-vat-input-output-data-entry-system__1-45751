VERSION 5.00
Begin VB.Form frmRestoreDba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Restore Database"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   ControlBox      =   0   'False
   Icon            =   "frmRestoreDba.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6210
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "..."
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtSource 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5535
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "&Restore Database"
      Enabled         =   0   'False
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
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
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
      Left            =   177
      TabIndex        =   7
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label lblSelectedDba 
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
      Left            =   139
      TabIndex        =   6
      Top             =   1560
      Width           =   5895
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Current Database Size"
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
      TabIndex        =   4
      Top             =   240
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Restore from"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmRestoreDba"
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
'#           area :  frmRestoreDba            #
'#    description :  Code File                #
'#                   Restore Database File    #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim dbasize As Long
Dim dbasize2 As Long
Dim PathName As String

Private Sub cmdClose_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdRestore_Click()
  If MsgBox("Restoring database from location " & txtSource & " will replace existing database files.Do you want to Contunue", vbYesNo) = vbYes Then
     DoRestore txtSource.Text, App.Path
     If NoDba = True Then
        MsgBox "Database Restored Click Ok to Exit Program"
        frmRestoreDba.Hide
        Unload frmRestoreDba
     End If
  Else
     lblStatus.Caption = "Database Restore Canceled"
  End If
End Sub

Private Sub cmdSource_Click()
On Error GoTo Erro
Dim strTemp As String
strTemp = fBrowseForFolder(Me.hWnd, "Restore From")
If strTemp <> "" Then
    txtSource = strTemp
    dbasize2 = FileLen(txtSource & "\VES.mdb")
    lblSelectedDba = "Selected Backup Database is : " & Format((dbasize2 / 1024) / 1024, "standard") & "MB."
    cmdRestore.Enabled = True
End If
Erro:
    Select Case err.Number
       Case 53 'File Not Found
          lblSelectedDba = "No Backup at this location"
          cmdRestore.Enabled = False
    End Select
End Sub

Private Sub Form_Activate()
lblSize = Format((dbasize / 1024) / 1024, "standard") & "MB."
End Sub

Private Sub Form_Load()
   Me.Top = 2000
   Me.Left = 3000
   PathName = App.Path & "\VES.mdb"
   On Error GoTo err
   dbasize = FileLen(PathName)
   Call Disable_System_Menu
err:
   Exit Sub
End Sub

