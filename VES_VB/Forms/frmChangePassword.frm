VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " « Change Password"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   4695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "C&hange"
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtRetypePassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtNewPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtOldPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Retype Password"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "New Password"
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
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Old Password"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmChangePassword"
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
'#         area :  frmChangePassword          #
'#  description :  Code File Change Password  #
'#      e-mail  :  adm@rfm.com.ph             #
'#      url     :  http://www.rfm.com.ph      #
'#                                            #
'##############################################

Dim rsLog As DAO.Recordset

Private Sub cmdCancel_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdChange_Click()
   If Me.txtNewPassword <> "" And Me.txtRetypePassword <> "" Then
      If Me.txtRetypePassword = Me.txtNewPassword Then
         Set rsLog = db.OpenRecordset("SELECT * FROM Users")
         rsLog.FindFirst "UserName = '" & pubUserName & "'"
         If Not rsLog.NoMatch Then
            rsLog.Edit
            rsLog.Fields!Password = Me.txtNewPassword
            rsLog.Update
            '---------------------'
            Call Enable_System_Menu
            Unload Me
         End If
      Else
         sagot = MsgBox("The PASSWORD you entered does not match. Try again...", vbInformation, Me.Caption)
         If sagot = vbOK Then
            Call Clear_Set
         End If
      End If
   Else
      sagot = MsgBox("You must enter a PASSWORD.", vbInformation, Me.Caption)
      If sagot = vbOK Then
         Call Clear_Set
      End If
   End If
End Sub

Private Sub Form_Load()
   ' Disable System Menu
   Call Disable_System_Menu
   Me.txtOldPassword.Text = pubPassword
   Me.Top = 250
   Me.Left = 250
   Me.WindowState = 0
End Sub

Private Sub txtNewPassword_Change()
   If Me.txtNewPassword <> "" Then
      If Len(Me.txtNewPassword) = 15 Then
         Me.txtRetypePassword.SetFocus
      End If
   End If
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txtNewPassword <> "" Then
         Me.txtRetypePassword.SetFocus
      End If
   End If
End Sub

Private Sub txtRetypePassword_Change()
   If Me.txtRetypePassword <> "" Then
      If Len(Me.txtRetypePassword) = 15 Then
         If Me.txtRetypePassword <> Me.txtNewPassword Then
            sagot = MsgBox("The PASSWORD you entered does not match. Try again...", vbInformation, Me.Caption)
            If sagot = vbOK Then
               Call Clear_Set
            End If
         Else
            Me.cmdChange.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtRetypePassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txtRetypePassword <> "" Then
         If Me.txtRetypePassword <> Me.txtNewPassword Then
            sagot = MsgBox("The PASSWORD you entered does not match. Try again...", vbInformation, Me.Caption)
            If sagot = vbOK Then
               Call Clear_Set
            End If
         Else
            Me.cmdChange.SetFocus
         End If
      End If
   End If
End Sub

Private Sub Clear_Set()
   Me.txtNewPassword.Text = ""
   Me.txtRetypePassword.Text = ""
   Me.txtNewPassword.SetFocus
End Sub
