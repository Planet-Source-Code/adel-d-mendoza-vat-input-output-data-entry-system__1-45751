VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « User's Information"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7290
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
      Height          =   420
      Left            =   5625
      TabIndex        =   8
      Top             =   2115
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4365
      TabIndex        =   5
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5625
      TabIndex        =   7
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4365
      TabIndex        =   6
      Top             =   2115
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   7020
      Begin VB.ComboBox cboAccLevel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmUsers.frx":0000
         Left            =   1440
         List            =   "frmUsers.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1755
         Width           =   1410
      End
      Begin VB.TextBox txtPosition 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   2
         Top             =   1395
         Width           =   1410
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1035
         Width           =   5295
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   0
         Top             =   675
         Width           =   2130
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   15
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   14
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User's Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   12
         Top             =   180
         Width           =   2550
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Access Level"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   11
         Top             =   1800
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   10
         Top             =   2205
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmUsers"
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
'#           area :  frmUsers                 #
'#    description :  Code File                #
'#                :  User's Maintenance       #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsUSER As DAO.Recordset
Dim txt As Control
Dim esc As Byte
Dim tempUser As String

Private Sub cmdAdd_Click()
   On Error GoTo errHandler
   If cmdAdd.Caption = "&Add" Then
      rsUSER.AddNew
      Frame1.Enabled = True
      cmdDelete.Enabled = False
      cmdUpdate.Enabled = False
      Call clrTxt
      If pubAccLevel < 3 Then
         Me.cboAccLevel.Text = 1
         Me.cboAccLevel.Enabled = False
      End If
      txtUserName.SetFocus
      cmdAdd.Caption = "&Save"
      cmdClose.Caption = "&Cancel"
      esc = 1
   Else
      If txtUserName.Text <> "" And txtName.Text <> "" _
         And txtPosition.Text <> "" And cboAccLevel.Text <> "" _
         And txtPassword.Text <> "" Then
         Call ConFields
         rsUSER.Update
         MsgBox "New user has been added.", vbInformation, Me.Caption
         Call clrTxt
         If pubAccLevel < 3 Then
            Me.cboAccLevel.Enabled = True
         End If
         cmdAdd.Caption = "&Add"
         cmdClose.Caption = "&Close"
         Frame1.Enabled = False
         cmdDelete.Enabled = True
         cmdUpdate.Enabled = True
         esc = 0
      Else
         MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
         Exit Sub
      End If
   End If
   Exit Sub
   
errHandler:
If err.Number = 3022 Then
   MsgBox "User Name: " & txtUserName.Text & " already exists.", vbInformation, Me.Caption
Else
   MsgBox err.Description, vbInformation, Me.Caption
End If
End Sub

Private Sub cmdClose_Click()
   If cmdClose.Caption = "&Close" Then
      Call Enable_System_Menu
      Unload Me
   Else
      rsUSER.CancelUpdate
      Call clrTxt
      If pubAccLevel < 3 Then
         Me.cboAccLevel.Enabled = True
      End If
      cmdAdd.Caption = "&Add"
      cmdClose.Caption = "&Close"
      Frame1.Enabled = False
      cmdDelete.Enabled = True
      cmdUpdate.Enabled = True
      esc = 0
   End If
End Sub

Private Sub cmdDelete_Click()
   If txtUserName.Text <> "" Then
      Set rsUSER = db.OpenRecordset("SELECT * FROM Users WHERE UserName = '" & txtUserName.Text & "'")
      Reply = MsgBox("User Name: " & rsUSER.Fields!UserName & Chr(13) & "Password: " & String(Len(rsUSER.Fields!Password), "*") & Chr(13) & "Are you sure you want to Delete this user?", vbQuestion + vbYesNo, Me.Caption)
      If Reply = vbYes Then
         rsUSER.Delete
         Call clrTxt
         MsgBox "User deleted.", vbInformation, Me.Caption
         cmdUpdate.Caption = "&Search"
         cmdAdd.Enabled = True
         Frame1.Enabled = False
         Set rsUSER = db.OpenRecordset("SELECT * FROM Users")
      End If
   Else
      MsgBox "No user to delete.", vbInformation, Me.Caption
   End If
End Sub

Private Sub cmdUpdate_Click()
   If cmdUpdate.Caption = "&Search" Then
      entry = InputBox("Enter User Name: ", Me.Caption)
      rsUSER.FindFirst "UserName = '" & entry & "'"
      If rsUSER.NoMatch = False Then
         Call RetFields
         cmdUpdate.Caption = "&Update"
         Frame1.Enabled = True
         cmdAdd.Enabled = False
         tempUser = rsUSER.Fields!UserName
         If pubAccLevel < 3 Then
            Me.cboAccLevel.Enabled = False
            If UCase(entry) = "ADMINISTRATOR" Then
               Me.txtUserName.Locked = True
               Me.txtName.Locked = True
               Me.txtPosition.Locked = True
               Me.txtPassword.Locked = True
               Me.cmdDelete.Enabled = False
               Me.cmdUpdate.Enabled = False
            End If
         End If
      Else
         MsgBox "User Name: " & entry & " not found!", vbInformation, Me.Caption
      End If
   Else
      If txtUserName.Text <> "" And txtName.Text <> "" _
         And txtPosition.Text <> "" And cboAccLevel.Text <> "" _
         And txtPassword.Text <> "" Then
         If tempUser = txtUserName.Text Then
            Call updateRec
            MsgBox "User updated", vbInformation, Me.Caption
            Set rsUSER = db.OpenRecordset("SELECT * FROM Users")
            Frame1.Enabled = False
            cmdAdd.Enabled = True
            cmdUpdate.Caption = "&Search"
            Call clrTxt
            If pubAccLevel < 3 Then
               Me.cboAccLevel.Enabled = True
               If UCase(entry) = "ADMINISTRATOR" Then
                  Me.txtUserName.Locked = False
                  Me.txtName.Locked = False
                  Me.txtPosition.Locked = False
                  Me.txtPassword.Locked = False
                  Me.cmdDelete.Enabled = True
                  Me.cmdUpdate.Enabled = False
               End If
            End If
            Exit Sub
         End If
         If tempUser <> txtUserName.Text Then
            Set rsUSER = db.OpenRecordset("SELECT * FROM Users WHERE UserName = '" & txtUserName.Text & "'")
            If rsUSER.BOF = True Then
               Call updateRec
               MsgBox "User updated", vbInformation, Me.Caption
               Set rsUSER = db.OpenRecordset("SELECT * FROM Users")
               Frame1.Enabled = False
               cmdAdd.Enabled = True
               cmdUpdate.Caption = "&Search"
               Call clrTxt
               If pubAccLevel < 3 Then
                  Me.cboAccLevel.Enabled = True
                  If UCase(entry) = "ADMINISTRATOR" Then
                     Me.txtUserName.Locked = False
                     Me.txtName.Locked = False
                     Me.txtPosition.Locked = False
                     Me.txtPassword.Locked = False
                     Me.cmdDelete.Enabled = True
                     Me.cmdUpdate.Enabled = True
                  End If
               End If
            Else
               MsgBox "User Name: " & txtUserName.Text & " already exists!", vbInformation, Me.Caption
               txtUserName.SetFocus
               SendKeys "{home}+{end}"
               Exit Sub
            End If
            Exit Sub
         End If
      Else
         MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
         Exit Sub
      End If
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13
         SendKeys "{tab}"
      Case vbKeyEscape
         If esc = 1 Then
            Reply = MsgBox("Cancel operation?", vbQuestion + vbYesNo, Me.Caption)
            If Reply = vbYes Then
               rsUSER.CancelUpdate
               Call clrTxt
               If pubAccLevel < 3 Then
                  Me.cboAccLevel.Enabled = True
               End If
               cmdAdd.Caption = "&Add"
               If cmdClose.Caption = "&Cancel" Then
                  cmdClose.Caption = "&Close"
               End If
               Frame1.Enabled = False
               cmdDelete.Enabled = True
               cmdUpdate.Enabled = True
               esc = 0
            End If
         End If
   End Select
End Sub

Private Sub Form_Load()
   Me.Top = 2000
   Me.Left = 2250
   Me.WindowState = 0
   Call Disable_System_Menu
   Set rsUSER = db.OpenRecordset("SELECT * FROM Users")
   esc = 0
End Sub

Private Sub clrTxt()
   For Each txt In Me.Controls
       If TypeOf txt Is TextBox Then
          txt.Text = ""
       End If
   Next
   cboAccLevel.ListIndex = -1
End Sub

Private Sub ConFields()
   rsUSER.Fields!UserName = txtUserName.Text
   rsUSER.Fields!UserFullName = txtName.Text
   rsUSER.Fields!UserPosition = txtPosition.Text
   rsUSER.Fields!AccLevel = cboAccLevel.Text
   rsUSER.Fields!Password = txtPassword.Text
End Sub

Private Sub RetFields()
   txtUserName.Text = rsUSER.Fields!UserName
   txtName.Text = rsUSER.Fields!UserFullName
   txtPosition.Text = rsUSER.Fields!UserPosition
   cboAccLevel.Text = rsUSER.Fields!AccLevel
   txtPassword.Text = rsUSER.Fields!Password
End Sub

Private Sub updateRec()
   sqlStr = "UPDATE Users SET " _
     & ", UserName = '" & UCase(txtUserName.Text) & "'" _
     & ", UserFullName = '" & txtName.Text & "'" _
     & ", UserPosition = '" & txtPosition.Text & "'" _
     & ", AccLevel = '" & cboAccLevel.Text & "'" _
     & ", Password = '" & txtPassword.Text & "' WHERE UserName = '" & txtUserName.Text & "'"
     execQuery (sqlStr)
End Sub
