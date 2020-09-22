VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "Lvbuttons.ocx"
Begin VB.Form frmUserPass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter User name and Password to Login ..."
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmUserPass.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unmask the Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Login ..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "frmUserPass.frx":0E42
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   120
      Picture         =   "frmUserPass.frx":0E5E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WARNING : Be sure when you use this option, it will show password in alphanumeric characters."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserPass"
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
'#           area :  frmUserPass              #
'#    description :  Code File Login          #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsLog As DAO.Recordset
Dim rsUSER As DAO.Recordset
  
Private Sub Check1_Click()
  If Check1.Value = 1 Then
    Text1.PasswordChar = ""
  Else
    Text1.PasswordChar = "*"
  End If
End Sub

Private Sub Combo1_Click()
  Text1.Text = Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     On Error Resume Next
     Sagot = MsgBox("Are you sure you want to close the application?", vbQuestion Or vbYesNo, "Want to Close Application ..")
     If Sagot = vbYes Then
        Unload Me
     End If
  End If
End Sub

Private Sub Form_Load()
  'open database file
  openDB
  'initialize form
  Dim rsUSR As DAO.Recordset
  Set rsUSR = db.OpenRecordset("SELECT * FROM USERS")
  Combo1.Clear
  Do While Not rsUSR.EOF
    Combo1.addItem rsUSR.Fields!UserName
    rsUSR.MoveNext
  Loop
  rsUSR.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  db.Close
End Sub

Private Sub LaVolpeButton1_Click()
  If Len(Combo1.Text) > 0 Then
     Set rsUSER = db.OpenRecordset("SELECT * FROM USERS WHERE USERNAME = '" & Me.Combo1.Text & "'")
     If rsUSER.RecordCount <> 0 Then
        If UCase(rsUSER.Fields!Password) = UCase(Text1.Text) Then
        
           pubPassword = rsUSER.Fields!Password
           pubUserName = rsUSER.Fields!UserName
           pubAccLevel = rsUSER.Fields!AccLevel
           
           mdiAPAY.cur_user = Combo1.Text
           mdiAPAY.Label1(1).Caption = LCase(Combo1.Text)
           mdiAPAY.Label1(3).Caption = Time
           
           Set rsLog = db.OpenRecordset("SELECT * FROM LOGIN")
           rsLog.AddNew
           rsLog.Fields!User = Combo1.Text
           rsLog.Fields!Login = Now
           rsLog.Fields!Logout = "CURRENT"
           rsLog.Update

           Me.Visible = False
           Text1.Text = Clear
           'check access level
           Call Check_Access_Level
           'call main menu
           mdiAPAY.Show
        Else
           MsgBox "Invalid Password ...", vbCritical, "Enter Proper Password ..."
           Text1.Text = Clear
        End If
     Else
        MsgBox "Username Not Selected ...", vbCritical
     End If
  End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    Call LaVolpeButton1_Click
  End If
End Sub

Private Sub Check_Access_Level()
   If pubAccLevel < 3 Then
      mdiAPAY.mnuVendorMaintenance.Enabled = False
      mdiAPAY.mnuCustomerMaintenance.Enabled = False
      mdiAPAY.mnuBranchMaintenance.Enabled = False
      mdiAPAY.mnuDebitAcctMaintenance.Enabled = False
      mdiAPAY.mnuTaxCodeMaintenance.Enabled = False
      mdiAPAY.mnuClearLogDetails.Enabled = False
      mdiAPAY.btnBranch.Enabled = False
      mdiAPAY.btnCustomer.Enabled = False
      mdiAPAY.btnDebitAcct.Enabled = False
      mdiAPAY.btnTaxCodes.Enabled = False
      mdiAPAY.btnVendor.Enabled = False
   End If
End Sub

