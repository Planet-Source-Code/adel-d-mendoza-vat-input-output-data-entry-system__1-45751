VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLogDetails 
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log Details ..."
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
   Icon            =   "frmLogDetails.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7980
   Begin MSComctlLib.ListView lv1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Login"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Logout"
         Object.Width           =   4410
      EndProperty
   End
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
      Left            =   6480
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogDetails"
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
'#           area :  frmLogDetails            #
'#    description :  Code File Log Details    #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim R As DAO.Recordset

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub

Private Sub Form_Load()
  On Error GoTo errF
  Me.Left = 20
  Me.Top = 20
  KeyPreview = True
  Set R = db.OpenRecordset("SELECT * FROM LOGIN WHERE LOGOUT <> 'CURRENT'")
  Me.lv1.ListItems.Clear
  Do While Not R.EOF
     Set x = Me.lv1.ListItems.Add(, , R.Fields!User)
     x.SubItems(1) = R.Fields!Login
     x.SubItems(2) = R.Fields!Logout
     R.MoveNext
  Loop
  Exit Sub
 
errF:
  MsgBox err.Description, vbInformation, Me.Caption
End Sub

