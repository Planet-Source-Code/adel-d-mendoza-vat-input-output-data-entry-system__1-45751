VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewDebit 
   BackColor       =   &H00000040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Debit Accounts"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click or Enter to select"
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11456
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmViewDebit"
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
'#         area :  frmViewDebit               #
'#  description :  Code File View Debit Acct  #
'#      e-mail  :  adm@rfm.com.ph             #
'#      url     :  http://www.rfm.com.ph      #
'#                                            #
'##############################################

Private Sub Form_Load()
  Me.Top = 50
  Me.Left = 5100
  Me.WindowState = 0
End Sub

Private Sub lv1_DblClick()
   rw = Me.lv1.SelectedItem.Index
   frmTax.txtDebit.Text = Me.lv1.ListItems.Item(rw).SubItems(1)
   
   On Error Resume Next
   Call test
   frmTax.cmdAddItem.SetFocus
   Unload Me
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13
          rw = Me.lv1.SelectedItem.Index
          frmTax.txtDebit.Text = Me.lv1.ListItems.Item(rw).SubItems(1)
          
          On Error Resume Next
          Call test
          frmTax.cmdAddItem.SetFocus
          Unload Me
   End Select
End Sub

Private Sub test()
   If frmTax.txtNumber.Text <> "" And frmTax.txtDebit.Text <> "" Then
      frmTax.cmdAddItem.Enabled = True
   Else
      frmTax.cmdAddItem.Enabled = False
   End If
End Sub

