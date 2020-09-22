VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewBranch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Branch"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lv1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click or Enter to select"
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
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
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Branch Name"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "frmViewBranch"
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
'#           area :  frmViewBranch            #
'#    description :  Code File View Branch    #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Private Sub Form_Load()
  Me.Top = 250
  Me.Left = 5800
  Me.WindowState = 0
End Sub

Private Sub lv1_DblClick()
   rw = Me.lv1.SelectedItem.Index
   frmBranch.txtBranchCode = Me.lv1.ListItems.Item(rw).Text
   frmBranch.txtBranchName = Me.lv1.ListItems.Item(rw).SubItems(1)
   Unload Me
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13
        rw = Me.lv1.SelectedItem.Index
        frmBranch.txtBranchCode = Me.lv1.ListItems.Item(rw).Text
        frmBranch.txtBranchName = Me.lv1.ListItems.Item(rw).SubItems(1)
        Unload Me
   End Select
End Sub


