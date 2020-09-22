VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVESCUS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Customer Maintenance"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7110
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Top             =   3360
      Width           =   3090
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2400
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   "Double click an item to edit or delete."
      Top             =   3780
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   9526
      EndProperty
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
      Height          =   330
      Left            =   2880
      TabIndex        =   7
      Top             =   2760
      Width           =   1230
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   2760
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
      Height          =   330
      Left            =   1560
      TabIndex        =   6
      Top             =   2760
      Width           =   1230
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
      Height          =   330
      Left            =   5505
      TabIndex        =   9
      Top             =   2760
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6900
      Begin VB.TextBox txtCustomer 
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
         MaxLength       =   5
         TabIndex        =   1
         Top             =   720
         Width           =   615
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
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtVatNo 
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
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtAddress2 
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
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox txtAddress1 
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
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "VAT Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Address 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Address 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cust Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Customers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   180
         TabIndex        =   10
         Top             =   180
         Width           =   1665
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Search for Customer Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   3420
      Width           =   2295
   End
End
Attribute VB_Name = "frmVESCUS"
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
'#           area :  frmVESCUS                #
'#    Nameription :  Code File                #
'#                :  Customer Maintenance     #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsVESCUS As DAO.Recordset
Dim txt As Control
Dim esc As Byte

Private Sub cmdAdd_Click()
   On Error GoTo errHandler
   If cmdAdd.Caption = "&Add" Then
      rsVESCUS.AddNew
      Frame1.Enabled = True
      cmdDelete.Enabled = False
      cmdUpdate.Enabled = False
      
      lv1.Enabled = False
      txtSearch.Locked = True
      cmdRefresh.Enabled = False

      Call clrTxt
      txtCustomer.SetFocus
      cmdAdd.Caption = "&Save"
      cmdClose.Caption = "&Cancel"
      esc = 1
   Else
      If txtCustomer.Text <> "" And txtName.Text <> "" And _
         txtAddress1.Text <> "" And txtAddress2.Text <> "" And _
         txtVatNo.Text <> "" Then
         
         Call ConFields
         rsVESCUS.Update
         rsVESCUS.MoveFirst
         Call loadLV
         MsgBox "New customer has been added.", vbInformation, Me.Caption
         Call clrTxt
         cmdAdd.Caption = "&Add"
         cmdClose.Caption = "&Close"
         
         Frame1.Enabled = False
         cmdDelete.Enabled = True
         cmdUpdate.Enabled = True
         
         lv1.Enabled = True
         txtSearch.Locked = False
         cmdRefresh.Enabled = True
         
         esc = 0
      Else
         MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
         Exit Sub
      End If
   End If
   Exit Sub
   
errHandler:
MsgBox err.Nameription, vbInformation, Me.Caption
End Sub

Private Sub cmdClose_Click()
   If cmdClose.Caption = "&Close" Then
      Call Enable_System_Menu
      Unload Me
   Else
      rsVESCUS.CancelUpdate
      Call clrTxt
      cmdAdd.Caption = "&Add"
      cmdClose.Caption = "&Close"
      Frame1.Enabled = False
      cmdDelete.Enabled = True
      cmdUpdate.Enabled = True
      lv1.Enabled = True
      txtSearch.Locked = False
      cmdRefresh.Enabled = True
      esc = 0
   End If
End Sub

Private Sub cmdDelete_Click()
   If txtCustomer.Text <> "" Then
      Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS WHERE Customer = '" & txtCustomer.Text & "'")
      Reply = MsgBox("Code: " & rsVESCUS.Fields!Customer & Chr(13) & "Name: " & rsVESCUS.Fields!Name & Chr(13) & "Are you sure you want to delete this Customer?", vbQuestion + vbYesNo, Me.Caption)
      If Reply = vbYes Then
         rsVESCUS.Delete
         Call clrTxt
         Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS")
         Call loadLV
         cmdAdd.Enabled = True
         Frame1.Enabled = False
         MsgBox "Debit Account deleted.", vbInformation, Me.Caption
      End If
   Else
      MsgBox "No Debit Account to delete.", vbInformation, Me.Caption
   End If
End Sub

Private Sub cmdRefresh_Click()
   txtSearch.Text = ""
End Sub

Private Sub cmdUpdate_Click()
   If txtCustomer.Text <> "" And txtName.Text <> "" And _
      txtAddress1.Text <> "" And txtAddress2.Text <> "" And _
      txtVatNo.Text <> "" Then
      
      Call updateRec
      Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS")
      Frame1.Enabled = False
      cmdAdd.Enabled = True
      Call clrTxt
      Call loadLV
      MsgBox "Customer updated", vbInformation, Me.Caption
   Else
      MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
      Exit Sub
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
               rsVESCUS.CancelUpdate
               Call clrTxt
               cmdAdd.Caption = "&Add"
               cmdClose.Caption = "&Close"
               Frame1.Enabled = False
               cmdDelete.Enabled = True
               cmdUpdate.Enabled = True
               lv1.Enabled = True
               txtSearch.Locked = False
               cmdRefresh.Enabled = True
               esc = 0
            End If
         End If
   End Select
End Sub

Private Sub Form_Load()
   Call Disable_System_Menu
   Me.Top = 500
   Me.Left = 2300
   Me.WindowState = 0
   Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS")
   esc = 0
   Call loadLV
End Sub

Private Sub clrTxt()
   For Each txt In Me.Controls
       If TypeOf txt Is TextBox Then
          txt.Text = ""
       End If
   Next
End Sub

Private Sub ConFields()
   rsVESCUS.Fields!Customer = UCase(txtCustomer.Text)
   rsVESCUS.Fields!Name = UCase(txtName.Text)
   rsVESCUS.Fields!Address1 = UCase(txtAddress1.Text)
   rsVESCUS.Fields!Address2 = UCase(txtAddress2.Text)
   rsVESCUS.Fields!Vatno = UCase(txtVatNo.Text)
End Sub

Private Sub RetFields()
   If rsVESCUS.Fields!Customer <> "" Then
      txtCustomer.Text = rsVESCUS.Fields!Customer
   Else
      txtCustomer.Text = ""
   End If
   If rsVESCUS.Fields!Name <> "" Then
      txtName.Text = rsVESCUS.Fields!Name
   Else
      txtName.Text = ""
   End If
   If rsVESCUS.Fields!Address1 <> "" Then
      txtAddress1.Text = rsVESCUS.Fields!Address1
   Else
      txtAddress1.Text = ""
   End If
   If rsVESCUS.Fields!Address2 <> "" Then
      txtAddress2.Text = rsVESCUS.Fields!Address2
   Else
      txtAddress2.Text = ""
   End If
   If rsVESCUS.Fields!Vatno <> "" Then
      txtVatNo.Text = rsVESCUS.Fields!Vatno
   Else
      txtVatNo.Text = ""
   End If
End Sub

Private Sub updateRec()
   rsVESCUS.Edit
   rsVESCUS.Fields!Customer = UCase(txtCustomer.Text)
   rsVESCUS.Fields!Name = UCase(txtName.Text)
   rsVESCUS.Fields!Address1 = UCase(txtAddress1.Text)
   rsVESCUS.Fields!Address2 = UCase(txtAddress2.Text)
   rsVESCUS.Fields!Vatno = UCase(txtVatNo.Text)
   rsVESCUS.Update
End Sub

Private Sub loadLV()
   lv1.ListItems.Clear
   Do While Not rsVESCUS.EOF
      Set j = lv1.ListItems.Add(, , rsVESCUS.Fields!Customer)
      If rsVESCUS.Fields!Name <> "" Then
         j.SubItems(1) = rsVESCUS.Fields!Name
      Else
         j.SubItems(1) = "*** NO NAME ***"
      End If
      rsVESCUS.MoveNext
   Loop
End Sub

Private Sub lv1_DblClick()
   If lv1.ListItems.Count <> 0 Then
      Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS WHERE Customer = '" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'")
      Call RetFields
      Frame1.Enabled = True
      txtCustomer.SetFocus
      SendKeys "{home}+{end}"
      cmdAdd.Enabled = False
   End If
End Sub

Private Sub txtSearch_Change()
   Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS WHERE Name Like '" & txtSearch.Text & "*'")
   Call loadLV
End Sub


