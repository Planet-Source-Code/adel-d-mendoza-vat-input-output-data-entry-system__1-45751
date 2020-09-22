VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVESVEN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Vendor Maintenance"
   ClientHeight    =   6270
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
   ScaleHeight     =   6270
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
      Left            =   2280
      TabIndex        =   12
      Top             =   3360
      Width           =   3405
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2400
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   "Double click an item to edit or delete."
      Top             =   3840
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
         Text            =   "Vendor Name"
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
      Width           =   1215
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
      Left            =   5520
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6900
      Begin VB.TextBox txtVendor 
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
         MaxLength       =   7
         TabIndex        =   1
         Top             =   720
         Width           =   1215
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
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Vendor Name"
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
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "VAT  Number"
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
         Top             =   2160
         Width           =   1215
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
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vendors"
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
         Width           =   1260
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Search for Vendor Name:"
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
      Top             =   3360
      Width           =   2070
   End
End
Attribute VB_Name = "frmVESVEN"
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
'#           area :  frmVESVEN                #
'#    description :  Code File                #
'#                :  Vendor Maintenance       #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsVESVEN As DAO.Recordset
Dim txt As Control
Dim esc As Byte

Private Sub cmdAdd_Click()
   On Error GoTo errHandler
   If cmdAdd.Caption = "&Add" Then
      rsVESVEN.AddNew
      Frame1.Enabled = True
      cmdDelete.Enabled = False
      cmdUpdate.Enabled = False
      
      lv1.Enabled = False
      txtSearch.Locked = True
      cmdRefresh.Enabled = False

      Call clrTxt
      txtVendor.SetFocus
      cmdAdd.Caption = "&Save"
      cmdClose.Caption = "&Cancel"
      esc = 1
   Else
      If txtVendor.Text <> "" And txtName.Text <> "" And _
         txtAddress1.Text <> "" And txtAddress2.Text <> "" And _
         txtVatNo.Text <> "" Then
         
         Call ConFields
         rsVESVEN.Update
         rsVESVEN.MoveFirst
         Call loadLV
         MsgBox "New vendor has been added.", vbInformation, Me.Caption
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
MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdClose_Click()
   If cmdClose.Caption = "&Close" Then
      Call Enable_System_Menu
      Unload Me
   Else
      rsVESVEN.CancelUpdate
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
   If txtVendor.Text <> "" Then
      Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN WHERE Vendor = '" & txtVendor.Text & "'")
      Reply = MsgBox("Code: " & rsVESVEN.Fields!Vendor & Chr(13) & "Name: " & rsVESVEN.Fields!Name & Chr(13) & "Are you sure you want to delete this Vendor?", vbQuestion + vbYesNo, Me.Caption)
      If Reply = vbYes Then
         rsVESVEN.Delete
         Call clrTxt
         Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN")
         Call loadLV
         cmdAdd.Enabled = True
         Frame1.Enabled = False
         MsgBox "Vendor deleted.", vbInformation, Me.Caption
      End If
   Else
      MsgBox "No Vendor to delete.", vbInformation, Me.Caption
   End If
End Sub

Private Sub cmdRefresh_Click()
   txtSearch.Text = ""
End Sub

Private Sub cmdUpdate_Click()
   If txtVendor.Text <> "" And txtName.Text <> "" And _
      txtAddress1.Text <> "" And txtAddress2.Text <> "" And _
      txtVatNo.Text <> "" Then
      Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN where vendor = '" & txtVendor.Text & "'")
      Call updateRec
      Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN")
      Frame1.Enabled = False
      cmdAdd.Enabled = True
      Call clrTxt
      Call loadLV
      MsgBox "Vendor updated", vbInformation, Me.Caption
      Exit Sub
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
               rsVESVEN.CancelUpdate
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
   Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN")
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
   rsVESVEN.Fields!Vendor = UCase(txtVendor.Text)
   rsVESVEN.Fields!Name = UCase(txtName.Text)
   rsVESVEN.Fields!Address1 = UCase(txtAddress1.Text)
   rsVESVEN.Fields!Address2 = UCase(txtAddress2.Text)
   rsVESVEN.Fields!Vatno = UCase(txtVatNo.Text)
End Sub

Private Sub RetFields()
   If rsVESVEN.Fields!Vendor <> "" Then
      txtVendor.Text = rsVESVEN.Fields!Vendor
   Else
      txtVendor.Text = ""
   End If
   If rsVESVEN.Fields!Name <> "" Then
      txtName.Text = rsVESVEN.Fields!Name
   Else
      txtVendor.Text = ""
   End If
   If rsVESVEN.Fields!Address1 <> "" Then
      txtAddress1.Text = rsVESVEN.Fields!Address1
   Else
      txtAddres1.Text = ""
   End If
   If rsVESVEN.Fields!Address2 <> "" Then
      txtAddress2.Text = rsVESVEN.Fields!Address2
   Else
      txtAddress2.Text = ""
   End If
   If rsVESVEN.Fields!Vatno <> "" Then
      txtVatNo.Text = rsVESVEN.Fields!Vatno
   Else
      txtVatNo.Text = ""
   End If
End Sub

Private Sub updateRec()
   rsVESVEN.Edit
   rsVESVEN.Fields!Vendor = UCase(txtVendor.Text)
   rsVESVEN.Fields!Name = UCase(txtName.Text)
   rsVESVEN.Fields!Address1 = UCase(txtAddress1.Text)
   rsVESVEN.Fields!Address2 = UCase(txtAddress2.Text)
   rsVESVEN.Fields!Vatno = UCase(txtVatNo.Text)
   rsVESVEN.Update
End Sub

Private Sub loadLV()
   lv1.ListItems.Clear
   Do While Not rsVESVEN.EOF
      Set j = Me.lv1.ListItems.Add(, , rsVESVEN.Fields!Vendor)
      j.SubItems(1) = rsVESVEN.Fields!Name
      rsVESVEN.MoveNext
   Loop
End Sub

Private Sub lv1_DblClick()
   If lv1.ListItems.Count <> 0 Then
      Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN WHERE Vendor = '" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'")
      Call RetFields
      Frame1.Enabled = True
      txtVendor.SetFocus
      SendKeys "{home}+{end}"
      cmdAdd.Enabled = False
   End If
End Sub

Private Sub txtSearch_Change()
   Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN WHERE Name Like '" & txtSearch.Text & "*'")
   Call loadLV
End Sub

