VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVESDBS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Debit Accounts Maintenance"
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
      TabIndex        =   12
      Top             =   2280
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
      Left            =   2400
      TabIndex        =   10
      Top             =   2280
      Width           =   3210
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3480
      Left            =   90
      TabIndex        =   9
      ToolTipText     =   "Double click an item to edit or delete."
      Top             =   2700
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   6138
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
         Text            =   "Account Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
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
      TabIndex        =   3
      Top             =   1710
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
      TabIndex        =   4
      Top             =   1710
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
      TabIndex        =   2
      Top             =   1710
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
      TabIndex        =   5
      Top             =   1710
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
      Height          =   2085
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6900
      Begin VB.TextBox txtAcctNo 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   675
         Width           =   810
      End
      Begin VB.TextBox txtDesc 
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
         TabIndex        =   1
         Top             =   1035
         Width           =   4650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Code"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Debit Accounts"
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
         TabIndex        =   6
         Top             =   180
         Width           =   2325
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Search for Account Name:"
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
      TabIndex        =   11
      Top             =   2340
      Width           =   2130
   End
End
Attribute VB_Name = "frmVESDBS"
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
'#           area :  frmVESDBS                #
'#    description :  Code File                #
'#                :  Debit Acct Maintenance   #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsVESDBS As DAO.Recordset
Dim txt As Control
Dim esc As Byte

Private Sub cmdAdd_Click()
   On Error GoTo errHandler
   If cmdAdd.Caption = "&Add" Then
      rsVESDBS.AddNew
      Frame1.Enabled = True
      cmdDelete.Enabled = False
      cmdUpdate.Enabled = False
      
      lv1.Enabled = False
      txtSearch.Locked = True
      cmdRefresh.Enabled = False

      Call clrTxt
      txtAcctNo.SetFocus
      cmdAdd.Caption = "&Save"
      cmdClose.Caption = "&Cancel"
      esc = 1
   Else
      If txtAcctNo.Text <> "" And txtDesc.Text <> "" Then
         Call ConFields
         rsVESDBS.Update
         rsVESDBS.MoveFirst
         Call loadLV
         MsgBox "New debit account has been added.", vbInformation, Me.Caption
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
      rsVESDBS.CancelUpdate
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
   If txtAcctNo.Text <> "" Then
      Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS WHERE Acctno = '" & txtAcctNo.Text & "'")
      Reply = MsgBox("Account Code: " & rsVESDBS.Fields!Acctno & Chr(13) & "Description: " & rsVESDBS.Fields!Desc & Chr(13) & "Are you sure you want to Delete this debit account?", vbQuestion + vbYesNo, Me.Caption)
      If Reply = vbYes Then
         rsVESDBS.Delete
         Call clrTxt
         Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS")
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
   If txtAcctNo.Text <> "" And txtDesc.Text <> "" Then
      Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS where acctno = '" & txtAcctNo.Text & "'")
      Call updateRec
      Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS")
      Frame1.Enabled = False
      cmdAdd.Enabled = True
      Call clrTxt
      Call loadLV
      MsgBox "Debit Account updated", vbInformation, Me.Caption
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
               rsVESDBS.CancelUpdate
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
   Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS")
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
   rsVESDBS.Fields!Acctno = txtAcctNo.Text
   rsVESDBS.Fields!Desc = UCase(txtDesc.Text)
End Sub

Private Sub RetFields()
   txtAcctNo.Text = rsVESDBS.Fields!Acctno
   If rsVESDBS.Fields!Desc <> "" Then
      txtDesc.Text = rsVESDBS.Fields!Desc
   Else
      txtDesc.Text = "--- NO ENTRY ---"
   End If
End Sub

Private Sub updateRec()
   rsVESDBS.Edit
   rsVESDBS.Fields!Acctno = txtAcctNo.Text
   rsVESDBS.Fields!Desc = UCase(txtDesc.Text)
   rsVESDBS.Update
End Sub

Private Sub loadLV()
   lv1.ListItems.Clear
   Do While Not rsVESDBS.EOF
      Set j = lv1.ListItems.Add(, , rsVESDBS.Fields!Acctno)
      If rsVESDBS.Fields!Desc <> "" Then
         j.SubItems(1) = rsVESDBS.Fields!Desc
      Else
         j.SubItems(1) = "--- NO ENTRY ---"
      End If
      rsVESDBS.MoveNext
   Loop
End Sub

Private Sub lv1_DblClick()
   If lv1.ListItems.Count <> 0 Then
      Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS WHERE Acctno = '" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'")
      Call RetFields
      Frame1.Enabled = True
      txtAcctNo.SetFocus
      SendKeys "{home}+{end}"
      cmdAdd.Enabled = False
   End If
End Sub

Private Sub txtSearch_Change()
   Set rsVESDBS = db.OpenRecordset("SELECT * FROM VESDBS WHERE Desc Like '" & txtSearch.Text & "*'")
   Call loadLV
End Sub
