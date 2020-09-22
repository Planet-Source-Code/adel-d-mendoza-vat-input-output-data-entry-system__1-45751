VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « TAX Input"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   ControlBox      =   0   'False
   Icon            =   "frmTax.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10830
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   10575
      Begin VB.Frame Frame5 
         Height          =   2415
         Left            =   9120
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            Picture         =   "frmTax.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            Picture         =   "frmTax.frx":0D0C
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   8895
         Begin VB.TextBox txtTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4560
            TabIndex        =   24
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   " Total Amount :"
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
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   4335
         End
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   3375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Double Click to edit, Press Delete key to delete"
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Number"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vendor Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "T.I.N."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "TC"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "TAX Base"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Debit"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   1665
         Left            =   9120
         Picture         =   "frmTax.frx":114E
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   10575
      Begin VB.Label txtLabel 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   10330
      End
   End
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10575
      Begin VB.TextBox txtDebit 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         MaxLength       =   7
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   4
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtVendorName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         MaxLength       =   30
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   9120
         TabIndex        =   11
         Top             =   120
         Width           =   1335
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear"
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
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "&Add Entry"
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
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.TextBox txtNetAmount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         MaxLength       =   11
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtTIN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtNumber 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Debit Act"
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
         Height          =   375
         Left            =   7200
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tax Base"
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
         Height          =   375
         Left            =   4800
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tax Code"
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
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T.I.N."
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
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTax"
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
'#           area :  frmTAX                   #
'#    description :  Code File TAX input      #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsVESFLE As DAO.Recordset
Dim rsVESDBS As DAO.Recordset
Dim rsVESCOD As DAO.Recordset
Dim rsVESVEN As DAO.Recordset

Private Sub cmdAddItem_Click()
  Dim Valid As Boolean
  Call ValidateData(Valid)
  If Not (Valid) Then
     Exit Sub
  End If

  Set itmFound = Me.lv2.FindItem(txtNumber.Text)
  If itmFound Is Nothing Then
     Call addItem
  Else
     sagot = MsgBox("VENDOR already exist in the list. ADD vendor anyway?", vbQuestion + vbYesNo, Me.Caption)
     If sagot = vbYes Then
        Call addItem
     Else
        Call Clear_Fields
        Me.txtNumber.SetFocus
     End If
  End If
End Sub

Private Sub cmdSave_Click()
   sagot = MsgBox("Are you sure you want to SAVE this transaction?", vbQuestion + vbYesNo, Me.Caption)
   If sagot = vbYes Then
   
      Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & pubYEAR & "' and vesfle.month = '" & pubMONTH & "' and vesfle.code = 'T'")

      ' delete old records to replace by new/updated records
      Do While Not rsVESFLE.EOF
        If rsVESFLE.Fields!Year = pubYEAR And _
           rsVESFLE.Fields!Month = pubMONTH And _
           rsVESFLE.Fields!Code = "T" Then
           '--------------
           rsVESFLE.Delete
         End If
         rsVESFLE.MoveNext
      Loop
   
      ' save new/updated records
      For i = 1 To Me.lv2.ListItems.Count
          rsVESFLE.AddNew
          rsVESFLE.Fields!Year = pubYEAR
          rsVESFLE.Fields!Month = pubMONTH
          rsVESFLE.Fields!Code = "T"
       
          rsVESFLE.Fields!Vendor = Me.lv2.ListItems.Item(i).Text            'Vendor
          rsVESFLE.Fields!Name = Me.lv2.ListItems.Item(i).SubItems(1)       'Vendor name
          rsVESFLE.Fields!Tin = Me.lv2.ListItems.Item(i).SubItems(2)        'TIN number
          rsVESFLE.Fields!Customer = Me.lv2.ListItems.Item(i).SubItems(3)   'Tax Code
          rsVESFLE.Fields!Grsamt = Me.lv2.ListItems.Item(i).SubItems(4)     'Gross Amount
          rsVESFLE.Fields!Acctno = Me.lv2.ListItems.Item(i).SubItems(5)     'GL Acct (Debit)
          
          ' Update Vendor File for TIN number
          Set rsVESVEN = db.OpenRecordset("SELECT * FROM VESVEN WHERE Vendor = '" & Me.lv2.ListItems.Item(i).Text & "'")
          If rsVESVEN.RecordCount <> 0 Then
             rsVESVEN.Edit
             rsVESVEN.Fields!Vatno = Me.lv2.ListItems.Item(i).SubItems(2) 'TIN number
             rsVESVEN.Update
          End If
          ' Update VAT Input File
          rsVESFLE.Update
       Next
       ' Save data then Enable System Menu
       ' and Out to VAT input data entry
       Call Enable_System_Menu
       Unload Me
   End If
End Sub

Private Sub cmdCancel_Click()
   sagot = MsgBox("Are you sure you want to CANCEL this document entry?", vbQuestion + vbYesNo, Me.Caption)
   If sagot = vbYes Then
      Call Enable_System_Menu
      Unload Me
   End If
End Sub

Private Sub cmdClear_Click()
   Call Clear_Fields
   Me.txtNumber.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Private Sub Form_Load()
   Call load_Month
   Me.txtLabel = "WITHOLDING TAX FOR " + iMonth(pubMONTH) + " " + pubYEAR
   Me.Top = 30
   Me.Left = 10
   ' Initialize Fields
   Call Clear_Fields
   ' Load Detail List or Previous Entry
   Call loadLV2
   Me.WindowState = 0
   Call test
End Sub

Private Sub txtNumber_Change()
   On Error Resume Next
   test
   If Me.txtNumber.Text <> "" Then
      Set rsVESVEN = db.OpenRecordset("SELECT name, vatno FROM VESVEN WHERE Vendor = '" & txtNumber.Text & "'")
      If rsVESVEN.RecordCount <> 0 Then
         Me.txtVendorName.Text = rsVESVEN.Fields!Name
                 
         If Left(Trim(rsVESVEN.Fields!Vatno), 1) = "0" Or _
            Left(Trim(rsVESVEN.Fields!Vatno), 1) = "1" Or _
            Left(Trim(rsVESVEN.Fields!Vatno), 1) = "2" Or _
            Left(Trim(rsVESVEN.Fields!Vatno), 1) = "9" Then
            Me.txtTIN.Text = Mid(Trim(rsVESVEN.Fields!Vatno), 1, 9)
         Else
            Me.txtTIN.Text = ""
         End If
         Me.txtTIN.SetFocus
      Else
         If Len(Me.txtNumber) = 7 Then
            sagot = MsgBox("INVALID VENDOR NUMBER. Please check your VENDOR list..", vbInformation, Me.Caption)
            Call Clear_Fields
         End If
      End If
   Else
      Call Clear_Fields
   End If
End Sub

Private Sub txtNumber_GotFocus()
   Me.Frame.Caption = "Press F2 to view Vendor List"
End Sub

Private Sub txtNumber_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
           Set rsVESVEN = db.OpenRecordset("SELECT vesven.name, vesven.vendor FROM VESVEN order by name")
           frmViewVendor.lv1.ListItems.Clear
           rsVESVEN.MoveFirst
           Do While Not rsVESVEN.EOF
              Set x = frmViewVendor.lv1.ListItems.Add(, , rsVESVEN.Fields!Name)
              x.SubItems(1) = rsVESVEN.Fields!Vendor
              rsVESVEN.MoveNext
           Loop
           frmViewVendor.lv1.SetFocus
           frmViewVendor.Show
   End Select
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   test
   If KeyAscii = 13 Then
      If Me.txtNumber.Text <> "" Then
         Set rsVESVEN = db.OpenRecordset("SELECT name, vatno FROM VESVEN WHERE Vendor = '" & txtNumber.Text & "'")
         If rsVESVEN.RecordCount <> 0 Then
            Me.txtVendorName.Text = rsVESVEN.Fields!Name
            Me.txtTIN.Text = rsVESVEN.Fields!Vatno
            Me.txtTIN.SetFocus
         Else
            Call Clear_Fields
         End If
      Else
         Call Clear_Fields
      End If
   End If
End Sub

Private Sub txtNumber_LostFocus()
   Me.Frame.Caption = ""
End Sub

Private Sub txtTIN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Me.txtType.SetFocus
    End If
End Sub

Private Sub Clear_Fields()
   Me.txtNumber.Text = ""
   Me.txtVendorName.Text = ""
   Me.txtTIN.Text = ""
   Me.txtType.Text = ""
   Me.txtNetAmount.Text = ""
   Me.txtDebit.Text = ""
End Sub

Private Sub test()
   If Me.txtNumber.Text <> "" And Me.txtDebit.Text <> "" Then
      Me.cmdAddItem.Enabled = True
   Else
      Me.cmdAddItem.Enabled = False
   End If
End Sub

Private Sub txtType_Change()
   If Me.txtType.Text <> "" Then
      Set rsVESCOD = db.OpenRecordset("SELECT code, desc FROM VESCOD WHERE code = '" & txtType.Text & "'")
      If rsVESCOD.RecordCount <> 0 Then
         Me.txtNetAmount.SetFocus
      Else
         If Len(Me.txtType) = 2 Then
            sagot = MsgBox("INVALID TAX CODE. Please check your TAX CODE list..", vbInformation, Me.Caption)
            If sagot = vbOK Then
               Me.txtType.Text = ""
            End If
         End If
      End If
   End If
End Sub

Private Sub txtType_GotFocus()
   Me.Frame.Caption = "Press F2 to view TAX Codes"
End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txtType.Text <> "" Then
         Set rsVESCOD = db.OpenRecordset("SELECT code, desc FROM VESCOD WHERE code = '" & txtType.Text & "'")
         If rsVESCOD.RecordCount <> 0 Then
            Me.txtNetAmount.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtType_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
           Set rsVESCOD = db.OpenRecordset("SELECT vescod.code, vescod.desc FROM VESCOD order by code")
           frmViewTax.lv1.ListItems.Clear
           rsVESCOD.MoveFirst
           Do While Not rsVESCOD.EOF
              Set x = frmViewTax.lv1.ListItems.Add(, , rsVESCOD.Fields!Code)
              x.SubItems(1) = rsVESCOD.Fields!Desc
              rsVESCOD.MoveNext
           Loop
           frmViewTax.lv1.SetFocus
           frmViewTax.Show
   End Select
End Sub

Private Sub txtType_LostFocus()
   Me.Frame.Caption = ""
End Sub

Private Sub txtNetAmount_Change()
   If Len(Trim(Me.txtNetAmount)) = 11 Then
      Me.txtDebit.SetFocus
   End If
End Sub

Private Sub txtNetAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Me.txtNumber.Text <> "" And Me.txtNetAmount.Text <> "" Then
          Me.txtDebit.SetFocus
       End If
    End If
End Sub

Private Sub txtDebit_Change()
   On Error Resume Next
   test
   If Me.txtDebit.Text <> "" Then
      Set rsVESDBS = db.OpenRecordset("SELECT desc, acctno FROM VESDBS WHERE acctno = '" & txtDebit.Text & "'")
      If rsVESDBS.RecordCount <> 0 Then
         Me.cmdAddItem.SetFocus
      Else
         If Len(Me.txtDebit) = 7 Then
            sagot = MsgBox("INVALID DEBIT ACCOUNT. Please check your DEBIT ACCOUNT list..", vbInformation, Me.Caption)
            If sagot = vbOK Then
               Me.txtDebit.Text = ""
            End If
         End If
      End If
   End If
End Sub

Private Sub txtDebit_GotFocus()
   Me.Frame.Caption = "Press F2 to view DEBIT ACCOUNT list"
End Sub

Private Sub txtDebit_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
           Set rsVESDBS = db.OpenRecordset("SELECT vesdbs.desc, vesdbs.acctno FROM VESDBS order by acctno")
           frmViewDebit.lv1.ListItems.Clear
           rsVESDBS.MoveFirst
           Do While Not rsVESDBS.EOF
              If Len(rsVESDBS.Fields!Desc) <> 0 Then
                 Set x = frmViewDebit.lv1.ListItems.Add(, , rsVESDBS.Fields!Desc)
                 If Len(rsVESDBS.Fields!Acctno) <> 0 Then
                    x.SubItems(1) = rsVESDBS.Fields!Acctno
                 Else
                    x.SubItems(1) = "NONUMB"
                 End If
                 rsVESDBS.MoveNext
              Else
                 Set x = frmViewDebit.lv1.ListItems.Add(, , "NO DESCRIPTION")
                 If Len(rsVESDBS.Fields!Acctno) <> 0 Then
                    x.SubItems(1) = rsVESDBS.Fields!Acctno
                 Else
                    x.SubItems(1) = "NONUMB"
                 End If
                 rsVESDBS.MoveNext
              End If
           Loop
           frmViewDebit.lv1.SetFocus
           frmViewDebit.Show
   End Select
End Sub

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   test
   If KeyAscii = 13 Then
      If Me.txtDebit.Text <> "" Then
         Set rsVESDBS = db.OpenRecordset("SELECT desc, acctno FROM VESDBS WHERE acctno = '" & txtDebit.Text & "'")
         If rsVESDBS.RecordCount <> 0 Then
            Me.cmdAddItem.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtDebit_LostFocus()
   Me.Frame.Caption = ""
End Sub

Private Sub loadLV2()
   Dim valTemp As Double
   
   ' Load Previous Entry. If there is..
   Set rsVESFLE = db.OpenRecordset("SELECT vesfle.vendor, vesven.name, vesfle.grsamt, vesfle.customer, vesfle.tin, vesfle.acctno FROM VESFLE INNER JOIN VESVEN ON vesfle.vendor = vesven.vendor WHERE vesfle.year = '" & pubYEAR & "' and vesfle.month = '" & pubMONTH & "' and vesfle.code = 'T'")
   Me.lv2.ListItems.Clear
   
   valTemp = 0
   Do While Not rsVESFLE.EOF
      Set x = Me.lv2.ListItems.Add(, , rsVESFLE.Fields!Vendor)
      x.SubItems(1) = rsVESFLE.Fields!Name
      x.SubItems(2) = rsVESFLE.Fields!Tin
      x.SubItems(3) = Mid(Trim(rsVESFLE.Fields!Customer), 1, 2)
      x.SubItems(4) = Format(Val(rsVESFLE.Fields!Grsamt), "###,###,##0.00")
      x.SubItems(5) = rsVESFLE.Fields!Acctno
      
      valTemp = valTemp + rsVESFLE.Fields!Grsamt
      rsVESFLE.MoveNext
   Loop
   
   Me.txtTotal.Text = Format(valTemp, "###,###,##0.00")
End Sub

Private Sub lv2_DblClick()
   Dim valTemp As Double
   
   rw = Me.lv2.SelectedItem.Index
   Me.txtNumber.Text = Me.lv2.ListItems.Item(rw).Text
   Me.txtVendorName.Text = Me.lv2.ListItems.Item(rw).SubItems(1)
   Me.txtTIN.Text = Me.lv2.ListItems.Item(rw).SubItems(2)
   Me.txtType.Text = Me.lv2.ListItems.Item(rw).SubItems(3)
   Me.txtNetAmount.Text = Format$(Me.lv2.ListItems.Item(rw).SubItems(4), "#.00")
   Me.txtDebit.Text = Me.lv2.ListItems.Item(rw).SubItems(5)
   
   Me.lv2.ListItems.Remove (rw)
   valTemp = 0
   For i = 1 To Me.lv2.ListItems.Count
       valTemp = valTemp + Me.lv2.ListItems.Item(i).SubItems(4)
   Next
   Me.txtTotal.Text = Format(valTemp, "###,###,##0.00")
   Me.cmdAddItem.Enabled = True
   Me.txtNumber.SetFocus
End Sub

Private Sub addItem()
   Dim valTemp As Double
   
   Set j = Me.lv2.ListItems.Add(, , Me.txtNumber.Text)
   j.SubItems(1) = Me.txtVendorName.Text
   j.SubItems(2) = Me.txtTIN.Text
   j.SubItems(3) = Me.txtType.Text
   j.SubItems(4) = Format(Me.txtNetAmount.Text, "###,###,##0.00")
   j.SubItems(5) = Me.txtDebit.Text
     
   Call Clear_Fields
        
   Me.txtNumber.SetFocus
   valTemp = 0
   For i = 1 To Me.lv2.ListItems.Count
       valTemp = valTemp + Me.lv2.ListItems.Item(i).SubItems(4)
   Next
   Me.txtTotal.Text = Format$(valTemp, "###,###,##0.00")
End Sub

Private Sub ValidateData(AllOK As Boolean)
   Dim Message As String
   AllOK = True
   Message = ""
   If Len(Me.txtNumber.Text) = 0 Then
      Message = "You must enter a Vendor number." + vbCrLf
      Me.txtNumber.SetFocus
      AllOK = False
   ElseIf Len(Me.txtNumber.Text) < 7 Then
      Message = "Vendor number must be 7 characters long." + vbCrLf
      Me.txtNumber.SetFocus
      AllOK = False
   ElseIf Len(Me.txtType.Text) = 0 Then
      Message = "You must enter a Tax code." + vbCrLf
      Me.txtType.SetFocus
      AllOK = False
   ElseIf Len(Me.txtType.Text) < 2 Then
      Message = "Tax code must be 2 characters long." + vbCrLf
      Me.txtType.SetFocus
      AllOK = False
   ElseIf Len(Me.txtDebit.Text) = 0 Then
      Message = "You must enter a Debit Account." + vbCrLf
      Me.txtDebit.SetFocus
      AllOK = False
   ElseIf Len(Me.txtDebit.Text) < 7 Then
      Message = "Debit Account must be 7 characters long." + vbCrLf
      Me.txtDebit.SetFocus
      AllOK = False
   End If
   If Not (AllOK) Then
      MsgBox Message, vbOKOnly + vbInformation, "Validation Error"
   End If
End Sub

