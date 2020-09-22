VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVAToutput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « VAT Output"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10830
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
      Height          =   735
      Left            =   4920
      Picture         =   "frmVAToutput.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Width           =   1695
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
      Height          =   735
      Left            =   3120
      Picture         =   "frmVAToutput.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   6720
      TabIndex        =   14
      Top             =   6120
      Width           =   3975
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   3495
      Left            =   3120
      TabIndex        =   13
      Top             =   2640
      Width           =   7575
      Begin MSComctlLib.ListView lv2 
         Height          =   3135
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Double Click to edit, Press Delete key to delete"
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   5530
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Net Amount"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   3120
      TabIndex        =   12
      Top             =   960
      Width           =   7575
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
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
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
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtNetAmount 
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
         Left            =   2040
         MaxLength       =   11
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtCustomerName 
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
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   2
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txtNumber 
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
         Left            =   2040
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Width           =   7575
      Begin VB.Label txtLabel 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
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
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "&Enable List"
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
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Double Click or Enter to select"
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   10186
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Customer List"
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
      TabIndex        =   10
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmVAToutput"
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
'#           area : frmVAToutput              #
'#    description : Code File VAT output      #
'#                : VAT output (exempt)       #
'#    cType codes : 1 for VAT output          #
'#                : 3 for VAT output (exempt) #
'#        e-mail  : adm@rfm.com.ph            #
'#        url     : http://www.rfm.com.ph     #
'#                                            #
'##############################################

Dim rsVESCUS As DAO.Recordset
Dim rsVESFLE As DAO.Recordset

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
     sagot = MsgBox("CUSTOMER already exist in the list. ADD customer anyway?", vbQuestion + vbYesNo, Me.Caption)
     If sagot = vbYes Then
        Call addItem
     Else
        Call Clear_Fields
        Me.txtNumber.SetFocus
     End If
  End If
End Sub

Private Sub cmdClear_Click()
   Call Clear_Fields
   Me.txtNumber.SetFocus
End Sub

Private Sub cmdSave_Click()
   sagot = MsgBox("Are you sure you want to SAVE this transaction?", vbQuestion + vbYesNo, Me.Caption)
   If sagot = vbYes Then
   
      Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & pubYEAR & "' and vesfle.month = '" & pubMONTH & "' and vesfle.code = 'O' and vesfle.type = '" & cType & "'")

      ' delete old records to replace by new/updated records
      Do While Not rsVESFLE.EOF
        If rsVESFLE.Fields!Year = pubYEAR And _
           rsVESFLE.Fields!Month = pubMONTH And _
           rsVESFLE.Fields!Code = "O" And _
           rsVESFLE.Fields!Type = cType Then
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
          
          rsVESFLE.Fields!Code = "O"
          rsVESFLE.Fields!Type = cType                                  'type
          
          rsVESFLE.Fields!Customer = Me.lv2.ListItems.Item(i).Text      'customer
          rsVESFLE.Fields!Name = Me.lv2.ListItems.Item(i).SubItems(1)   'customer name
          rsVESFLE.Fields!NetAmt = Me.lv2.ListItems.Item(i).SubItems(2) 'net amount
       
          ' Compute VAT amount (10% VAT is applied)
          vatAmtTemp = 0
          vatAmtTemp = Me.lv2.ListItems.Item(i).SubItems(2) / 10
          rsVESFLE.Fields!Vatamt = vatAmtTemp
       
          ' Compute Gross Amount
          grsAmtTemp = 0
          grsAmtTemp = Me.lv2.ListItems.Item(i).SubItems(2) + vatAmtTemp
          rsVESFLE.Fields!Grsamt = grsAmtTemp
       
          ' Update VAT Input File
          rsVESFLE.Update
       Next
       ' Save data then Enable System Menu
       ' and Out to VAT input data entry
       Call Enable_System_Menu
       Unload Me
   End If
End Sub

Private Sub txtNumber_Change()
   On Error Resume Next
   test
   If Me.txtNumber.Text <> "" Then
      Set rsVESCUS = db.OpenRecordset("SELECT name FROM VESCUS WHERE Customer = '" & txtNumber.Text & "'")
      If rsVESCUS.RecordCount <> 0 Then
         Me.txtCustomerName.Text = rsVESCUS.Fields!Name
         Me.txtNetAmount.SetFocus
      Else
         If Len(Me.txtNumber) = 5 Then
            sagot = MsgBox("INVALID CUSTOMER NUMBER. Please check your CUSTOMER list..", vbInformation, Me.Caption)
            Call Clear_Fields
         End If
      End If
   Else
      Call Clear_Fields
   End If
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   test
   If KeyAscii = 13 Then
      If Me.txtNumber.Text <> "" Then
         Set rsVESCUS = db.OpenRecordset("SELECT name FROM VESVEN WHERE Customer = '" & txtNumber.Text & "'")
         If rsVESCUS.RecordCount <> 0 Then
            Me.txtCustomerName.Text = rsVESVEN.Fields!Name
            Me.txtNetAmount.SetFocus
         Else
            Call Clear_Fields
         End If
      Else
         Call Clear_Fields
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   sagot = MsgBox("Are you sure you want to CANCEL this document entry?", vbQuestion + vbYesNo, Me.Caption)
   If sagot = vbYes Then
      Call Enable_System_Menu
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Call load_Month
   Select Case cType
     Case 1
        Me.txtLabel = "VAT OUTPUT FOR " + iMonth(pubMONTH) + " " + pubYEAR
        Me.Caption = " « Vat Output"
     Case 3
        Me.Caption = " « Vat Output (Exempt)"
        Me.txtLabel = "VAT OUTPUT (EXEMPT) FOR " + iMonth(pubMONTH) + " " + pubYEAR
   End Select

   Me.Top = 30
   Me.Left = 10
   
   ' Initialize Fields
   Call Clear_Fields
   ' Load Cstomer List
   Call loadLV
   ' Load Detail List or Previous Entry
   Call loadLV2
   
   Me.lv1.Enabled = False
   Me.cmdEnable.Caption = "&Enable List"
   Me.cmdRefresh.Enabled = False
   
   Me.WindowState = 0
   Call test
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = "{tab}" Then
   End If
End Sub

Private Sub loadLV()
   '' Load Members ' Once
   Set rsVESCUS = db.OpenRecordset("SELECT vescus.name, vescus.customer FROM VESCUS order by name")
   Me.lv1.ListItems.Clear

   rsVESCUS.MoveFirst
   Do While Not rsVESCUS.EOF
      Set x = Me.lv1.ListItems.Add(, , rsVESCUS.Fields!Name)
      x.SubItems(1) = rsVESCUS.Fields!Customer
      rsVESCUS.MoveNext
   Loop
End Sub

Private Sub lv1_DblClick()
   rw = Me.lv1.SelectedItem.Index
   Me.txtNumber.Text = Me.lv1.ListItems.Item(rw).SubItems(1)
   Me.lv1.Enabled = False
   Me.cmdRefresh.Enabled = False
   Me.cmdEnable.Caption = "&Enable List"
   Me.txtNetAmount.SetFocus
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13
          rw = Me.lv1.SelectedItem.Index
          Me.txtNumber.Text = Me.lv1.ListItems.Item(rw).SubItems(1)
          Me.lv1.Enabled = False
          Me.cmdRefresh.Enabled = False
          Me.cmdEnable.Caption = "&Enable List"
          Me.txtNetAmount.SetFocus
   End Select
End Sub

Private Sub loadLV2()
   Dim valTemp As Double
   
   ' Load Previous Entry. If there is..
   Set rsVESFLE = db.OpenRecordset("SELECT vesfle.customer, vescus.name, vesfle.netamt FROM VESFLE INNER JOIN VESCUS ON vesfle.customer = vescus.customer WHERE vesfle.year = '" & pubYEAR & "' and vesfle.month = '" & pubMONTH & "' and vesfle.code = 'O' and vesfle.type = '" & cType & "'")
   Me.lv2.ListItems.Clear
   
   valTemp = 0
   Do While Not rsVESFLE.EOF
      Set x = Me.lv2.ListItems.Add(, , rsVESFLE.Fields!Customer)
      x.SubItems(1) = rsVESFLE.Fields!Name
      x.SubItems(2) = Format(Val(rsVESFLE.Fields!NetAmt), "###,###,##0.00")
      valTemp = valTemp + rsVESFLE.Fields!NetAmt
      rsVESFLE.MoveNext
   Loop
   
   Me.txtTotal.Text = Format(valTemp, "###,###,##0.00")
End Sub

Private Sub lv2_DblClick()
   Dim valTemp As Double
   
   rw = Me.lv2.SelectedItem.Index
   Me.txtNumber.Text = Me.lv2.ListItems.Item(rw).Text
   Me.txtCustomerName.Text = Me.lv2.ListItems.Item(rw).SubItems(1)
   Me.txtNetAmount.Text = Format$(Me.lv2.ListItems.Item(rw).SubItems(2), "#.00")
   
   Me.lv2.ListItems.Remove (rw)
   valTemp = 0
   For i = 1 To Me.lv2.ListItems.Count
       valTemp = valTemp + Me.lv2.ListItems.Item(i).SubItems(2)
   Next
   Me.txtTotal.Text = Format(valTemp, "###,###,##0.00")
   Me.cmdAddItem.Enabled = True
   Me.txtNumber.SetFocus
End Sub

Private Sub lv2_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDelete
       If Me.lv2.ListItems.Count <> 0 Then
          rw = Me.lv2.SelectedItem.Index
          Me.lv2.ListItems.Remove (rw)
          valTemp = 0
          For i = 1 To Me.lv2.ListItems.Count
              valTemp = valTemp + Me.lv2.ListItems.Item(i).SubItems(2)
          Next
          Me.txtTotal.Text = Format(valTemp, "###,###,##0.00")
          Me.lv2.SetFocus
       End If
    End Select
End Sub

Private Sub cmdEnable_Click()
   ' Enable / Disable Button
   If Me.lv1.Enabled Then
      Me.lv1.Enabled = False
      Me.cmdRefresh.Enabled = False
      Me.cmdEnable.Caption = "&Enable List"
      Me.txtNumber.SetFocus
   Else
      Me.lv1.Enabled = True
      Me.cmdRefresh.Enabled = True
      Me.cmdEnable.Caption = "&Disable List"
      Me.lv1.SetFocus
   End If
End Sub

Private Sub cmdRefresh_Click()
   'Refresh Customer List
   Call loadLV
   Me.lv1.SetFocus
End Sub

Private Sub Clear_Fields()
   Me.txtNumber.Text = ""
   Me.txtCustomerName.Text = ""
   Me.txtNetAmount.Text = ""
End Sub

Private Sub addItem()
   Dim valTemp As Double
   
   Set j = Me.lv2.ListItems.Add(, , Me.txtNumber.Text)
   j.SubItems(1) = Me.txtCustomerName.Text
   j.SubItems(2) = Format(Me.txtNetAmount.Text, "###,###,##0.00")
     
   Call Clear_Fields
        
   Me.txtNumber.SetFocus
   valTemp = 0
   For i = 1 To Me.lv2.ListItems.Count
       valTemp = valTemp + Me.lv2.ListItems.Item(i).SubItems(2)
   Next
   Me.txtTotal.Text = Format$(valTemp, "###,###,##0.00")
End Sub

Private Sub txtNetAmount_Change()
   On Error Resume Next
   test
   If Len(Trim(Me.txtNetAmount)) = 11 Then
      Me.cmdAddItem.SetFocus
   End If
End Sub

Private Sub txtNetAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Me.txtNumber.Text <> "" And Me.txtNetAmount.Text <> "" Then
          Me.cmdAddItem.SetFocus
       End If
    End If
End Sub

Private Sub test()
   If Me.txtNumber.Text <> "" And Me.txtNetAmount.Text <> "" Then
      Me.cmdAddItem.Enabled = True
   Else
      Me.cmdAddItem.Enabled = False
   End If
End Sub

Private Sub ValidateData(AllOK As Boolean)
   Dim Message As String
   AllOK = True
   Message = ""
   If Len(Me.txtNumber.Text) = 0 Then
      Message = "You must enter a Customer number." + vbCrLf
      Me.txtNumber.SetFocus
      AllOK = False
   ElseIf Len(Me.txtNumber.Text) < 5 Then
      Message = "Customer number must be 5 characters long." + vbCrLf
      Me.txtNumber.SetFocus
      AllOK = False
   End If
   If Not (AllOK) Then
      MsgBox Message, vbOKOnly + vbInformation, "Validation Error"
   End If
End Sub


