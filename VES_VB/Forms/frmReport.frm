VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Report"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   10830
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   6360
      Width           =   10575
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   2160
      TabIndex        =   9
      Top             =   5640
      Width           =   8535
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
         Left            =   1080
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtVatAmount 
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
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtGrossAmount 
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
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl3 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl2 
         Caption         =   "V.A.T."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   3120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "Gross"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClose 
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
      Height          =   975
      Left            =   120
      Picture         =   "frmReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Picture         =   "frmReport.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton cmdFlexToExcel 
      Caption         =   "Copy to &Excel"
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
      Picture         =   "frmReport.frx":2BE4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate List"
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
      Picture         =   "frmReport.frx":3026
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   8535
      Begin MSComctlLib.ListView lv1 
         Height          =   4575
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8070
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblHead 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reqd parameters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReport.frx":3CF0
         Left            =   720
         List            =   "frmReport.frx":3D8D
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReport.frx":3EC3
         Left            =   720
         List            =   "frmReport.frx":3EEB
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Month"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmReport"
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
'#           area :  frmReport                #
'#    description :  Code File Listing        #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim varCompany As String
Dim varBranch As String
Dim varHeader As String

Dim rsVESREP As DAO.Recordset
Dim rsSETFLE As DAO.Recordset
Dim rsVESCOD As DAO.Recordset
Dim rsVESFLE As DAO.Recordset

Private Sub cmdClose_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdFlexToExcel_Click()
   MousePointer = vbHourglass
   Call CopyListDataToExcel(lv1)
   MousePointer = vbDefault
End Sub

Private Sub cmdGenerate_Click()
   If cboYear.Text <> "" And cboMonth.Text <> "" Then
      Select Case rptPointer
      Case 1   'VAT INPUT
         Set rsVESREP = db.OpenRecordset("SELECT * FROM VESREP")
         Do While Not rsVESREP.EOF
            rsVESREP.Delete
            rsVESREP.MoveNext
         Loop
         Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & cboYear & "' and vesfle.month = '" & cboMonth & "' and vesfle.code = 'I'")
         varHeader = "VAT INPUT FOR " + iMonth(cboMonth) + " " + cboYear
      Case 2   'VAT OUTPUT
         Set rsVESREP = db.OpenRecordset("SELECT * FROM VESREP")
         Do While Not rsVESREP.EOF
            rsVESREP.Delete
            rsVESREP.MoveNext
         Loop
         Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & cboYear & "' and vesfle.month = '" & cboMonth & "' and vesfle.code = 'O' and vesfle.type = '1'")
         varHeader = "VAT OUTPUT FOR " + iMonth(cboMonth) + " " + cboYear
      Case 3   'VAT OUTPUT (EXEMPT)
         Set rsVESREP = db.OpenRecordset("SELECT * FROM VESREP")
         Do While Not rsVESREP.EOF
            rsVESREP.Delete
            rsVESREP.MoveNext
         Loop
         Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & cboYear & "' and vesfle.month = '" & cboMonth & "' and vesfle.code = 'O' and vesfle.type = '3'")
         varHeader = "VAT OUTPUT (EXEMPT) FOR " + iMonth(cboMonth) + " " + cboYear
      Case 4   'WITHOLDING TAX
         Set rsVESREP = db.OpenRecordset("SELECT * FROM VESREP")
         Do While Not rsVESREP.EOF
            rsVESREP.Delete
            rsVESREP.MoveNext
         Loop
         Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & cboYear & "' and vesfle.month = '" & cboMonth & "' and vesfle.code = 'T'")
         Set rsVESCOD = db.OpenRecordset("SELECT * FROM VESCOD")
         varHeader = "WITHOLDING TAX FOR " + iMonth(cboMonth) + " " + cboYear
      End Select
      Call do_ListView_Header
      Call loadLV
   Else
      MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
      cboYear.SetFocus
   End If
End Sub

Private Sub do_ListView_Header()
   Select Case rptPointer
   Case 1   'VAT INPUT
     Call clear_LV_headers
     Me.lv1.ColumnHeaders.Add(1, , , 1440) = "Number"
     Me.lv1.ColumnHeaders.Add(2, , , 4000) = "Vendor Name"
     Me.lv1.ColumnHeaders.Add(3, , , 1000) = "T.I.N."
     Me.lv1.ColumnHeaders.Add(4, , , 700) = "Type"
     Me.lv1.ColumnHeaders.Add(5, , , 1500, 1) = "Amount"
     Me.lv1.ColumnHeaders.Add(6, , , 1500, 1) = "V.A.T."
     Me.lv1.ColumnHeaders.Add(7, , , 1500, 1) = "Gross"
   Case 2   'VAT OUTPUT
     Call clear_LV_headers
     Me.lv1.ColumnHeaders.Add(1, , , 1440) = "Number"
     Me.lv1.ColumnHeaders.Add(2, , , 4000) = "Customer Name"
     Me.lv1.ColumnHeaders.Add(3, , , 1500, 1) = "Amount"
     Me.lv1.ColumnHeaders.Add(4, , , 1500, 1) = "V.A.T."
     Me.lv1.ColumnHeaders.Add(5, , , 1500, 1) = "Gross"
   Case 3   'VAT OUTPUT (EXEMPT)
     Call clear_LV_headers
     Me.lv1.ColumnHeaders.Add(1, , , 1440) = "Number"
     Me.lv1.ColumnHeaders.Add(2, , , 4000) = "Customer Name"
     Me.lv1.ColumnHeaders.Add(3, , , 1500, 1) = "Amount"
     Me.lv1.ColumnHeaders.Add(4, , , 1500, 1) = "V.A.T."
     Me.lv1.ColumnHeaders.Add(5, , , 1500, 1) = "Gross"
   Case 4   'WITHOLDING TAX
     Call clear_LV_headers
     Me.lv1.ColumnHeaders.Add(1, , , 1440) = "Number"
     Me.lv1.ColumnHeaders.Add(2, , , 4000) = "Vendor Name"
     Me.lv1.ColumnHeaders.Add(3, , , 1000) = "TIN"
     Me.lv1.ColumnHeaders.Add(4, , , 700) = "TC"
     Me.lv1.ColumnHeaders.Add(5, , , 900) = "Rate"
     Me.lv1.ColumnHeaders.Add(6, , , 1500, 1) = "TAX Base"
     Me.lv1.ColumnHeaders.Add(7, , , 1500, 1) = "Tax"
     Me.lv1.ColumnHeaders.Add(8, , , 1500, 1) = "Net"
     Me.lv1.ColumnHeaders.Add(9, , , 1440) = "Debit"
   End Select
End Sub

Private Sub display_Caption()
   Select Case rptPointer
   Case 1    'VAT INPUT
     Me.Caption = " « VAT Input Editlist"
     Me.lblHead = "Value Added Tax Input"
   Case 2    'VAT OUTPUT
     Me.Caption = " « VAT Output Editlist"
     Me.lblHead = "Value Added Tax Output"
   Case 3    'VAT OUTPUT (EXEMPT)
     Me.Caption = " « VAT Output (Exempt) Editlist"
     Me.lblHead = "Value Added Tax Output (Exempt)"
   Case 4    'WITHOLDING TAX
     Me.Caption = " « Witholding TAX Editlist"
     Me.lblHead = "Witholding Tax"
     Me.lbl1.Caption = "Base"
     Me.lbl2.Caption = "Tax"
     Me.lbl3.Caption = "Net Amt"
   End Select
End Sub

Private Sub loadLV()
   Dim varGrossAmount As Double
   Dim varVatAmount As Double
   Dim varNetAmount As Double
   Dim varRATE As Double
   Dim varTAX As Double
   Dim varNET As Double
   
   'Initialize variables
   varGrossAmount = 0
   varVatAmount = 0
   varNetAmount = 0
   varRATE = 0
   varTAX = 0
   varNET = 0

   Me.lv1.Gridlines = True
   Select Case rptPointer
   Case 1      'VAT INPUT
      Me.lv1.ListItems.Clear
      Do While Not rsVESFLE.EOF
         Set x = Me.lv1.ListItems.Add(, , rsVESFLE.Fields!Vendor)
         x.SubItems(1) = rsVESFLE.Fields!Name
         x.SubItems(2) = rsVESFLE.Fields!Tin
         x.SubItems(3) = rsVESFLE.Fields!Type
         x.SubItems(4) = Format(Val(rsVESFLE.Fields!NetAmt), "###,###,##0.00")
         x.SubItems(5) = Format(Val(rsVESFLE.Fields!Vatamt), "###,###,##0.00")
         x.SubItems(6) = Format(Val(rsVESFLE.Fields!Grsamt), "###,###,##0.00")
         
         varGrossAmount = varGrossAmount + Format(Val(rsVESFLE.Fields!Grsamt), "########0.00")
         varVatAmount = varVatAmount + Format(Val(rsVESFLE.Fields!Vatamt), "########0.00")
         varNetAmount = varNetAmount + Format(Val(rsVESFLE.Fields!NetAmt), "########0.00")
         
         'ADD RECORD TO REPORT FILE
         rsVESREP.AddNew
         rsVESREP.Fields!Number = rsVESFLE.Fields!Vendor
         rsVESREP.Fields!Name = rsVESFLE.Fields!Name
         rsVESREP.Fields!Tin = rsVESFLE.Fields!Tin
         rsVESREP.Fields!Type = rsVESFLE.Fields!Type
         rsVESREP.Fields!Amount = rsVESFLE.Fields!NetAmt
         rsVESREP.Fields!Vat = rsVESFLE.Fields!Vatamt
         rsVESREP.Fields!Gross = rsVESFLE.Fields!Grsamt
         rsVESREP.Fields!Header = varHeader
         rsVESREP.Fields!Branch = varBranch
         rsVESREP.Update
         'ADD - END
         
         rsVESFLE.MoveNext
      Loop
      
   Case 2      'VAT OUTPUT
      Me.lv1.ListItems.Clear
      Do While Not rsVESFLE.EOF
         Set x = Me.lv1.ListItems.Add(, , rsVESFLE.Fields!Customer)
         x.SubItems(1) = rsVESFLE.Fields!Name
         x.SubItems(2) = Format(Val(rsVESFLE.Fields!NetAmt), "###,###,##0.00")
         x.SubItems(3) = Format(Val(rsVESFLE.Fields!Vatamt), "###,###,##0.00")
         x.SubItems(4) = Format(Val(rsVESFLE.Fields!Grsamt), "###,###,##0.00")
         
         varGrossAmount = varGrossAmount + Format(Val(rsVESFLE.Fields!Grsamt), "########0.00")
         varVatAmount = varVatAmount + Format(Val(rsVESFLE.Fields!Vatamt), "########0.00")
         varNetAmount = varNetAmount + Format(Val(rsVESFLE.Fields!NetAmt), "########0.00")
         
         'ADD RECORD TO REPORT FILE
         rsVESREP.AddNew
         rsVESREP.Fields!Number = rsVESFLE.Fields!Customer
         rsVESREP.Fields!Name = rsVESFLE.Fields!Name
         rsVESREP.Fields!Amount = rsVESFLE.Fields!NetAmt
         rsVESREP.Fields!Vat = rsVESFLE.Fields!Vatamt
         rsVESREP.Fields!Gross = rsVESFLE.Fields!Grsamt
         rsVESREP.Fields!Header = varHeader
         rsVESREP.Fields!Branch = varBranch
         rsVESREP.Update
         'ADD - END
         
         rsVESFLE.MoveNext
      Loop
      
   Case 3      'VAT OUTPUT (EXEMPT)
      Me.lv1.ListItems.Clear
      Do While Not rsVESFLE.EOF
         Set x = Me.lv1.ListItems.Add(, , rsVESFLE.Fields!Customer)
         x.SubItems(1) = rsVESFLE.Fields!Name
         x.SubItems(2) = Format(Val(rsVESFLE.Fields!NetAmt), "###,###,##0.00")
         x.SubItems(3) = Format(Val(rsVESFLE.Fields!Vatamt), "###,###,##0.00")
         x.SubItems(4) = Format(Val(rsVESFLE.Fields!Grsamt), "###,###,##0.00")
         
         varGrossAmount = varGrossAmount + Format(Val(rsVESFLE.Fields!Grsamt), "########0.00")
         varVatAmount = varVatAmount + Format(Val(rsVESFLE.Fields!Vatamt), "########0.00")
         varNetAmount = varNetAmount + Format(Val(rsVESFLE.Fields!NetAmt), "########0.00")
         
         'ADD RECORD TO REPORT FILE
         rsVESREP.AddNew
         rsVESREP.Fields!Number = rsVESFLE.Fields!Customer
         rsVESREP.Fields!Name = rsVESFLE.Fields!Name
         rsVESREP.Fields!Amount = rsVESFLE.Fields!NetAmt
         rsVESREP.Fields!Vat = rsVESFLE.Fields!Vatamt
         rsVESREP.Fields!Gross = rsVESFLE.Fields!Grsamt
         rsVESREP.Fields!Header = varHeader
         rsVESREP.Fields!Branch = varBranch
         rsVESREP.Update
         'ADD - END
         
         rsVESFLE.MoveNext
      Loop
      
   Case 4      'WITHOLDING TAX
      Me.lv1.ListItems.Clear
      Do While Not rsVESFLE.EOF
         Set x = Me.lv1.ListItems.Add(, , rsVESFLE.Fields!Vendor)
         
         rsVESCOD.FindFirst "Code = '" & Mid(Trim(rsVESFLE.Fields!Customer), 1, 2) & "'"
         If Not rsVESCOD.NoMatch Then
            varRATE = rsVESCOD.Fields!Rate
         End If
                  
         x.SubItems(1) = rsVESFLE.Fields!Name
         x.SubItems(2) = rsVESFLE.Fields!Tin
         x.SubItems(3) = Mid(Trim(rsVESFLE.Fields!Customer), 1, 2)
         x.SubItems(4) = Format(varRATE * 100, "#0.00") + "%"
         
         x.SubItems(5) = Format(Val(rsVESFLE.Fields!Grsamt), "###,###,##0.00")
         varGrossAmount = varGrossAmount + Format(Val(rsVESFLE.Fields!Grsamt), "########0.00")
         
         'COMPUTE TAX
         varTAX = rsVESFLE.Fields!Grsamt * varRATE
         x.SubItems(6) = Format(Val(varTAX), "###,###,##0.00")
         varVatAmount = varVatAmount + Format(Val(varTAX), "########0.00")
         
         'COMPUTE NET
         varNET = rsVESFLE.Fields!Grsamt - varTAX
         x.SubItems(7) = Format(Val(varNET), "###,###,##0.00")
         varNetAmount = varNetAmount + Format(Val(varNET), "########0.00")
         
         x.SubItems(8) = rsVESFLE.Fields!Acctno
         
         'ADD RECORD TO REPORT FILE
         rsVESREP.AddNew
         rsVESREP.Fields!Number = rsVESFLE.Fields!Vendor
         rsVESREP.Fields!Name = rsVESFLE.Fields!Name
         rsVESREP.Fields!Tin = rsVESFLE.Fields!Tin
         rsVESREP.Fields!Tc = Mid(Trim(rsVESFLE.Fields!Customer), 1, 2)
         rsVESREP.Fields!Rate = varRATE
         rsVESREP.Fields!Gross = rsVESFLE.Fields!Grsamt
         rsVESREP.Fields!Vat = varTAX
         rsVESREP.Fields!Amount = varNET
         rsVESREP.Fields!Debit = rsVESFLE.Fields!Acctno
         rsVESREP.Fields!Header = varHeader
         rsVESREP.Fields!Branch = varBranch
         rsVESREP.Update
         'ADD - END
         
         rsVESFLE.MoveNext
      Loop
   End Select
   
   Me.txtGrossAmount = Format(varGrossAmount, "###,###,##0.00")
   Me.txtVatAmount = Format(varVatAmount, "###,###,##0.00")
   Me.txtNetAmount = Format(varNetAmount, "###,###,##0.00")
   
End Sub

Private Sub CopyListDataToExcel(LV As ListView)
   On Error GoTo errHandler
   Dim EXCELApp As Excel.Application
   Dim EXCELWorkBook As Excel.Workbook
   
   If LV.ListItems.Count = 0 Then
      MsgBox "No Data to extract.", vbInformation, App.Title
      Exit Sub
   End If
    
   Set EXCELApp = CreateObject("Excel.application")
   Set EXCELWorkBook = EXCELApp.Workbooks.Add
    
   Dim New_Column As Boolean
   
   varCounter = 0
   varSwitch = 0
   For Row = 1 To LV.ListItems.Count
       ' -------------------- Progress Bar ---------------------- '
       varCounter = varCounter + 1
       Me.pb.Value = Int((varCounter / LV.ListItems.Count) * 100)
       ' -------------------------------------------------------- '
       'COPY LIST HEADER TO EXCEL
       If varSwitch = 0 Then
          varSwitch = 1
          cl = 1
          For P = 1 To lv1.ColumnHeaders.Count
              EXCELApp.Cells(Row, cl).Value = UCase(LV.ColumnHeaders.Item(P).Text)
              EXCELApp.Cells(Row, cl).Font.Bold = True
              cl = cl + 1
          Next
       End If
       'COPY LIST DATA TO EXCEL
       Cols = 1
       EXCELApp.Cells(Row + 1, Cols).Value = LV.ListItems.Item(Row).Text
       Cols = Cols + 1
       For Col = 1 To LV.ListItems.Item(Row).ListSubItems.Count
           EXCELApp.Cells(Row + 1, Cols).Value = LV.ListItems.Item(Row).SubItems(Col)
           Cols = Cols + 1
       Next
   Next
   EXCELApp.Columns.AutoFit
   EXCELApp.Cells(LV.ListItems.Count + 3, 1).Value = varHeader
   EXCELApp.Cells(LV.ListItems.Count + 3, 1).Font.Bold = True
   EXCELApp.Application.Visible = True
   Set EXCELWorkBook = Nothing
   Set EXCELApp = Nothing
   LV.SetFocus
   MsgBox "Data has been successfully copied to Excel. ", vbInformation, "Success"
   Me.pb.Value = 0
   Exit Sub

errHandler:
   MousePointer = vbDefault
   MsgBox err.Description, vbInformation, "Info "
End Sub

Private Sub clear_LV_headers()
   Me.lv1.ColumnHeaders.Clear
End Sub

Private Sub cmdPrint_Click()

  If lv1.ListItems.Count = 0 Then
     MsgBox "No Data to print.", vbInformation, App.Title
     Exit Sub
  End If

  Me.MousePointer = vbHourglass
  Select Case rptPointer
  Case 1
     With Me.CrystalReport1
       Me.CrystalReport1.DataFiles(0) = App.Path & "\VES.MDB"
       Me.CrystalReport1.ReportFileName = App.Path & "\Reports\vatinput.rpt"
       Me.CrystalReport1.Action = 1
       Me.CrystalReport1.PageZoom (100)
     End With
  Case 2
     With Me.CrystalReport1
       Me.CrystalReport1.DataFiles(0) = App.Path & "\VES.MDB"
       Me.CrystalReport1.ReportFileName = App.Path & "\Reports\vatotput.rpt"
       Me.CrystalReport1.Action = 1
       Me.CrystalReport1.PageZoom (100)
     End With
  Case 3
     With Me.CrystalReport1
       Me.CrystalReport1.DataFiles(0) = App.Path & "\VES.MDB"
       Me.CrystalReport1.ReportFileName = App.Path & "\Reports\vatotput.rpt"
       Me.CrystalReport1.Action = 1
       Me.CrystalReport1.PageZoom (100)
     End With
  Case 4
     With Me.CrystalReport1
       Me.CrystalReport1.DataFiles(0) = App.Path & "\VES.MDB"
       Me.CrystalReport1.ReportFileName = App.Path & "\Reports\withtax.rpt"
       Me.CrystalReport1.Action = 1
       Me.CrystalReport1.PageZoom (100)
     End With
  End Select
  Me.MousePointer = vbDefault
End Sub

Private Sub cboYear_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Private Sub cboMonth_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Private Sub Form_Load()
   Call Disable_System_Menu
   Call load_Month
   Call display_Caption
   Set rsSETFLE = db.OpenRecordset("SELECT * FROM SETFLE")
   If rsSETFLE.RecordCount <> 0 Then
      Me.cboMonth = rsSETFLE.Fields!Month
      Me.cboYear = rsSETFLE.Fields!Year
      varCompany = rsSETFLE.Fields!Coname
      varBranch = rsSETFLE.Fields!Brname + " - " + rsSETFLE.Fields!Branch
   End If
   Me.Top = 30
   Me.Left = 10
End Sub

