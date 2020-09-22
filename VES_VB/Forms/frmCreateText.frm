VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCreateText 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Create VAT Output / Input Text File"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7455
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   7215
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1575
      Left            =   4440
      TabIndex        =   8
      Top             =   1440
      Width           =   2895
      Begin VB.Frame Frame5 
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   2655
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
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   2175
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "Create &TEXT File"
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
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Enter Filename"
      Height          =   1335
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "&Generate Name"
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
         Left            =   600
         TabIndex        =   9
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtFileName 
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Enter Destination"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4215
      Begin VB.CommandButton cmdDirectory 
         Caption         =   "Select &Destination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtDestination 
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
         Left            =   240
         TabIndex        =   5
         Text            =   "A:\"
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Year and Quarter"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
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
         ItemData        =   "frmCreateText.frx":0000
         Left            =   1200
         List            =   "frmCreateText.frx":009D
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save This"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         Picture         =   "frmCreateText.frx":01D3
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cboQuarter 
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
         ItemData        =   "frmCreateText.frx":0615
         Left            =   1200
         List            =   "frmCreateText.frx":0625
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblSave 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   17
         Top             =   840
         Width           =   255
      End
      Begin VB.Label label2 
         Caption         =   "Quarter"
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   735
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCreateText"
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
'#           area :  frmCreateText            #
'#    description :  Code File VES            #
'#                :  Text File Creator        #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim varBranch As String * 3
Dim varYear As String * 4
Dim varMonth As String * 2
Dim varVendor As String * 7
Dim varCustomer As String * 5
Dim varName As String * 30
Dim varNetamt As String * 15
Dim varAddress1 As String * 30
Dim varAddress2 As String * 30
Dim varTin As String * 9
Dim varType As String * 1
Dim varAcctno As String * 7
   
Dim cFilter1
Dim cFilter2
Dim cFilter3

Dim rsVESVEN As DAO.Recordset
Dim rsVESCUS As DAO.Recordset
Dim rsVESFLE As DAO.Recordset
Dim rsSETFLE As DAO.Recordset

Private Sub cboQuarter_Click()
   Me.lblSave.Caption = ""
End Sub

Private Sub cmdCancel_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdCreate_Click()
   Dim tempYear
   tempYear = Year(Date)
   If Me.cboYear > tempYear Then
      MsgBox "Invalid Year. Try again...", vbInformation, Me.Caption
      Me.cboYear.SetFocus
      Exit Sub
   End If
   If Me.txtDestination.Text = "" Then
      MsgBox "You must enter a Destination or Path.", vbInformation, Me.Caption
      Me.txtDestination.SetFocus
      Exit Sub
   ElseIf Me.txtFileName.Text = "" Then
      MsgBox "You must enter a FileName.", vbInformation, Me.Caption
      Me.txtFileName.SetFocus
      Exit Sub
   End If
   Me.Frame6.Caption = ""
   Me.pb.Value = 0
   Call Vat_Input_Output_Tax_Disk_Create
End Sub

Private Sub cmdDirectory_Click()
   Dim strTemp As String
   strTemp = fBrowseForFolder(Me.hWnd, "Select destination / path")
   If strTemp <> "" Then
      txtDestination = strTemp
   End If
End Sub

Private Sub cmdGenerate_Click()
   Dim tempYear
   tempYear = Year(Date)
   
   If Me.cboYear > tempYear Then
      MsgBox "Invalid Year. Try again...", vbInformation, Me.Caption
      Me.cboYear.SetFocus
      Exit Sub
   End If
   If Me.cboYear <> "" And Me.cboQuarter <> "" Then
      Me.txtFileName.Text = pubBRANCH + cboYear + Me.cboQuarter.Text + ".V" + cCode + "F"
   Else
      MsgBox "You must enter a Year and Quarter.", vbInformation, Me.Caption
      If Me.cboYear <> "" Then
         Me.cboQuarter.SetFocus
      Else
         Me.cboYear.SetFocus
      End If
   End If
End Sub

Private Sub cmdSave_Click()
   Dim tempYear
   tempYear = Year(Date)
   
   If Me.cboYear <> "" And Me.cboQuarter <> "" Then
      If Me.cboYear > tempYear Then
         sagot = MsgBox("Invalid Year. Check your entry...", vbInformation, Me.Caption)
         Me.cboYear.SetFocus
         Exit Sub
      End If
      Set rsSETFLE = db.OpenRecordset("SELECT * FROM SETFLE")
      rsSETFLE.Edit
      rsSETFLE.Fields!Year = Me.cboYear
      rsSETFLE.Fields!Quarter = Me.cboQuarter.Text
      rsSETFLE.Update
      Me.lblSave.Caption = ":-)"
   Else
      MsgBox "You must enter a Year & Quarter.", vbInformation, Me.Caption
      If Me.cboYear <> "" Then
         Me.cboQuarter.SetFocus
      Else
         Me.cboYear.SetFocus
      End If
   End If
End Sub

Private Sub Form_Load()
   Call Disable_System_Menu
   Set rsSETFLE = db.OpenRecordset("SELECT * FROM SETFLE")
   Me.cboYear = rsSETFLE.Fields!Year
   Me.cboQuarter = rsSETFLE.Fields!Quarter
   
   If cCode = "I" Then
      Me.Caption = " « Create VAT Input Diskette"
   ElseIf cCode = "O" Then
      Me.Caption = " « Create VAT Output Diskette"
   ElseIf cCode = "T" Then
      Me.Caption = " « Create TAX Diskette"
   End If
   
   Me.Frame6.Caption = ""
   Me.pb.Value = 0
   
   Me.Top = 1700
   Me.Left = 2000
   Me.WindowState = 0
End Sub

Private Sub Vat_Input_Output_Tax_Disk_Create()

   On Error GoTo Create_Error
   
   If Me.cboQuarter = 1 Then
      cFilter1 = "01"
      cFilter2 = "02"
      cFilter3 = "03"
   ElseIf Me.cboQuarter = 2 Then
      cFilter1 = "04"
      cFilter2 = "05"
      cFilter3 = "06"
   ElseIf Me.cboQuarter = 3 Then
      cFilter1 = "07"
      cFilter2 = "08"
      cFilter3 = "09"
   ElseIf Me.cboQuarter = 4 Then
      cFilter1 = "10"
      cFilter2 = "11"
      cFilter3 = "12"
   End If
   
   ' Collect data based on year, quarter and code
   Set rsVESFLE = db.OpenRecordset("SELECT * FROM VESFLE WHERE vesfle.year = '" & Me.cboYear & "' and vesfle.code = '" & cCode & "' and vesfle.month = '" & cFilter1 & "' or vesfle.month = '" & cFilter2 & "' or vesfle.month = '" & cFilter3 & "'")
   If rsVESFLE.RecordCount = 0 Then
       sagot = MsgBox("RECORDSET is empty. There are no transactions for extraction...", vbInformation, Me.Caption)
       Exit Sub
   End If
   
   ' ----------------------------------- '
   ' Create Text File                    '
   ' ----------------------------------- '
   ' First, check if textfile already exist
   If FileExists(Me.txtDestination & "\" & Me.txtFileName) = True Then
      sagot = MsgBox("File " & Me.txtFileName & " already exist. Overwrite?", vbQuestion + vbYesNo, Me.Caption)
      If sagot = vbNo Then
         Exit Sub
      End If
   End If
   
   Set Obj = CreateObject("Scripting.FileSystemObject")
   Set F1 = Obj.CreateTextFile(Me.txtDestination & "\" & Me.txtFileName)
 
   varCounter = 0
      
   rsVESFLE.MoveFirst
   Do While Not rsVESFLE.EOF
      ' -------------------- Progress Bar ---------------------- '
      varCounter = varCounter + 1
      Me.pb.Value = Int((varCounter / rsVESFLE.RecordCount) * 100)
      ' -------------------------------------------------------- '
   
      varBranch = pubBRANCH
      varYear = Me.cboYear.Text
      varMonth = rsVESFLE.Fields!Month
      
      If rsVESFLE.Fields!Vendor <> "" Then
         varVendor = rsVESFLE.Fields!Vendor
      End If
      If rsVESFLE.Fields!Customer <> "" Then
         varCustomer = rsVESFLE.Fields!Customer
      End If
      If rsVESFLE.Fields!Tin <> "" Then
         varTin = rsVESFLE.Fields!Tin
      End If
      If rsVESFLE.Fields!Type <> "" Then
         varType = rsVESFLE.Fields!Type
      End If
      If rsVESFLE.Fields!Acctno <> "" Then
         varAcctno = rsVESFLE.Fields!Acctno
      End If
      
      If cCode = "I" Or cCode = "T" Then
         Set rsVESVEN = db.OpenRecordset("SELECT address1, address2, name FROM VESVEN WHERE Vendor = '" & rsVESFLE.Fields!Vendor & "'")
         If rsVESVEN.RecordCount <> 0 Then
            If rsVESVEN.Fields!Address1 <> "" Then
               varAddress1 = rsVESVEN.Fields!Address1
            End If
            If rsVESVEN.Fields!Address2 <> "" Then
               varAddress2 = rsVESVEN.Fields!Address2
            End If
            If rsVESVEN.Fields!Name <> "" Then
               varName = rsVESVEN.Fields!Name
            End If
         End If
      ElseIf cCode = "O" Then
         Set rsVESCUS = db.OpenRecordset("SELECT address1, address2, name FROM VESCUS WHERE Customer = '" & rsVESFLE.Fields!Customer & "'")
         If rsVESCUS.RecordCount <> 0 Then
            If rsVESCUS.Fields!Address1 <> "" Then
               varAddress1 = rsVESCUS.Fields!Address1
            End If
            If rsVESCUS.Fields!Address2 <> "" Then
               varAddress2 = rsVESCUS.Fields!Address2
            End If
            If rsVESCUS.Fields!Name <> "" Then
               varName = rsVESCUS.Fields!Name
            End If
         End If
      End If
      
      If cCode = "I" Or cCode = "O" Then
         RSet varNetamt = Format(rsVESFLE.Fields!NetAmt, "#0.00")
         ' Write
         F1.writeline varBranch + varYear + varMonth + cCode + varVendor + varCustomer + varName + varNetamt + varAddress1 + varAddress2 + varTin + varType
      ElseIf cCode = "T" Then
         RSet varNetamt = Format(rsVESFLE.Fields!Grsamt, "#0.00")
         ' Write
         F1.writeline varBranch + varYear + varMonth + cCode + varVendor + varCustomer + varName + varNetamt + varAddress1 + varAddress2 + varTin + varAcctno
      End If
      rsVESFLE.MoveNext
   Loop
   F1.Close
   Me.Frame6.Caption = Me.txtFileName + " created successfully..."
   Exit Sub
   
Create_Error:
   MsgBox err.Description
   Exit Sub
End Sub

Private Function FileExists(FullFileName As String) As Boolean
   On Error GoTo MakeF
   'If file does Not exist, there will be an Error
   Open FullFileName For Input As #1
   Close #1
   'no error, file exists
   FileExists = True
   Exit Function
MakeF:
   'error, file does Not exist
   FileExists = False
   Exit Function
End Function



