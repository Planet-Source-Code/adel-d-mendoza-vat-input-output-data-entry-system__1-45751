VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUpdateCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « Update Customer Masterdata"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   7845
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   7575
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extract Downloaded Customer Masterdata"
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
      Left            =   3600
      TabIndex        =   11
      Top             =   2520
      Width           =   3735
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   2778
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
         Text            =   "Filler"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Address 1"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address 2"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "VAT Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.FileListBox flb1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2820
      Left            =   120
      Pattern         =   "*.txt"
      TabIndex        =   9
      Top             =   120
      Width           =   3015
   End
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
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   4800
      Width           =   3135
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Customer Masterdata File"
      Enabled         =   0   'False
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
      Left            =   480
      TabIndex        =   7
      Top             =   4800
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   3240
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdDestination 
         Caption         =   "..."
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtDestination 
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
         Left            =   240
         TabIndex        =   2
         Text            =   "C:\"
         Top             =   720
         Width           =   3495
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   4215
         Begin VB.TextBox txtFileName 
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
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   3975
         End
         Begin VB.Label Label2 
            BackColor       =   &H00404040&
            Caption         =   " Source filename"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   " Source path"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmUpdateCustomer"
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
'#           area :  frmUpdateCustomer        #
'#    description :  Code File                #
'#                :  Update Customer Master   #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsVESCUS As DAO.Recordset

Private Sub cmdUpdate_Click()
   On Error GoTo ErrorHandler
   varCounter = 0
   Me.MousePointer = vbHourglass
   For i = 1 To Me.lv1.ListItems.Count
       ' --------------------- Progress Bar ----------------------- '
       varCounter = varCounter + 1
       Me.pb.Value = Int((varCounter / Me.lv1.ListItems.Count) * 100)
       ' ---------------------------------------------------------- '
       
       rsVESCUS.FindFirst "Customer = '" & Me.lv1.ListItems(i).SubItems(1) & "'"
       If Not rsVESCUS.NoMatch Then
          rsVESCUS.Edit
       Else
          rsVESCUS.AddNew
          rsVESCUS.Fields!Customer = Mid(Me.lv1.ListItems(i).SubItems(1), 1, 5)
       End If
       rsVESCUS.Fields!Name = Mid(Me.lv1.ListItems(i).SubItems(2), 1, 30)
       rsVESCUS.Fields!Address1 = Mid(Me.lv1.ListItems(i).SubItems(3), 1, 30)
       rsVESCUS.Fields!Address2 = Mid(Me.lv1.ListItems(i).SubItems(4), 1, 30)
       rsVESCUS.Fields!Vatno = Mid(Me.lv1.ListItems(i).SubItems(5), 1, 15)
       rsVESCUS.Update
   Next
   Me.MousePointer = vbDefault
   MsgBox "Customer masterdata file has been updated.", vbInformation, Me.Caption
   Me.lv1.ListItems.Clear
   Me.cmdUpdate.Enabled = False
   Me.pb.Value = 0
   Exit Sub
   
ErrorHandler:
MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdCancel_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdDestination_Click()
   Dim strTemp As String
   On Error GoTo errHandler
   strTemp = fBrowseForFolder(Me.hWnd, "Select source path")
   If strTemp <> "" Then
      Me.txtDestination = strTemp
      Me.flb1.Path = Me.txtDestination.Text
   End If
   Exit Sub
   
errHandler:
MsgBox err.Description, vbInformation, Me.Caption
End Sub

Private Sub cmdExtract_Click()
   Dim sFile As String
   If Me.txtDestination.Text <> "" And Me.txtFileName.Text <> "" Then
      If UCase(Mid(Me.txtFileName.Text, Len(Me.txtFileName.Text) - 2, 3)) = "TXT" Then
         If Mid(Me.txtDestination.Text, Len(Me.txtDestination.Text), 1) = "\" Then
            sFile = Me.txtDestination.Text + Me.txtFileName.Text
         Else
            sFile = Me.txtDestination.Text + "\" + Me.txtFileName.Text
         End If
         If FileExists(sFile) = False Then
            MsgBox "Downloaded CUSTOMER masterdata file does not exist.", vbInformation, Me.Caption
            Exit Sub
         End If
         Call ExtractCustomer(sFile)
         If Me.lv1.ListItems.Count <> 0 Then
            Me.cmdUpdate.Enabled = True
         Else
            MsgBox "There are no records for extraction.", vbInformation, Me.Caption
         End If
      Else
         MsgBox "Invalid file type.", vbInformation, Me.Caption
         Exit Sub
      End If
   Else
      MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
   End If
End Sub

Private Sub flb1_DblClick()
   Me.txtFileName.Text = Me.flb1.FileName
End Sub

Private Sub Form_Load()
   Call Disable_System_Menu
   Set rsVESCUS = db.OpenRecordset("SELECT * FROM VESCUS")
   Me.Left = 1800
   Me.Top = 600
   Me.flb1.Path = Me.txtDestination.Text
End Sub

Public Sub ExtractCustomer(varFile As String)
   Dim fnum As Integer
   Dim sTemp As String * 150
   
   Dim varTemp0 As String * 5           'Filler
   Dim varTemp1 As String * 5           'Customer Number
   Dim varTemp2 As String * 35          'Customer Name
   Dim varTemp3 As String * 35          'Address 1
   Dim varTemp4 As String * 35          'Address 2
   Dim varTemp5 As String * 15          'VAT Number
   
   Me.MousePointer = vbHourglass
   Me.lv1.ListItems.Clear
   
   fnum = FreeFile()
   Open varFile For Input As fnum
   While Not EOF(fnum)
      Line Input #fnum, sTemp
      varTemp0 = Mid(sTemp, 1, 5)
      varTemp1 = Mid(sTemp, 6, 5)
      varTemp2 = Mid(sTemp, 11, 35)
      varTemp3 = Mid(sTemp, 46, 35)
      varTemp4 = Mid(sTemp, 81, 35)
      varTemp5 = Mid(sTemp, 116, 15)

      Set j = Me.lv1.ListItems.Add(, , varTemp0)
      j.SubItems(1) = varTemp1
      j.SubItems(2) = varTemp2
      j.SubItems(3) = varTemp3
      j.SubItems(4) = varTemp4
      j.SubItems(5) = varTemp5
      
   Wend
   Close fnum
   Me.MousePointer = vbDefault
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

