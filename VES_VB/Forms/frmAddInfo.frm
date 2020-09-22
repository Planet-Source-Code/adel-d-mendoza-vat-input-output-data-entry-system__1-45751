VERSION 5.00
Begin VB.Form frmAddInfo 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " « VAT Input/Output Data Entry System"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3855
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   3375
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmAddInfo.frx":000C
         Left            =   2160
         List            =   "frmAddInfo.frx":00A9
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmAddInfo.frx":01DF
         Left            =   2160
         List            =   "frmAddInfo.frx":0207
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         Caption         =   " Enter Month"
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Caption         =   " Enter Year"
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H8000000B&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Additional Information"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmAddInfo"
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
'#         area :  frmADDinfo                 #
'#  description :  Code File Additional Info  #
'#      e-mail  :  adm@rfm.com.ph             #
'#      url     :  http://www.rfm.com.ph      #
'#                                            #
'##############################################

Dim rsSETFLE As DAO.Recordset

Private Sub cboMonth_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Private Sub cmdCancel_Click()
   ' Enable system menu
   rsSETFLE.CancelUpdate
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdOk_Click()
   Dim varYear
   varYear = Year(Date)
   
   If Me.cboYear <> "" And Me.cboMonth.Text <> "" Then
      If Me.cboYear > varYear Then
         MsgBox "Invalid year.", vbInformation, Me.Caption
         Me.cboYear.SetFocus
         Exit Sub
      End If
      rsSETFLE.Fields!Year = Me.cboYear
      rsSETFLE.Fields!Month = Me.cboMonth
      rsSETFLE.Update
      pubYEAR = Me.cboYear
      pubMONTH = Me.cboMonth
      ' Call DATA ENTRY FORMS
      Select Case mnuPointer
         Case 1
           Unload Me
           frmVATinput.Show     'for VAT Input
         Case 2
           Unload Me
           cType = "1"
           frmVAToutput.Show    'for VAT Output
         Case 3
           Unload Me
           cType = "3"
           frmVAToutput.Show    'for VAT Output (Exempt)
         Case 4
           Unload Me
           frmTax.Show          'for TAX input
      End Select
   Else
      MsgBox "Fill up all the required fields.", vbInformation, Me.Caption
      Me.cboYear.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Set rsSETFLE = db.OpenRecordset("SELECT * FROM SETFLE")
   If rsSETFLE.RecordCount <> 0 Then
      pubYEAR = rsSETFLE.Fields!Year
      pubMONTH = rsSETFLE.Fields!Month
      rsSETFLE.Edit
      '==========='
      Me.Top = 2200
      Me.Left = 4100
      Me.WindowState = 0
      Me.cboYear = pubYEAR
      Me.cboMonth = pubMONTH
      ' Disable System Menu
      Call Disable_System_Menu
   End If
End Sub

Private Sub cboYear_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{tab}"
   End If
End Sub
