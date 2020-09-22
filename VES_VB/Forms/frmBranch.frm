VERSION 5.00
Begin VB.Form frmBranch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " « Setup Branch"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4815
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
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   2175
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
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[F2] view branch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtBranchName 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox txtBranchCode 
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
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Branch Name"
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
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Branch Code"
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
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmBranch"
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
'#           area :  frmBRANCH                #
'#    description :  Code File Branch Update  #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Dim rsSETFLE As DAO.Recordset
Dim rsVESBRN As DAO.Recordset

Private Sub cmdCancel_Click()
   Call Enable_System_Menu
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   If Me.txtBranchCode.Text <> "" And Me.txtBranchName.Text <> "" Then
      Set rsVESBRN = db.OpenRecordset("SELECT brname, branch FROM VESBRN WHERE Branch = '" & txtBranchCode.Text & "'")
      If rsVESBRN.RecordCount <> 0 Then
          ' Update Set File
          Set rsSETFLE = db.OpenRecordset("SELECT * FROM SETFLE")
          If rsSETFLE.RecordCount <> 0 Then
             rsSETFLE.Edit
             rsSETFLE.Fields!Branch = Me.txtBranchCode.Text
             rsSETFLE.Fields!Brname = Me.txtBranchName.Text
             rsSETFLE.Update
             pubBRANCH = rsSETFLE.Fields!Branch
             pubBRNAME = rsSETFLE.Fields!Brname
          Else
             rsSETFLE.AddNew
             rsSETFLE.Fields!Branch = Me.txtBranchCode.Text
             rsSETFLE.Fields!Brname = Me.txtBranchName.Text
             rsSETFLE.Update
             Set rsSETFLE = db.OpenRecordset("SELECT * FROM SETFLE")
             pubBRANCH = rsSETFLE.Fields!Branch
             pubBRNAME = rsSETFLE.Fields!Brname
          End If
          
          mdiAPAY.StatusBar1.Panels(1).Text = "  Branch:  " + pubBRANCH + " - " + pubBRNAME
          Call Enable_System_Menu
          Unload Me
      Else
          sagot = MsgBox("INVALID BRANCH CODE. Try again..", vbInformation, Me.Caption)
          Me.txtBranchCode.SetFocus
      End If
   Else
      sagot = MsgBox("Enter BRANCH CODE. Try again..", vbInformation, Me.Caption)
      Me.txtBranchCode.SetFocus
   End If
End Sub

Private Sub Form_Load()
   ' Disable System Menu
   Call Disable_System_Menu
   Me.txtBranchCode = pubBRANCH
   Me.txtBranchName = pubBRNAME
   
   Me.Top = 250
   Me.Left = 250
   Me.WindowState = 0
End Sub

Private Sub txtBranchCode_Change()
   If Me.txtBranchCode.Text <> "" Then
      Set rsVESBRN = db.OpenRecordset("SELECT brname, branch FROM VESBRN WHERE Branch = '" & txtBranchCode.Text & "'")
      If rsVESBRN.RecordCount <> 0 Then
         Me.txtBranchCode.Text = rsVESBRN.Fields!Branch
         Me.txtBranchName.Text = rsVESBRN.Fields!Brname
      Else
         If Len(Me.txtBranchCode) = 3 Then
            sagot = MsgBox("INVALID BRANCH CODE. Please check your BRANCH list..", vbInformation, Me.Caption)
            If sagot = vbOK Then
               Me.txtBranchCode.Text = ""
               Me.txtBranchName.Text = ""
               Me.txtBranchCode.SetFocus
            End If
         End If
      End If
   Else
      Me.txtBranchName.Text = ""
   End If
End Sub

Private Sub txtBranchCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Me.txtBranchCode.Text <> "" Then
         Set rsVESBRN = db.OpenRecordset("SELECT brname, branch FROM VESBRN WHERE Branch = '" & txtBranchCode.Text & "'")
         If rsVESBRN.RecordCount <> 0 Then
            Me.txtBranchCode.Text = rsVESBRN.Fields!Branch
            Me.txtBranchName.Text = rsVESBRN.Fields!Brname
            Me.cmdUpdate.SetFocus
         Else
            Me.txtBranchName.Text = ""
         End If
      Else
         Me.txtBranchName.Text = ""
      End If
   End If
End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2
           Set rsVESBRN = db.OpenRecordset("SELECT vesbrn.brname, vesbrn.branch FROM VESBRN order by brname")
           frmViewBranch.lv1.ListItems.Clear
           rsVESBRN.MoveFirst
           Do While Not rsVESBRN.EOF
              Set x = frmViewBranch.lv1.ListItems.Add(, , rsVESBRN.Fields!Branch)
              x.SubItems(1) = rsVESBRN.Fields!Brname
              rsVESBRN.MoveNext
           Loop
           frmViewBranch.lv1.SetFocus
           frmViewBranch.Show
   End Select
End Sub

