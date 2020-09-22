Attribute VB_Name = "dbModule"
'##############################################
'#          Coded by Adel D. Mendoza          #
'#        Designed by Ronald S. Abian         #
'#  VRS - VAT Input/Output Reporting System   #
'#           for SWIFT FOODS, INC.            #
'#                                            #
'#           area :  dbModule                 #
'#    description :  Code File Module 1       #
'#        e-mail  :  adm@rfm.com.ph           #
'#        url     :  http://www.rfm.com.ph    #
'#                                            #
'##############################################

Option Explicit

Public pubAccLevel
Public pubUserName
Public pubPassword
Public iMonth(12)
Public cCode
Public rptPointer
Public cType As String
Public mnuPointer As Integer
Public pubDIV As String
Public pubBRNAME As String
Public pubCONAME As String
Public pubRDRIVE As String
Public pubPWORD As String
Public pubLASDAT As String
Public pubBRANCH As String
Public pubYEAR As String
Public pubMONTH As String
Public pubQUARTER As String
Public db As DAO.Database

Public Sub openDB()
   Set db = OpenDatabase(App.Path & "\VES.mdb")
End Sub

Public Function execQuery(ByVal sqlStr As String) As String
   db.Execute (sqlStr)
End Function

Public Sub load_Month()
   iMonth(1) = "JANUARY"
   iMonth(2) = "FEBRUARY"
   iMonth(3) = "MARCH"
   iMonth(4) = "APRIL"
   iMonth(5) = "MAY"
   iMonth(6) = "JUNE"
   iMonth(7) = "JULY"
   iMonth(8) = "AUGUST"
   iMonth(9) = "SEPTEMBER"
   iMonth(10) = "OCTOBER"
   iMonth(11) = "NOVEMBER"
   iMonth(12) = "DECEMBER"
End Sub

Public Sub Enable_System_Menu()
   mdiAPAY.mnuDataEntry.Enabled = True
   mdiAPAY.mnuFileUpdate.Enabled = True
   mdiAPAY.mnuFileMaintenance.Enabled = True
   mdiAPAY.mnuUtilities.Enabled = True
   mdiAPAY.mnuTools.Enabled = True
   mdiAPAY.mnuReports.Enabled = True
   mdiAPAY.mnuSystem.Enabled = True
End Sub

Public Sub Disable_System_Menu()
   mdiAPAY.mnuDataEntry.Enabled = False
   mdiAPAY.mnuFileUpdate.Enabled = False
   mdiAPAY.mnuFileMaintenance.Enabled = False
   mdiAPAY.mnuUtilities.Enabled = False
   mdiAPAY.mnuTools.Enabled = False
   mdiAPAY.mnuReports.Enabled = False
   mdiAPAY.mnuSystem.Enabled = False
End Sub
