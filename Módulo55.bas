Attribute VB_Name = "Módulo55"
Option Explicit
   
   Private Sub CommandButton1_Click()
   
   MsgBox "This macro demonstrates how to read data from an Excel sheet and submit this information to a website."
   
   Dim iim1, iret, row, totalrows
   
   Set iim1 = CreateObject("imacros")
   iret = iim1.iimInit
   'Firefox iret = iim1.iimInit ("-fx")
   iret = iim1.iimDisplay("Submitting Data from Excel")
   
   totalrows = ActiveSheet.UsedRange.Rows.Count
   For row = 2 To totalrows
      'Set the variables
      iret = iim1.iimSet("FNAME", Cells(row, 1).Value)
      iret = iim1.iimSet("LNAME", Cells(row, 2).Value)
      iret = iim1.iimSet("ADDRESS", Cells(row, 3).Value)
      iret = iim1.iimSet("CITY", Cells(row, 4).Value)
      iret = iim1.iimSet("ZIP", Cells(row, 5).Value)
      iret = iim1.iimSet("STATE-ID", Cells(row, 6).Value)
      iret = iim1.iimSet("COUNTRY-ID", Cells(row, 7).Value)
      iret = iim1.iimSet("EMAIL", Cells(row, 8).Value)
      'Set the display
      iret = iim1.iimDisplay("Row# " + CStr(row))
      'Run the macro
      'Same macro as in database-2-web.vbs example!
      iret = iim1.iimPlay("wsh-submit-2-web")
      If iret < 0 Then
         MsgBox iim1.iimGetLastError()
      End If
   Next row
   
   iret = iim1.iimDisplay("Submission complete")
   iret = iim1.iimExit
   
   End Sub
