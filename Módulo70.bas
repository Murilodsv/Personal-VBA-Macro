Attribute VB_Name = "Módulo70"
Sub Macro55()
'
' Macro55 Macro
'

'

For x = 123035 To 126283


Windows("lista.xlsx").Activate
arquivo = Sheets("lista").Range("A" & x).Value
Workbooks.Open Filename:="C:\Murilo\DOUTORADO\Series Clima\SERIES_CLIMA\Data\" & arquivo & ""

Application.DisplayAlerts = False

ActiveWorkbook.SaveAs Filename:= _
        "C:\Murilo\DOUTORADO\Series Clima\SERIES_CLIMA\CSV\" & arquivo & "" _
        , FileFormat:=xlCSVMSDOS, CreateBackup:=False


    ActiveWindow.Close


Next



End Sub

