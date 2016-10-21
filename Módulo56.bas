Attribute VB_Name = "Módulo56"
Sub iMacro_Extract()


Set iim1 = CreateObject("imacros")
iret = iim1.iimInit

For x = 1 To 2
iret = iim1.iimSet("myloop", CStr(x))
iim1.iimPlay ("TESTE")
Next

iret = iim1.iimExit

End Sub

