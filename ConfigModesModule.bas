Attribute VB_Name = "ConfigModesModule"
Public Sub showConfigModesForm()
    
    ConfigModesForm.TextBoxAir.Value = CStr(ThisWorkbook.Sheets("register").Range("air"))
    ConfigModesForm.TextBoxSea.Value = CStr(ThisWorkbook.Sheets("register").Range("sea"))
    
    
    ConfigModesForm.Show
End Sub


Public Sub ribbon_showConfigModesForm(ictrl As IRibbonControl)
    showConfigModesForm
End Sub
