VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigModesForm 
   Caption         =   "MODE Config"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   OleObjectBlob   =   "ConfigModesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConfigModesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitBtn_Click()
    
    
    ConfigModesForm.Hide
    ThisWorkbook.Sheets("register").Range("air") = CStr(ConfigModesForm.TextBoxAir.Value)
    ThisWorkbook.Sheets("register").Range("sea") = CStr(ConfigModesForm.TextBoxSea.Value)
End Sub
