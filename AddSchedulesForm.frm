VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddSchedulesForm 
   Caption         =   "Add schedules"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
   OleObjectBlob   =   "AddSchedulesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddSchedulesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' poniewaz w jakis sposob chcialem zostac z nienaruszona implementacja z poprzedniej powersji coverage
' zrobilem tak ze labelki nie maja juz wpisanego poprawnie textu
' wpisuje sie dynamicznie dopiero po uruchomieniu add schedule / add schedules z ribbona :D

Public nm As String

Private Sub BtnSubmit_Click()
    
    Hide
    If nm <> "1" And nm <> "2" Then
        On Error Resume Next
        If CLng(Me.TextBoxQty.Value) < 0 Then
            MsgBox "sth went wrong!"
            End
        End If
    End If
    
End Sub

