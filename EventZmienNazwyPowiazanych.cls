VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventZmienNazwyPowiazanych"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'      ___           ___                       ___                    ___           ___           ___           ___
'     /\  \         /\  \          ___        /\  \                  /\  \         /\  \         /\  \         /\__\
'    /::\  \       /::\  \        /\  \      /::\  \                 \:\  \       /::\  \       /::\  \       /:/  /
'   /:/\:\  \     /:/\:\  \       \:\  \    /:/\:\  \                 \:\  \     /:/\:\  \     /:/\:\  \     /:/  /
'  /::\~\:\  \   /::\~\:\  \      /::\__\  /:/  \:\  \                /::\  \   /:/  \:\  \   /:/  \:\  \   /:/  /
' /:/\:\ \:\__\ /:/\:\ \:\__\  __/:/\/__/ /:/__/ \:\__\              /:/\:\__\ /:/__/ \:\__\ /:/__/ \:\__\ /:/__/
' \/__\:\/:/  / \/_|::\/:/  / /\/:/  /    \:\  \ /:/  /             /:/  \/__/ \:\  \ /:/  / \:\  \ /:/  / \:\  \
'      \::/  /     |:|::/  /  \::/__/      \:\  /:/  /             /:/  /       \:\  /:/  /   \:\  /:/  /   \:\  \
'       \/__/      |:|\/__/    \:\__\       \:\/:/  /              \/__/         \:\/:/  /     \:\/:/  /     \:\  \
'                  |:|  |       \/__/        \::/  /                              \::/  /       \::/  /       \:\__\
'                   \|__|                     \/__/                                \/__/         \/__/         \/__/
'


Public WithEvents zdarzenie As EventCzyZmienionoNazweArkusza
Attribute zdarzenie.VB_VarHelpID = -1


Private Sub Class_Initialize()
    Set zdarzenie = New EventCzyZmienionoNazweArkusza
End Sub

Private Sub Class_Terminate()
    Set zdarzenie = Nothing
End Sub

Private Sub zdarzenie_zmienNazwyPowiazanychArkuszyWTymPliku(arg As Worksheet, theoldonename As String)


    If arg.Name <> "Parts" Then
        If arg.Name <> "register" Then
    
            Dim mstr As String
            Dim newone As String
            mstr = CStr(theoldonename)
            mstr = CStr(remove_prefix(mstr))
            newone = arg.Name
            newone = CStr(remove_prefix(newone))
            
            Dim sh As Worksheet
            For Each sh In ThisWorkbook.Sheets
                ' tutaj pelta sprawdzjaca wszystie powiazane arkusze nazwami
                If not_config_sheets(sh.Name) Then
                    If sh.Name Like "*" & mstr Then
                        If sh.Name <> arg.Name Then
                            On Error Resume Next
                            sh.Name = Replace(sh.Name, mstr, newone)
                        End If
                    End If
                End If
            Next sh
            
            ThisWorkbook.Sheets("register").Range("actSheetName") = CStr(arg.Name)
        End If
    End If
End Sub


Private Function remove_prefix(mstr As String) As String
    remove_prefix = Replace(mstr, "FLAT_", "")
    remove_prefix = Replace(remove_prefix, "PIVOTD_", "")
    remove_prefix = Replace(remove_prefix, "PIVOTW_", "")
    remove_prefix = Replace(remove_prefix, "order_", "")
    remove_prefix = Replace(remove_prefix, "copy_", "")
    remove_prefix = Replace(remove_prefix, "c_", "")
End Function

Private Function not_config_sheets(s As String) As Boolean
    If CStr(s) = "register" Or CStr(s) = "Parts" Then
        not_config_sheets = False
    Else
        not_config_sheets = True
    End If
End Function
