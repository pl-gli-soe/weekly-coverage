Attribute VB_Name = "DeleteModule"
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


Public Sub delete_this_sheet(ictrl As IRibbonControl)
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If (ActiveSheet.Name Like "*Parts*") Or (ActiveSheet.Name Like "*register*") Then
        MsgBox "you can't delete this sheet!"
    Else
        ActiveSheet.Delete
    End If
    
    
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub


Public Sub delete_all_sheets(ictrl As IRibbonControl)

    Application.EnableEvents = False

    ret = MsgBox("Are you sure?", vbQuestion + vbYesNo)
    If ret = vbYes Then
        Application.DisplayAlerts = False
        
        x = 1
        Do
            If (Sheets(x).Name Like "*Parts*") Or (Sheets(x).Name Like "*register*") Then
                x = x + 1
            Else
                Sheets(x).Delete
            End If
        Loop Until x > Sheets.Count
        Application.DisplayAlerts = True
    End If
    
    
    Application.EnableEvents = True
End Sub
