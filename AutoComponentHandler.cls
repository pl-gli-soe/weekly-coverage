VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoComponentHandler"
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



Private h As Range
Private P As Range
Private plants As Range

Private Sub Class_Initialize()

    Set h = ThisWorkbook.Sheets("register").Range("h2")
    Set P = ThisWorkbook.Sheets("Parts").Range("a2")
    
    If ThisWorkbook.Sheets("Parts").Range("a3") <> "" Then
        Set plants = Range(P, P.End(xlDown))
    Else
        Set plants = P
    End If
    

    ' we can run it already in constructor
    ' we already define configuration
    ' so no check the init list for this coverage
    ' if there's any C to should put
    
    Dim r As Range
    For Each r In plants
        If check_if_component(r, h) Then
            r.offset(0, 3) = "C"
        Else
            r.offset(0, 3) = ""
        End If
    Next r
    
    
End Sub

Private Function check_if_component(r As Range, h As Range) As Boolean

    check_if_component = False
    
    Dim i As Range
    For Each i In Range(h, h.End(xlDown))
        If CStr(i) = CStr(r) Then
            check_if_component = True
            Exit For
        End If
    Next i
End Function



