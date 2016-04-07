Attribute VB_Name = "ScanarioModule"
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




' all asns
Public Sub scenario0(ictrl As IRibbonControl)
    ThisWorkbook.Sheets("register").Range("togglehandler") = 0
End Sub


' st < 1
Public Sub scenario1(ictrl As IRibbonControl)
    ThisWorkbook.Sheets("register").Range("togglehandler") = 1
End Sub

' st < 2
Public Sub scenario2(ictrl As IRibbonControl)
    ThisWorkbook.Sheets("register").Range("togglehandler") = 2
End Sub


' st < 3
Public Sub scenario3(ictrl As IRibbonControl)
    ThisWorkbook.Sheets("register").Range("togglehandler") = 3
End Sub


' st < 4
Public Sub scenario4(ictrl As IRibbonControl)
    ThisWorkbook.Sheets("register").Range("togglehandler") = 4
End Sub



' doh
Public Sub scenario7(ictrl As IRibbonControl)
    ThisWorkbook.Sheets("register").Range("togglehandler") = 7
End Sub

