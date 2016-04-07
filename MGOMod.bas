Attribute VB_Name = "MGOMod"
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



Global Sessions As Object
Global System As Object
Global Sess0 As Object

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub mgoInit()
    Set System = CreateObject("EXTRA.System")   ' Gets the system object
    Set Sessions = System.Sessions
    Set Sess0 = System.ActiveSession
End Sub

Public Sub waitForMgo()
    Do
        DoEvents
    Loop Until Sess0.Screen.OIA.XStatus = 0
End Sub

Function MgoMonth(CurrDate) As Integer
    Select Case CurrDate
        Case "JA"
            MgoMonth = 1
        Case "FE"
            MgoMonth = 2
        Case "MR"
            MgoMonth = 3
        Case "AP"
            MgoMonth = 4
        Case "MY"
            MgoMonth = 5
        Case "JN"
            MgoMonth = 6
        Case "JL"
            MgoMonth = 7
        Case "AU"
            MgoMonth = 8
        Case "SE"
            MgoMonth = 9
        Case "OC"
            MgoMonth = 10
        Case "NO"
            MgoMonth = 11
        Case "DE"
            MgoMonth = 12
        Case Else
            MgoMonth = -1
    End Select
End Function

Sub ProgressIncrease()
        If StatusBox.ProgressBar.Value = StatusBox.ProgressBar.Max Then
            StatusBox.ProgressBar.Max = StatusBox.ProgressBar.Max + 1
        End If
        StatusBox.ProgressBar.Value = StatusBox.ProgressBar.Value + 1
End Sub




' ==============================================================================================

' for running macro
Public Sub run_prognoza_d(ictrl As IRibbonControl)
    Dim eh As ExtraHandler
    Set eh = New ExtraHandler
    
    eh.run_macro "prognoza"
    
    Set eh = Nothing
    
    ' MsgBox "ready!"
End Sub

Public Sub run_prognoza_h(ictrl As IRibbonControl)
    Dim eh As ExtraHandler
    Set eh = New ExtraHandler
    
    eh.run_macro "prognoza_h"
    
    Set eh = Nothing
    
    ' MsgBox "ready!"
End Sub

' ==============================================================================================
