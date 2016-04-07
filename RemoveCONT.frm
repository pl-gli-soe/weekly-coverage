VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveCONT 
   Caption         =   "Add layer of active containers in coverage"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4530
   OleObjectBlob   =   "RemoveCONT.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveCONT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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


Private Sub AllBtn_Click()
    select_all True
End Sub

Private Sub ClearAllbtn_Click()
    select_all False
End Sub



Private Sub select_all(arg As Boolean)
    
    With Me.ListBox1
        For x = 0 To .ListCount - 1
            .Selected(x) = arg
            
        Next x
    End With
End Sub




Private Sub SubmitBtn_Click()
    Hide
    
    'Dim beg_of_uniq_cont_list As Range
    'Set beg_of_uniq_cont_list = ThisWorkbook.Sheets("register").Range("h2")
    
    make_submit_inner
End Sub

Public Sub make_submit_inner()
    
    StatusBox.Show
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    StatusBox.ProgressBar.Value = 0
    StatusBox.ProgressBar.Max = Range("a2:a" & CStr(Range("a2").End(xlDown).row)).Count
    
    Dim r As Range
    For Each r In Range("a2:a" & CStr(Range("a2").End(xlDown).row))
    
    
        StatusBox.Description.Caption = CStr(r.offset(0, 5))
        StatusBox.Repaint
        DoEvents
    
        
        For x = 0 To Me.ListBox1.ListCount - 1
            If CStr(r.offset(0, 5)) = CStr(Me.ListBox1.List(x)) Then
                If Me.ListBox1.Selected(x) Then
                    ' 26 offset na cont manager
                    r.offset(0, 26) = 1
                    Exit For
                Else
                    r.offset(0, 26) = 0
                    Exit For
                End If
            End If
        Next x
        ProgressIncrease
    Next r
    
    
    ' Range(Range("t2"), Range("t2").End(xlDown)).Calculate
    StatusBox.Hide
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    
    MsgBox "ready!"
End Sub

Private Sub tranzakcja_zmiany_wartosci_kolumnt_cont_manager(s As String)



    ' ok zakladamy w prosty sposob ze aktywnym arkuszem z pewnoscia bedzie flat table so
    Dim flat As Worksheet
    Set flat = ActiveSheet
    
    
    Dim r As Range
    Set r = Range("a2")
    Do
        If CStr(s) = CStr(r.offset(0, 5)) Then
            ' 26 offset cont manager
            
            
        End If
        
        Set r = r.offset(1, 0)
    Loop While r <> ""
    
    
    
End Sub
