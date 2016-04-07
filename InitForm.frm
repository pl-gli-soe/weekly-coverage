VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InitForm 
   Caption         =   "Init Form"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3990
   OleObjectBlob   =   "InitForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InitForm"
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



Private Sub CheckBoxDelayFlag_Change()

    If Me.CheckBoxDelayFlag.Value Then
        Me.TextBoxDelayFlagBoundary.Enabled = True
        Me.TextBoxDelayFlagBoundary.Value = "7"
    Else
        Me.TextBoxDelayFlagBoundary.Enabled = False
        Me.TextBoxDelayFlagBoundary.Value = "0"
    End If

End Sub


Private Sub SubmitButton_Click()
    Me.Hide
    
    ' z TCForm
    ThisWorkbook.Sheets("register").Range("tc") = Me.TextBoxTC.Value
    ThisWorkbook.Sheets("register").Range("edaminussdate") = Me.TextBoxEDASDATE.Value
    
    Dim auto As AutoComponentHandler
    Set auto = New AutoComponentHandler

    If Me.OptionButtonSTD Then
        MakeFullCoverage STD, Me.CheckBoxCheckStatuses.Value, Me.CheckBoxCopyCritical.Value, Me.CheckBoxDelayFlag.Value, Me.CheckBoxQHDMISC.Value
    ElseIf Me.OptionButtonGREEN Then
        MakeFullCoverage GREEN, Me.CheckBoxCheckStatuses.Value, Me.CheckBoxCopyCritical.Value, Me.CheckBoxDelayFlag.Value, Me.CheckBoxQHDMISC.Value
    ElseIf Me.OptionButtonGB Then
        MakeFullCoverage GREEN_AND_BLUE, Me.CheckBoxCheckStatuses.Value, Me.CheckBoxCopyCritical.Value, Me.CheckBoxDelayFlag.Value, Me.CheckBoxQHDMISC.Value
    End If
    
    
    Range("A2") = Replace(CStr(CLng(Now)), ",", "")
    Set auto = Nothing
    
End Sub
