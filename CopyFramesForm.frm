VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyFramesForm 
   Caption         =   "Copy Frames Form"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3255
   OleObjectBlob   =   "CopyFramesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CopyFramesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCopyASNOffset_Click()
    Hide
    copy_asn_offset
End Sub

Private Sub BtnCopyMISCQHD_Click()
    Hide
    copy_misc_qhd
End Sub

Private Sub BtnCopyOrange_Click()
    Hide
    copy_orange
End Sub

Private Sub BtnJusCopy_Click()
    Hide
    just_copy_this_sheet
End Sub

Private Sub BtnRedFrames_Click()
    Hide
    copy_red_frames
End Sub


Private Sub BtnRunOutOrder_Click()
    Hide
    copy_with_runout_order
End Sub
