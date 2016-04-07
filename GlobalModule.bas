Attribute VB_Name = "GlobalModule"
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



Public Enum layout_type
    STD
    GREEN
    GREEN_AND_BLUE
End Enum

Public Enum pivot_layout
    EDACW
    EDA
End Enum

Public Enum starter_dla_listy_kontenerow_w_formularzu
    GUZIK
    PIVOT_SCENARIO
End Enum

Global zmiana_nazw_powiazanych_arkuszy As EventZmienNazwyPowiazanych
Global rng As Range

Global s As Worksheet
Global this_workbook As Workbook

Global G_ASN01 As Boolean
Global G_CP_CRITICAL As Boolean
Global G_DELAY_FLAG As Boolean
Global G_QHD_MISC As Boolean
Global G_DELAY_FLAG_BOUNDARY As Integer

Global Const G_COMMENT_WIDTH = 140
Global Const G_COMMENT_HEIGHT = 140
