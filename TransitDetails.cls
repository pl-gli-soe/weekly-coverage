VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransitDetails"
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


Public t As ITransit ' tutaj jest ukryta EDA
Public EDA_CW As String
Public pn As String
Public plt As String
Public duns As String
Public firstRunout As String ' CW
Public firstCW As String ' CW
Public scac As String
Public kontener As String
Public doh As Double
Public P As PositionHandler

Public Sub fill_eda_cw()

    ' tutaj wiac ze sam konstruktor new nie wystarczy poniewaz eda cw
    ' zapisana jest na poczatku jako ""
    ' i raczej liczyc na tym nie ma jak :D
    
    EDA_CW = CStr(Int(Year(t.mDeliveryDate))) & CStr(Application.WorksheetFunction.IsoWeekNum(t.mDeliveryDate))
End Sub

Private Sub Class_Initialize()
    Set t = New ASN
    Set P = New PositionHandler
    EDA_CW = ""
End Sub

Private Sub Class_Terminate()
    Set t = Nothing
    Set P = Nothing
End Sub
