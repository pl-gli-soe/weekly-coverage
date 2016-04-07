VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MS9PO400"
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


Private Sess0 As Object

Public Property Let sesja(ByRef arg As Object)
    Set Sess0 = arg
End Property

Public Property Get sesja() As Object
    Set sesja = Sess0
End Property

Public Property Get pn() As String
  pn = Trim(Sess0.Screen.getString(3, 19, 8))
End Property

Public Property Let pn(arg As String)
  Sess0.Screen.putString CStr(arg), 3, 19
End Property

Public Property Get plt() As String
  pn = Trim(Sess0.Screen.getString(3, 7, 2))
End Property

Public Property Let plt(arg As String)
  Sess0.Screen.putString CStr(arg), 3, 7
End Property

Public Property Get KANBAN() As String
  KANBAN = Trim(Sess0.Screen.getString(3, 35, 4))
End Property

Public Property Let KANBAN(arg As String)
  Sess0.Screen.putString CStr(arg), 3, 35
End Property



Public Function qty(i As Integer) As String
    qty = Trim(Sess0.Screen.getString(6 + 2 * i, 5, 8))
End Function
Public Function SDATE(i As Integer) As String
    SDATE = Trim(Sess0.Screen.getString(6 + 2 * i, 14, 6))
End Function
Public Function PCS(i As Integer) As String
    PCS = Trim(Sess0.Screen.getString(6 + 2 * i, 23, 3))
End Function
Public Function WGT(i As Integer) As String
    WGT = Trim(Sess0.Screen.getString(6 + 2 * i, 27, 6))
End Function
Public Function mode(i As Integer) As String
    mode = Trim(Sess0.Screen.getString(6 + 2 * i, 35, 2))
End Function
Public Function CD(i As Integer) As String
    CD = Trim(Sess0.Screen.getString(6 + 2 * i, 42, 1))
End Function
Public Function EDA(i As Integer) As String
    EDA = Trim(Sess0.Screen.getString(6 + 2 * i, 46, 6))
End Function
Public Function ETA(i As Integer) As String
    ETA = Trim(Sess0.Screen.getString(6 + 2 * i, 53, 4))
End Function
Public Function st(i As Integer) As String
    st = Trim(Sess0.Screen.getString(6 + 2 * i, 58, 1))
End Function
Public Function SID(i As Integer) As String
    SID = Trim(Sess0.Screen.getString(6 + 2 * i, 61, 9))
End Function
Public Function route(i As Integer) As String
    route = Trim(Sess0.Screen.getString(6 + 2 * i, 71, 7))
End Function
Public Function DOCK(i As Integer) As String
    DOCK = Trim(Sess0.Screen.getString(7 + 2 * i, 71, 7))
End Function
Public Function SHIP_CMNT(i As Integer) As String
    SHIP_CMNT = Trim(Sess0.Screen.getString(7 + 2 * i, 38, 32))
End Function
Public Function CARR(i As Integer) As String
    CARR = Trim(Sess0.Screen.getString(7 + 2 * i, 21, 16))
End Function
Public Function TRLR(i As Integer) As String
    TRLR = Trim(Sess0.Screen.getString(7 + 2 * i, 21, 16))
End Function
Public Function SUPPLIER(i As Integer) As String
    SUPPLIER = Trim(Sess0.Screen.getString(7 + 2 * i, 5, 15))
End Function
Public Function s(i As Integer) As String
    s = Trim(Sess0.Screen.getString(6 + 2 * i, 3, 1))
End Function

Public Sub C_SHIP_CMNT(str As String)
    Sess0.Screen.putString CStr(str), 7, 38
End Sub
