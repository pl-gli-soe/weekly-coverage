VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITransit"
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



Property Get mArrivalAtThePort() As Date
End Property

Property Let mArrivalAtThePort(arg As Date)
End Property

Public Sub tryToAssignToThisAsnDateArrivalAtThePort(str As String, imgo As MGO)
End Sub

Property Get mNotYetReceived() As Boolean
End Property

Property Let mNotYetReceived(arg As Boolean)
End Property

Property Get mOutOfScope() As Boolean
End Property

Property Let mOutOfScope(arg As Boolean)
End Property

Property Get mQty() As Long
End Property

Property Let mQty(arg As Long)
End Property

Property Get mPickupDate() As Date
End Property

Property Let mPickupDate(arg As Date)
End Property

Property Get mPickupTime() As Date
End Property

Property Let mPickupTime(arg As Date)
End Property

Property Get mDeliveryDate() As Date
End Property

Property Let mDeliveryDate(arg As Date)
End Property

Property Get mDeliveryTime() As Date
End Property

Property Let mDeliveryTime(arg As Date)
End Property

Property Get mName() As String
End Property

Property Let mName(arg As String)
End Property

Property Get mRoute() As String
End Property

Property Let mRoute(arg As String)
End Property

Property Get mOnPop() As Boolean
End Property

Property Let mOnPop(arg As Boolean)
End Property

Property Get mTTLessThan24h() As Boolean
End Property

Property Let mTTLessThan24h(arg As Boolean)
End Property

Property Get mTRLR() As String
End Property

Property Let mTRLR(arg As String)
End Property

Property Get mST() As String
End Property

Property Let mST(arg As String)
End Property

Property Get mMODE() As String
End Property

Property Let mMODE(arg As String)
End Property
