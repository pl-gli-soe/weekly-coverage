VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASN"
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



Implements ITransit

Private route As String
Private iNAME As String
Private qty As Long
Private delivery_date As Date
Private delivery_time As Date
Private SDATE As Date
Private kontener As String
Private st As String
Private mode As String


' estimated arrival at the port
Private arrival_at_the_port As Date

Private not_yet_received As Boolean
Private out_of_scope As Boolean
Private on_pop As Boolean





Private Sub Class_Initialize()
    not_yet_received = True
    out_of_scope = False
End Sub

Private Property Let ITransit_mArrivalAtThePort(RHS As Date)
    arrival_at_the_port = CDate(RHS)
End Property

Private Property Get ITransit_mArrivalAtThePort() As Date
    ITransit_mArrivalAtThePort = CDate(arrival_at_the_port)
End Property

Private Property Let ITransit_mDeliveryDate(RHS As Date)
    delivery_date = RHS
End Property

Private Property Get ITransit_mDeliveryDate() As Date
    ITransit_mDeliveryDate = delivery_date
End Property

Private Property Let ITransit_mDeliveryTime(RHS As Date)
    delivery_time = RHS
End Property

Private Property Get ITransit_mDeliveryTime() As Date
    ITransit_mDeliveryTime = delivery_time
End Property

Private Property Let ITransit_mMODE(RHS As String)
    mode = RHS
End Property

Private Property Get ITransit_mMODE() As String
    ITransit_mMODE = mode
End Property

Private Property Let ITransit_mName(RHS As String)
    iNAME = CStr(RHS)
End Property

Private Property Get ITransit_mName() As String
    ITransit_mName = iNAME
End Property

Private Property Let ITransit_mNotYetReceived(RHS As Boolean)
    not_yet_received = RHS
End Property

Private Property Get ITransit_mNotYetReceived() As Boolean
    ITransit_mNotYetReceived = not_yet_received
End Property

Private Property Let ITransit_mOnPop(RHS As Boolean)
    on_pop = RHS
End Property

Private Property Get ITransit_mOnPop() As Boolean
    ITransit_mOnPop = on_pop
End Property

Private Property Let ITransit_mOutOfScope(RHS As Boolean)
    out_of_scope = RHS
End Property

Private Property Get ITransit_mOutOfScope() As Boolean
    ITransit_mOutOfScope = out_of_scope
End Property

Private Property Let ITransit_mPickupDate(RHS As Date)
    SDATE = RHS
End Property

Private Property Get ITransit_mPickupDate() As Date
    ITransit_mPickupDate = SDATE
End Property

Private Property Let ITransit_mPickupTime(RHS As Date)
    
End Property

Private Property Get ITransit_mPickupTime() As Date
    ITransit_mPickupTime = CDate("00:00")
End Property

Private Property Let ITransit_mQty(RHS As Long)
    qty = RHS
End Property

Private Property Get ITransit_mQty() As Long
    ITransit_mQty = qty
End Property

Private Property Let ITransit_mRoute(RHS As String)
    route = RHS
End Property

Private Property Get ITransit_mRoute() As String
    ITransit_mRoute = route
End Property

Private Property Let ITransit_mST(RHS As String)
    st = CStr(RHS)
End Property

Private Property Get ITransit_mST() As String
    ITransit_mST = CStr(st)
End Property

Private Property Let ITransit_mTRLR(RHS As String)
    kontener = CStr(RHS)
End Property

Private Property Get ITransit_mTRLR() As String
    ITransit_mTRLR = CStr(kontener)
End Property

Private Property Let ITransit_mTTLessThan24h(RHS As Boolean)
    ITransit_mTTLessThan24h = RHS
End Property

Private Property Get ITransit_mTTLessThan24h() As Boolean
    ITransit_mTTLessThan24h = False
End Property


' connected with ORANGE feature
' ==========================================================================================
Private Sub ITransit_tryToAssignToThisAsnDateArrivalAtThePort(str As String, imgo As MGO)


    ' the str variable aggragate line of string which includes information such as comment of rdc arrival or port arrival
    ' need to implement simple parer, which will have capabilites to throw out important information on top and adjust it
    ' into prooper date type of variable - and the type is necessecity due to furhter calculations
    ' testing for today s date
    ' arrival_at_the_port = Date
    
    ' the format of the date is simple: dd.mm.rr
    
    For x = 1 To Len(str) - 8
        str_date = checkPattern(CStr(Mid(str, CLng(x), 8)))
        If Trim(str_date) <> "" Then
            Exit For
        End If
    Next x
    
    If Trim(str_date) <> "" Then
        ' now we have some date we can adjsut it to proper format
        arrival_at_the_port = CDate(imgo.convertToDateFromMS9POP00Date(CStr(str_date)))
    End If
End Sub
' ==========================================================================================

Private Function checkPattern(dd_dot_mm_dot_rr As String) As String
    checkPattern = ""
    
    ' 1st instance - check if first 2 characters is number
    If IsNumeric(Left(dd_dot_mm_dot_rr, 2)) Then
    
    
        ' check if it fits in month days (2nd instance)
        ' found and issue when name of month was not matched need to be defined  :P
        ' da function checkMonthName checking if the naming is OK
        If (CLng(Left(dd_dot_mm_dot_rr, 2)) > 0) And (CLng(Left(dd_dot_mm_dot_rr, 2)) < 32) And (checkMonthName(Mid(dd_dot_mm_dot_rr, 4, 2))) Then
        
            ' 3rd check dots - if there are in right places
            If (Mid(dd_dot_mm_dot_rr, 3, 1) = ".") And (Mid(dd_dot_mm_dot_rr, 6, 1) = ".") Then
            
                ' 4th check year - is num? - the last one - not we almost sure that we found in this comment date in proper format
                If IsNumeric(Right(dd_dot_mm_dot_rr, 2)) Then
                    
                    
                   ' 5th - check month if it is some letters (not numbers)
                   If Not IsNumeric(Mid(dd_dot_mm_dot_rr, 4, 2)) Then
                        checkPattern = CStr(dd_dot_mm_dot_rr)
                    End If
                End If
            End If
        End If
    End If
End Function

Private Function checkMonthName(m As String) As Boolean
    checkMonthName = False
    
    
    ' ta opcja z minus jeden wymusiala na  mnie zmiane ukladu select case w jendnej z najstarszych funkcji zawartych w tej implemntacji
    ' i musialem sie posluzyc select case z elementami else
    ' :D
    ' 2015-07-13
    If MgoMonth(m) = -1 Then
        ' tutaj faktycznie jest dupa, zatem zostaje false i koncze prace z ta funkcja
        Exit Function
    Else
        ' tutaj jest ok
        ' zatem zmieniam wartosc zmiennej
        checkMonthName = True
        
        ' no i koncze prace z funkcja - dalej nie trza :D
        Exit Function
    End If
End Function
