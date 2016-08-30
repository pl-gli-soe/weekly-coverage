VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TworzenieTranzytow"
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


Public Sub on_ms9po400(ByVal IIteracja_PlantName As String, _
    ByVal IIteracja_Partnum As String, ByRef i_mgo As MGO, _
    ByRef kolekcja As Collection)


        ' handler pod tranzyt
        Dim temp_trans As ITransit
        Dim is_f8_pressed As Boolean
        
        ' open po400
        i_mgo.sendKeys "<Clear>"
        i_mgo.sendKeys "ms9po400 <Enter>"
        i_mgo.pMS9PO400.plt = IIteracja_PlantName
        i_mgo.pMS9PO400.pn = IIteracja_Partnum
        i_mgo.pMS9PO400.KANBAN = ""
        i_mgo.sendKeys "<Enter>"
        
        
        If Not check_if_nsr_issue_appeared(i_mgo) Then
        
            
            
            ' tutaj petla zapelniajaca nyrt
            ' ===================================================================
            
            Do
                mgoInit
                For x = 0 To 7
                    
                    m_str = CStr(i_mgo.pMS9PO400.EDA(Int(x)))
                    If m_str <> "______" Then
                    
                        txt_time = Left(i_mgo.pMS9PO400.ETA(Int(x)), 2) & ":" & Right(i_mgo.pMS9PO400.ETA(Int(x)), 2)
                        If (txt_time = "__:__") Or (txt_time = "24:00") Then
                            txt_time = "00:00"
                        End If
                        
                        Set temp_trans = Nothing
                        Set temp_trans = New ASN
                        temp_trans.mDeliveryDate = i_mgo.convertToDateFromMS9PO400Date(i_mgo.pMS9PO400.EDA(Int(x)))
                        temp_trans.mDeliveryTime = CDate(Format(txt_time, "hh:mm"))
                        temp_trans.mNotYetReceived = True
                        temp_trans.mName = i_mgo.pMS9PO400.SID(Int(x))
           
                        
                        temp_trans.mTRLR = i_mgo.pMS9PO400.TRLR(Int(x))
                        temp_trans.mST = i_mgo.pMS9PO400.st(Int(x))
                        temp_trans.mPickupDate = i_mgo.convertToDateFromMS9PO400Date(CStr(i_mgo.pMS9PO400.SDATE(Int(x))))
                        temp_trans.mQty = CLng(i_mgo.pMS9PO400.qty(Int(x)))
                        temp_trans.mRoute = i_mgo.pMS9PO400.route(Int(x))
                        temp_trans.mMODE = Trim(i_mgo.pMS9PO400.mode(Int(x)))
                        
                        
                                 
                        If Trim(temp_trans.mName) Like "IP*" And _
                            Trim(temp_trans.mMODE) = "T" And _
                            Trim(temp_trans.mST) = "0" Then
                                temp_trans.mIsIP = True
                        Else
                            temp_trans.mIsIP = False
                        End If
                            
                        
                        ' new
                        If G_DELAY_FLAG Then
                            ' orange main procedure to recognize if the asn have some delay between eda and port arrival
                            temp_trans.tryToAssignToThisAsnDateArrivalAtThePort CStr(i_mgo.pMS9PO400.SHIP_CMNT(Int(x))), i_mgo
                            ' if mArrivalAtThePort were never assigned clng will throw 0
                            ' Debug.Print CLng(temp_trans.mArrivalAtThePort)
                        End If
                        
                        
                        temp_trans.mOnPop = True
    
                        kolekcja.Add temp_trans
                        
                    End If
                Next x
                
                If (CStr(i_mgo.getString(22, 2, 5)) = "I5487") Or (CStr(i_mgo.getString(22, 2, 5)) = "I6155") Or (CStr(i_mgo.getString(22, 2, 5)) = "I6293") Then
                    Exit Do
                End If
                
                i_mgo.sendKeys "<pf8>"
                is_f8_pressed = True
            Loop While True
        End If
        is_f8_pressed = False
        
        ' ===================================================================
End Sub


Private Function check_if_nsr_issue_appeared(i_mgo As MGO) As Boolean
    check_if_nsr_issue_appeared = False
    
    If CStr(i_mgo.actualScreen) = "MS3P9800" Then
        check_if_nsr_issue_appeared = True
    End If
    
    
End Function
