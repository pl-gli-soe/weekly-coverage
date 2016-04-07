Attribute VB_Name = "ScheduleModule"
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


' tutaj nieco odmiany w kodzie InputBox zostapilem zupelnie customwym
' formem
'
' wykorzystalem tutaj nieco inna strategie dzialania
' gdzie wystpuje jezcze logika po kliknieciu w submit
' co wiecej submit w sumie nie robi nic poza akcja Hide :P
' cala reszta znajduje sie bezposrednio w subach podlaczonych do ribbona (czyli tutaj w tym module)

' 2nd args is for add schedule or add schedules
' public function MyInputBox("How many schedules you want?", "1 or 2", "1")
Public Sub initMyInputBox(q As String, nm As String, qty As Long)
    
    ' it make sense
    If qty > 0 Then
    
        With AddSchedulesForm
            .nm = CStr(nm)
            .TextBoxQty = CStr(qty)
            .LabelQ.Caption = CStr(q)
            .CheckBoxInclCurrWeek.Value = True
            .Show
        End With
        
    
    Else
        MsgBox "only positive integers!"
    End If
End Sub

Public Sub add_schedule(ictrl As IRibbonControl)

    Dim y As Integer
    Dim s As String
    initMyInputBox "How many schedules you want?", "1", "1"
    
    With AddSchedulesForm
        s = .TextBoxQty
    End With
    If s = "" Then
        Exit Sub
    Else
        y = Int(s)
    End If
    
    
    If y > 7 Then
        MsgBox "you can have 7 as max! Click OK and work with add_sch = 7"
        y = 7
    End If

    Application.EnableEvents = False

    Dim the_layout As ILayout
    Set the_layout = New TheLayout

    Dim rng As Range
    Set rng = Range("i2")
    
    ' selected range
    Dim sel_r As Range
    Set sel_r = Selection
    Set sel_r = sel_r(1, 1)
    
    While rng.row + 5 <= sel_r.row
        Set rng = rng.offset(7, 0)
    Wend
    
    If CStr(rng) <> "Past due" Then
        Exit Sub
    End If
    
    mgoInit
    
    proc_on_screen rng, y, the_layout
    
    
    Set the_layout = Nothing
    Application.EnableEvents = True
    
    MsgBox "adding schedule for PN ready!"

End Sub

Public Sub add_schedules(ictrl As IRibbonControl)

    

    Dim y As Integer
    Dim s As String
    initMyInputBox "How many schedules you want?", "2", "1"
    
    
    With AddSchedulesForm
        s = .TextBoxQty
    End With
    
    If s = "" Then
        Exit Sub
    Else
        y = Int(s)
    End If
    
    
    If y > 7 Then
        MsgBox "you can have 7 as max! Click OK and work with add_sch = 7"
        y = 7
    End If

    Application.EnableEvents = False

    Dim the_layout As ILayout
    Set the_layout = New TheLayout

    Dim rng As Range
    Set rng = Range("i2")
    
    If CStr(rng) <> "Past due" Then
        Exit Sub
    End If
    
    mgoInit
    
    Do
        proc_on_screen rng, y, the_layout
        Set rng = rng.offset(7, 0)
    Loop While rng <> ""
    
    
    Set the_layout = Nothing
    Application.EnableEvents = True
    
    MsgBox "adding schedules ready!"
End Sub


Private Sub proc_on_screen(ByRef rng As Range, y As Integer, ByRef the_layout As ILayout)

    ' each past due on select
    rng.Select
        
    Sess0.Screen.sendKeys ("<Clear>")
    waitForMgo
    If rng.offset(0, -1).Value = "C" Then
        Sess0.Screen.sendKeys ("MYSP1F00 <Enter>")
    Else
        Sess0.Screen.sendKeys ("ZK7PWPSC <Enter>")
    End If
    
    waitForMgo
    
    If rng.offset(0, -1).Value = "C" Then
        Sess0.Screen.putString rng.offset(1, -8), 1, 21
        Sess0.Screen.putString rng.offset(1, -7), 4, 22
        Sess0.Screen.putString rng.offset(3, -7), 4, 49
    Else
        Sess0.Screen.putString rng.offset(1, -8), 3, 8
        Sess0.Screen.putString rng.offset(1, -7), 3, 21
        Sess0.Screen.putString rng.offset(3, -7), 3, 48
    End If
    Sess0.Screen.sendKeys ("<Enter>")
    waitForMgo
    
    delta_week = CLng(CLng(rng.offset(4, -3)) / (24# * CDbl(ThisWorkbook.Sheets("register").Range("tc"))))
    
    Dim pocz As Integer
    If rng.offset(0, -1).Value = "C" Then
        pocz = 12
    Else
        pocz = 11
    End If
    
    
    
    For x = pocz To pocz + (y - 1)
    
        If (Not AddSchedulesForm.CheckBoxInclCurrWeek) And (x = pocz) Then
            ' nope - no current week on coverage
        Else
        
            rng.offset(4, 1 + delta_week + (x - pocz)) = Val(Sess0.Screen.getString(x, 5, 9))
            rng.offset(4, 1 + delta_week + (x - pocz)).Font.Bold = True
            the_layout.BackColor rng.offset(4, 1 + delta_week + (x - pocz)), RGB(210, 110, 110)
        End If
    Next x
    
    
End Sub
