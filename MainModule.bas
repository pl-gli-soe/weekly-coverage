Attribute VB_Name = "MainModule"
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



' the main module
' have not only the main procedures
' but also subs from the copy group
' where you can find a copy critical
' a just copy
' and my favourite copy with runout order
' from version 4.21 I'm try to implement coping with duns or even container order
' need to think how to use the connections between flat table and std report on coverage.


Public Sub bank_formatting_on(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("BANKI") = 1
    MsgBox "Bank Formatting is ON"
End Sub

Public Sub bank_formatting_off(ictrl As IRibbonControl)
    ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("BANKI") = 0
    MsgBox "Bank Formatting is OFF"
End Sub

Public Sub run_simple2()
    InitForm.OptionButtonSTD.Value = True
    InitForm.OptionButtonGREEN.Value = False
    InitForm.OptionButtonGB.Value = False
    
    InitForm.CheckBoxCheckStatuses = True
    
    
    If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("BANKI") = 0 Then
        InitForm.CheckBoxBANKI = False
    Else
        InitForm.CheckBoxBANKI = True
    End If
    
    InitForm.Show
    InitForm.Repaint
End Sub


Public Sub run_simple(ictrl As IRibbonControl)
    run_simple2
End Sub


' obsolete from 4.4x
'
'Public Sub run_green(ictrl As IRibbonControl)
'    InitForm.OptionButtonSTD.Value = False
'    InitForm.OptionButtonGREEN.Value = True
'    InitForm.OptionButtonGB.Value = False
'
'    InitForm.CheckBoxCheckStatuses = True
'
'    InitForm.Show
'    InitForm.Repaint
'End Sub
'
'Public Sub run_green_and_blue(ictrl As IRibbonControl)
'
'    InitForm.OptionButtonSTD.Value = False
'    InitForm.OptionButtonGREEN.Value = False
'    InitForm.OptionButtonGB.Value = True
'
'    InitForm.CheckBoxCheckStatuses = True
'
'    InitForm.Show
'    InitForm.Repaint
'End Sub

Private Function check_sheets(s As String) As String
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Sheets
        If Sh.Name = CStr(s) Then
            check_sheets = check_sheets(s & "I")
        Else
            check_sheets = s
            Exit For
        End If
    Next Sh
End Function


Public Sub copy_frames(ictrl As IRibbonControl)
    CopyFramesForm.Show
End Sub

Public Sub copy_red_frames()

    Application.EnableEvents = False
    ' ThisWorkbook.Sheets("register").Range("autocopy") = 1

    Dim rng As Range
    Set rng = ActiveSheet.Range("i2")
    
    If rng <> "Past due" Then
        MsgBox "to nie jest kowerydz!"
        Exit Sub
    End If
    
    
    Dim critical_sheet As Worksheet
    Set critical_sheet = Sheets.Add
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    Dim critical_rng As Range
    Set critical_rng = critical_sheet.Range("a2")

    critical_sheet.Name = check_sheets(CStr("c_" & rng.Parent.Name))
    
    
    Do
    
        i = 0
        Do
            If rng.offset(0, i).Borders(xlEdgeTop).LineStyle = xlContinuous _
            And rng.offset(0, i).Borders(xlEdgeTop).Color = RGB(255, 0, 0) Then
                copy_this_range rng, critical_rng
                Exit Do
            ElseIf rng.offset(0, i).Value = "" Then
                Exit Do
            End If
            i = i + 1
        Loop While True
    
        Set rng = rng.offset(7, 0)
    Loop While rng <> ""
    
    Columns("A:AC").EntireColumn.AutoFit
    'Columns("C:D").Select
    'Selection.EntireColumn.Hidden = True
    'Columns("G:G").Select
    'Selection.EntireColumn.Hidden = True
    Range("A1").Select
    
    Application.EnableEvents = True
End Sub

Public Sub copy_misc_qhd()


    ' this orange is 240, 120, 20

    Application.EnableEvents = False
    ' ThisWorkbook.Sheets("register").Range("autocopy") = 1

    Dim rng As Range
    Set rng = ActiveSheet.Range("i2")
    
    If rng <> "Past due" Then
        MsgBox "to nie jest kowerydz!"
        Exit Sub
    End If
    
    
    Dim critical_sheet As Worksheet
    Set critical_sheet = Sheets.Add
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    Dim critical_rng As Range
    Set critical_rng = critical_sheet.Range("a2")

    critical_sheet.Name = check_sheets(CStr("miscqhd_" & rng.Parent.Name))
    
    
    Do
    
        i = 0
        Do
            If miscqhd(rng) Then
                copy_this_range rng, critical_rng
                Exit Do
            ElseIf rng.offset(0, i).Value = "" Then
                Exit Do
            End If
            i = i + 1
        Loop While True
    
        Set rng = rng.offset(7, 0)
    Loop While rng <> ""
    
    Columns("A:AC").EntireColumn.AutoFit
    'Columns("C:D").Select
    'Selection.EntireColumn.Hidden = True
    'Columns("G:G").Select
    'Selection.EntireColumn.Hidden = True
    Range("A1").Select
    
    Application.EnableEvents = True
End Sub



Public Sub copy_orange()


    ' this orange is 240, 120, 20

    Application.EnableEvents = False
    ' ThisWorkbook.Sheets("register").Range("autocopy") = 1

    Dim rng As Range
    Set rng = ActiveSheet.Range("i2")
    
    If rng <> "Past due" Then
        MsgBox "to nie jest kowerydz!"
        Exit Sub
    End If
    
    
    Dim critical_sheet As Worksheet
    Set critical_sheet = Sheets.Add
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    Dim critical_rng As Range
    Set critical_rng = critical_sheet.Range("a2")

    critical_sheet.Name = check_sheets(CStr("orange_" & rng.Parent.Name))
    
    
    Do
    
        i = 0
        Do
            If orange(rng) Then
                copy_this_range rng, critical_rng
                Exit Do
            ElseIf rng.offset(0, i).Value = "" Then
                Exit Do
            End If
            i = i + 1
        Loop While True
    
        Set rng = rng.offset(7, 0)
    Loop While rng <> ""
    
    Columns("A:AC").EntireColumn.AutoFit
    'Columns("C:D").Select
    'Selection.EntireColumn.Hidden = True
    'Columns("G:G").Select
    'Selection.EntireColumn.Hidden = True
    Range("A1").Select
    
    Application.EnableEvents = True
End Sub

Private Function miscqhd(ByRef r As Range) As Boolean
    miscqhd = False
    Dim tmp As Range, ir As Range
    Set tmp = r.Parent.Range(r.offset(2, -6), r.offset(5, -3))
    
    
    For Each ir In tmp
        If (ir = "MISC" And ir.offset(0, 1) > 0) Or (ir = "QHD" And ir.offset(0, 1) > 0) Then
            miscqhd = True
            Exit Function
        End If
    Next ir
End Function

Private Function orange(ByRef r As Range) As Boolean
    orange = False
    Dim tmp As Range, ir As Range
    Set tmp = r.Parent.Range(r.offset(1, 0), r.offset(5, 20))
    
    
    For Each ir In tmp
        If ir.Font.Color = RGB(240, 120, 20) Then
            orange = True
            Exit Function
        End If
    Next ir
End Function

Public Sub copy_with_runout_order()


    ' wlasciwie implementacja orderu na runoucie niczym sie nie rozni od copy critical
    ' w sumie to wlasnie temu drugiemu to zawdzieczamy poniewaz jego kod jest bardzo uogolniony tak ze
    ' wieksza jego czesc moze zostac uzyta ponownie
    ' na dzien 3 lipca 2014 nie ma jeszcze zbyt dobrze zrobionej zasady DRY chocby przy sprawdzeniu czy w ogole dane
    ' kopiowanie jest realne i ze naszym aktualnie aktywnym arkuszem jest arkusz raportu coverage'a
    ' no ale z drugiej strony nie kopiuje 500 linii kodu wiec chyba mimo wsyzstko moge sobie darowac
    ' z reszta zobacze jeszcze...
    ' ==========================================
    

    Dim coll As Collection
    Dim ohi As OrderHandlerItem
    

    Application.EnableEvents = False

    Dim rng As Range
    Set rng = ActiveSheet.Range("i2")
    Dim Sh As Worksheet
    Set Sh = ActiveSheet
    
    If rng <> "Past due" Then
        MsgBox "to nie jest kowerydz!"
        Exit Sub
    End If
    
    
    Dim critical_sheet As Worksheet
    Set critical_sheet = Sheets.Add
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    Dim critical_rng As Range
    Set critical_rng = critical_sheet.Range("a2")
    critical_sheet.Name = check_sheets(CStr("order_" & Sh.Name))
    
    
    Set coll = Nothing
    Set coll = New Collection
    
    i = 1
    Do
    
        Set ohi = Nothing
        Set ohi = New OrderHandlerItem
        Set ohi.kotiwca = rng
        Set ohi.runout = rng.offset(1, -2)
        ohi.indeks = i
        coll.Add ohi
        
        i = i + 1
        Set rng = rng.offset(7, 0)
    Loop While rng <> ""
    
    
    
    ' ta metoda szczegolnie mi sie podoba poniewaz dzieki niej nie musismy zbytnio sie przjemowac jak przemionaowac kolejnosc
    ' arkuszy kod jest tak przjerzysty jak jezioro bajkal 2000 lat temu na srodku gdy na lodce patrzac na dno ludzie z lekiem wysokosci
    ' mogliby miec problem :D:D:D.
    ' no po prostu zwykle sortowanie "bomblami"
    ' az geba sama sie usmiecha ze w tak prosty sposob mozna danymi zarzadzac
    bubble coll
    
    'For Each ohi In coll
    '    Debug.Print ohi.kotiwca.Address & " " & ohi.runout
    'Next ohi
    
    For Each ohi In coll
        copy_this_range ohi.kotiwca, critical_rng
    Next ohi
    
    Columns("A:AC").EntireColumn.AutoFit
    'Columns("C:D").Select
    'Selection.EntireColumn.Hidden = True
    'Columns("G:G").Select
    'Selection.EntireColumn.Hidden = True
    Range("A1").Select
    
    Application.EnableEvents = True
End Sub


Public Sub copy_asn_offset()


    ' this orange is 240, 120, 20

    Application.EnableEvents = False
    ' ThisWorkbook.Sheets("register").Range("autocopy") = 1

    Dim rng As Range
    Set rng = ActiveSheet.Range("i2")
    
    If rng <> "Past due" Then
        MsgBox "to nie jest kowerydz!"
        Exit Sub
    End If
    
    
    Dim new_sheet As Worksheet
    Set new_sheet = Sheets.Add
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    Dim n_rng As Range
    Set n_rng = new_sheet.Range("a2")

    On Error Resume Next
    new_sheet.Name = check_sheets(CStr("asnofst_" & rng.Parent.Name))
    
    
    Do
    
        i = 0
        Do
            If Not isempty(rng) Then
                copy_this_range rng, n_rng
                make_on_this_copy_asn_offset n_rng
                Exit Do
            ElseIf rng.offset(0, i).Value = "" Then
                Exit Do
            End If
            i = i + 1
        Loop While True
    
        Set rng = rng.offset(7, 0)
    Loop While rng <> ""
    
    Columns("A:AZ").EntireColumn.AutoFit
    'Columns("C:D").Select
    'Selection.EntireColumn.Hidden = True
    'Columns("G:G").Select
    'Selection.EntireColumn.Hidden = True
    Range("A1").Select
    
    Application.EnableEvents = True
End Sub


Private Function isempty(r As Range) As Boolean
    If Trim(r) = "" Then
        isempty = True
    Else
        isempty = False
    End If
End Function

Private Sub make_on_this_copy_asn_offset(r As Range)

    Dim tmprng As Range
    
    ' MsgBox r.Address -> to jest A9
    
    Set tmprng = r.End(xlUp).End(xlUp).End(xlToRight).End(xlToRight).End(xlToRight).offset(2, 0)
    
    ' 10 kolumna zaczynaja sie pierwsze regularne asny
    pierwsza_kolumna_asnow = 10
    
    Do
        
        ' =======================================
        
        For x = 0 To 2
            If CStr(tmprng.offset(x, 0)) <> "" Then
                On Error Resume Next
                If CLng(tmprng.offset(x, 0)) > 0 Then
                    tmprng.offset(x, 1) = tmprng.offset(x, 0)
                    tmprng.offset(x, 0) = "0"
                    'tmprng.offset(x, 0).Interior.Color = RGB(20, 200, 220)
                    'tmprng.offset(x, 1).Interior.Color = RGB(75, 160, 200)
                    With ThisWorkbook.Sheets("Parts")
                        ' colors
                        tmprng.offset(x, 0).Interior.Color = .Range("Q16").Interior.Color
                        tmprng.offset(x, 1).Interior.Color = .Range("R16").Interior.Color
                    End With
                    
                    ' move also comment
                    tmprng.offset(x, 1).ClearComments
                    tmprng.offset(x, 1).AddComment CStr(tmprng.offset(x, 0).Comment.Text)
                    tmprng.offset(x, 0).ClearComments
                End If
            End If
        Next x
        
        ' =======================================
    
        Set tmprng = tmprng.offset(0, -1)
    Loop Until Int(tmprng.Column) < Int(pierwsza_kolumna_asnow)

End Sub

Public Sub just_copy_this_sheet()


    Dim Sh As Worksheet
    Application.EnableEvents = False
    Set Sh = ThisWorkbook.ActiveSheet
    Sh.Copy After:=ThisWorkbook.Sheets(Sheets.Count)
    ActiveSheet.Name = "copy_" & CStr(Sh.Name)
    
    Application.EnableEvents = True
    
End Sub

Private Sub bubble(c As Collection)


    Dim tmp_c As Collection
    Set tmp_c = New Collection
    
    Dim ohi1 As OrderHandlerItem, ohi2 As OrderHandlerItem, tmp As OrderHandlerItem
    For x = 1 To c.Count - 1
        
        If Int(c.item(x).runout) > Int(c.item(x + 1).runout) Then
            swap_elements c.item(x), c.item(x + 1)
            x = 0
        End If

        
    Next x
End Sub


Private Sub swap_elements(ByRef ohi1 As OrderHandlerItem, ByRef ohi2 As OrderHandlerItem)


    ' ten prywatny sub wlasciwie pracuje tylko dla metody bubble
    ' swapy to dosyc sliski temat jesli chodzi o jezyki pokroju vba
    ' meczace sa zasady ktore nie zawsze obowiazuje i wszech obecne wyjatki
    ' az boli czasami ze glowni programisci windowsa
    ' na takie rzeczy pozwalaja
    ' jezyk powinien byc spojny i na slepo jesli jest jakas idea
    ' to nie musze znac miliona konstrukcji skladniowych
    ' wszystko powinno dzialac tak jak mysle bez wiekszych rozterek
    ' ze zmienna typu range ma mase wyjatkow w stosunku do innych obiektow
    ' i ze nawet byref nie pomoze i - nie wiem - moze trzeba patrzec na range z perspektywy
    ' bycia elementem static czy cos...
    ' eh :D
    
    Dim tmp_k As Range, tmp_r As Range
    
    tmp_i = ohi1.indeks
    Set tmp_r = ohi1.runout
    Set tmp_k = ohi1.kotiwca
    
    
    ohi1.indeks = ohi2.indeks
    Set ohi1.kotiwca = ohi2.kotiwca
    Set ohi1.runout = ohi2.runout
    
    ohi2.indeks = tmp_i
    Set ohi2.kotiwca = tmp_k
    Set ohi2.runout = tmp_r ' temporary runout

    
End Sub


Private Sub copy_this_range(ByRef rng As Range, ByRef dest As Range)
    Range(rng.offset(0, -8), rng.offset(5, 20)).Copy dest
    
    
    Set dest = dest.offset(7, 0)
End Sub

