Attribute VB_Name = "FormatModule"
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

' source will be range that defines classic or blue or purple and any
Public Sub classic_colors(ictrl As IRibbonControl)

    set_on ThisWorkbook.Sheets("Parts").Range("j26:j29"), ThisWorkbook.Sheets("Parts").Range("j15:j18")
    DoEvents
    change_layout_on_already_created_report
End Sub

Public Sub blue_colors(ictrl As IRibbonControl)

    set_on ThisWorkbook.Sheets("Parts").Range("l26:l29"), ThisWorkbook.Sheets("Parts").Range("j15:j18")
    DoEvents
    change_layout_on_already_created_report
End Sub

Public Sub purple_colors(ictrl As IRibbonControl)

    set_on ThisWorkbook.Sheets("Parts").Range("n26:n29"), ThisWorkbook.Sheets("Parts").Range("j15:j18")
    DoEvents
    change_layout_on_already_created_report
End Sub

Public Sub no_colors(ictrl As IRibbonControl)

    set_on ThisWorkbook.Sheets("Parts").Range("p26:p29"), ThisWorkbook.Sheets("Parts").Range("j15:j18")
    DoEvents
    change_layout_on_already_created_report
End Sub

Public Sub ecoprint_colors(ictrl As IRibbonControl)

    set_on ThisWorkbook.Sheets("Parts").Range("r26:r29"), ThisWorkbook.Sheets("Parts").Range("j15:j18")
    DoEvents
    change_layout_on_already_created_report
End Sub

Public Sub custom_colors(ictrl As IRibbonControl)

    set_on ThisWorkbook.Sheets("Parts").Range("t26:t29"), ThisWorkbook.Sheets("Parts").Range("j15:j18")
    DoEvents
    change_layout_on_already_created_report
End Sub

Public Sub set_on(source As Range, dest As Range)

    source.Copy
    dest.PasteSpecial xlPasteAll
End Sub



' this function in easy way checking if active sheet is a coverage report
Private Function check_if_active_worksheet_is_coverage_report() As Boolean
    check_if_active_worksheet_is_coverage_report = False
    

    If ActiveSheet.Range("g2") = "First runout" Then
        If ActiveSheet.Range("h3") = "req" Then
        
            check_if_active_worksheet_is_coverage_report = True
        End If
    End If
End Function


' this sub will change colors on report from layout on parts
' so this is obvious that you need to provide addtional logic before to change it
' if you dont want to make it in manual way
Private Sub change_layout_on_already_created_report()


    ' na dzien dobry pierwszy warunek ktory
    If check_if_active_worksheet_is_coverage_report() Then
    
        Dim l As layout_type
        Dim rng As Range
        Set rng = Range("a2")
        
        If Range("a1") = 0 Then
            l = STD
        ElseIf Range("a1") = 1 Then
            l = GREEN
        ElseIf Range("a1") = 2 Then
            l = GREEN_AND_BLUE
        End If
        
        
        
        Do
            FormatCoverRecord rng, l
            Set rng = rng.offset(7, 0)
        Loop While rng <> ""
    
    
        MsgBox "ready!"
    Else
        MsgBox "you cannot change layout on this worksheet!"
    End If
End Sub


Private Sub init_dzejs(j15 As Range, j16 As Range, j17 As Range, j18 As Range, j19 As Range)

    Set j15 = Sheets("Parts").Range("j15")
    Set j16 = Sheets("Parts").Range("j16")
    Set j17 = Sheets("Parts").Range("j17")
    Set j18 = Sheets("Parts").Range("j18")
    Set j19 = Sheets("Parts").Range("j19")
End Sub

Sub FormatCoverRecord(ByRef rng As Range, layout_type As layout_type)
    ' procedura ta zaczyna sie od pierwszej kolumny i konkretnego pierwszego wiersza iteracji danej czesci
    
    
    
    ' j15 domyslnie szara ramka
    ' j16 ciemny niebieski
    ' j17 jasny niebieski - razem te dwa tworza szachownice danych
    ' j18 kolor pod ebal rozowy domyslnie
    ' j19 bialy z niczym - prawdopodobnie w ogole nie bedzie go w implementacji
    Dim j15 As Range, j16 As Range, j17 As Range, j18 As Range, j19 As Range
    init_dzejs j15, j16, j17, j18, j19
    
    Dim the_layout As ILayout
    Set the_layout = New TheLayout
    
    ' wszystkie dane ogolne z popa
    Dim general_data As Range
    
    Set general_data = Range(rng.offset(1, 0), rng.offset(5, 6))
    ' general_data.Font.Bold = True
    ' the_layout.FillSolidGridLines general_data, RGB(0, 0, 0)
    the_layout.BackColor general_data, j16.Interior.Color
    
    
    ' szachownica
    ' supp name
    the_layout.BackColor rng.offset(2, 1), j17.Interior.Color
    ' empty
    the_layout.BackColor rng.offset(4, 1), j17.Interior.Color
    ' bank
    the_layout.BackColor rng.offset(1, 2), j17.Interior.Color
    the_layout.BackColor rng.offset(1, 3), j17.Interior.Color
    
    ' pcs to go
    the_layout.BackColor rng.offset(3, 2), j17.Interior.Color
    the_layout.BackColor rng.offset(3, 3), j17.Interior.Color
    
    ' misc
    the_layout.BackColor rng.offset(5, 2), j17.Interior.Color
    the_layout.BackColor rng.offset(5, 3), j17.Interior.Color
    
    ' doh
    the_layout.BackColor rng.offset(2, 4), j17.Interior.Color
    the_layout.BackColor rng.offset(2, 5), j17.Interior.Color
    
    ' std pack
    the_layout.BackColor rng.offset(4, 4), j17.Interior.Color
    the_layout.BackColor rng.offset(4, 5), j17.Interior.Color
    
    ' kolumna G
    ' dla dodatkowego first runout
    ' the_layout.BackColor rng.offset(1, 6), j17.Interior.Color
    ' Range(rng.offset(2, 6), rng.offset(5, 6)).Font.Color = j16.Interior.Color
    
    ' pierwszy wierz na szaro
    the_layout.BackColor Range(rng, rng.offset(0, 28)), j15.Interior.Color
    Range(rng, rng.offset(0, 28)).Font.Color = RGB(255, 255, 255)
    rng.offset(0, 7).Font.Color = j15.Interior.Color
    ' czerwony past due
    rng.offset(0, 8).Font.Color = RGB(240, 0, 0)
    Range(rng, rng.offset(0, 28)).Font.Bold = True
    
    ' legenda tez na szaro
    the_layout.BackColor Range(rng.offset(1, 7), rng.offset(5, 7)), j15.Interior.Color
    ' bialy font bez coverage
    Range(rng.offset(1, 7), rng.offset(4, 7)).Font.Color = RGB(255, 255, 255)
    
    
    ' rqm
    the_layout.BackColor Range(rng.offset(1, 8), rng.offset(1, 28)), j16.Interior.Color
    Range(rng.offset(1, 7), rng.offset(1, 28)).Font.Bold = True
    
    ' ebal
    the_layout.BackColor Range(rng.offset(5, 7), rng.offset(5, 28)), j18.Interior.Color
    Range(rng.offset(5, 7), rng.offset(5, 28)).Font.Bold = True
    
    ' komorka z nazwa plantu
    the_layout.ChangeTxtOrientation rng.offset(1, 0), 90
    the_layout.BackColor rng.offset(1, 0), j15.Interior.Color
    rng.offset(1, 0).Font.Color = RGB(255, 255, 255)
    rng.offset(1, 0).Font.Bold = True
    Range(rng.offset(1, 0), rng.offset(5, 0)).Merge
    rng.offset(1, 0).VerticalAlignment = xlCenter
    rng.offset(1, 0).HorizontalAlignment = xlCenter
    
    ' part number bold
    rng.offset(1, 1).Font.Bold = True
    
    
    Set the_layout = Nothing
End Sub



' ta funkcjonalnosc jest bardzo fajna pod wzdledem wylapywania wlasnie gdzie znajduje sie kreska
' ttime'u
' wazne w bankach jest to ze potrzebuje przeliczyc zaraz z nia
Public Sub kreska_ttime_u(ByRef the_layout As ILayout, ByRef rng As Range)

    If Not rng Is Nothing Then
        If rng.Count = 1 Then
    
            If czy_to_jest_arkusz_coverage(rng.Parent) Then
            
                ' kreska ttime'u
                ' ta 5 lub 7 jest tutaj ultra istotna :)
                ' ===================================================
                delta_week = CLng(CLng(rng.offset(4, 5)) / (24# * CDbl(ThisWorkbook.Sheets("register").Range("tc"))))
                
                ' indays
                rng.offset(4, 6) = Application.WorksheetFunction.Round((CDbl(Val(rng.offset(4, 5)) / 24#) / CDbl(ThisWorkbook.Sheets("register").Range("tc"))) * 7, 2)
                ' ===================================================
                
                Dim tt_rng As Range
                Dim i_rng As Range
                Dim bank_rng As Range
                ' plus jeden poniewaz oversiaki licza pierwszy tydzien od nastepnego poniewaz i tak
                ' dostawca muli z wyslaniem
                Set tt_rng = Range(rng.offset(0, 9 + delta_week), rng.offset(5, 9 + delta_week))
                
                ' usuniecie starych ramek
                For x = 1 To 20
                    the_layout.FillThinFrame Range(rng.offset(0, 8 + x), rng.offset(5, 8 + x)), RGB(255, 255, 255)
                Next x
                
                Set bank_rng = rng.offset(1, 3)
                
                
                
                If Trim(CStr(tt_rng.item(1))) <> "" Then
                
                
                    ' sprawdz czy choc jeden coverage przed ttime jest czerwony
                    Dim red_rng As Range
                    Dim red_flag As Boolean
                    
                    If G_CP_CRITICAL Then
                        Set red_rng = Range(rng.offset(5, 9), tt_rng.item(5, 1).offset(0, -1))
                    Else
                        Set red_rng = Range(rng.offset(5, 9), tt_rng.item(5, 1))
                    End If
                    
                    
                    
                    For Each i_rng In red_rng
                    
                        ' Debug.Print i_rng.Address
                        
                        If IsNumeric(i_rng) Then
                            Debug.Print i_rng.Value
                            If i_rng.Value < 0 Then
                                red_flag = True
                            End If
                        End If
                    Next i_rng
                    
                    If red_flag Then
                        the_layout.FillSolidFrame tt_rng, RGB(255, 0, 0)
                    Else
                        the_layout.FillSolidFrame tt_rng, RGB(0, 0, 0)
                    End If
                    
                    ' tt_rng to cala ramka potrzebuje drugiego pola w sumie
                    Dim fst_req_after_ttime As Range
                    Set fst_req_after_ttime = tt_rng.item(2).offset(0, 1)
                    
                    
                    Dim srednia_rqmow As Double, proporcja As Double
                    
                    srednia_rqmow = CDbl(CDbl(CLng(fst_req_after_ttime) + _
                        CLng(fst_req_after_ttime.offset(0, 1)) + _
                        CLng(fst_req_after_ttime.offset(0, 2))) / 3#)
                        
                    ' 1. teraz jesli wartosc jest ujemna to znaczy ze srednia jest wieksza niz bank
                    ' 2. jesli wartosc jest dodatnia to znaczybank jest wiekszy
                    ' samo liczenia to proste odejmowanie plus przemnozenie przez propocje zeby zawsze sie miescilo w zbiorze
                    ' 1 0 -1
                    proporcja = 1 - CDbl(CDbl(srednia_rqmow) / CDbl(bank_rng))
                    
                    If proporcja < -1 Then
                        bank_rng.Interior.Color = ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I2").Interior.Color
                    ElseIf proporcja >= -1 And propocja <= 1 Then
                    
                        For w = 2 To 21
                            Debug.Print ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I" & CStr(w)).Value & ", " & ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I" & CStr(w + 1)).Value
                            If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I" & CStr(w)).Value <= proporcja Then
                                If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I" & CStr(w + 1)).Value >= proporcja Then
                                    jakikolor = CDbl(ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I" & CStr(w)).Interior.Color)
                                    bank_rng.Interior.Color = jakikolor
                                    Exit For
                                End If
                            End If
                        Next w
                    ElseIf proporcja > 1 Then
                        bank_rng.Interior.Color = ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("I22").Interior.Color
                    End If
                    
                    
                End If
            
            
            End If
        End If
    End If
End Sub

Public Function znajdz_pn_coverage_dla(ByRef r As Range) As Range
    Set znajdz_pn_coverage_dla = Nothing
    
    If r.Count = 1 Then
    
        ile = (r.row Mod 7) - 2
        
        Set znajdz_pn_coverage_dla = r.Parent.Cells(r.offset(0 - ile, 0).row, 1)
    End If
End Function

Private Function czy_to_jest_arkusz_coverage(Sh As Worksheet) As Boolean

    czy_to_jest_arkusz_coverage = False
    
    ' If sh.Cells(2, 1) = "PLT" Then
        If Sh.Cells(2, 7) = "First runout" Then
            czy_to_jest_arkusz_coverage = True
        End If
    ' End If
End Function

Sub Autoformatowanie()

    ActiveCell.offset(5, 9).Range("A1:S1").Select
    'Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=J14<K10"
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=WK<W[-4]K[1]"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=WK>SUMA(W[-4]K[1]:W[-4]K[3])"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 15773696
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    ActiveCell.offset(-5, -9).Range("A1").Select
End Sub


Public Sub przelicz_arkusz(the_layout As layout_type, Optional fst_time As Boolean)

    Dim rng As Range
    Dim ir As Range
    
    Set rng = ThisWorkbook.ActiveSheet.Range("h1")
    Do
        If rng Like "*coverage*" Then
            ' tutaj jestesmy w wierszu ktorym mamy formuly na coveregu :)
            Set ir = rng.offset(0, 2)
            
            
            ' tutja blad sie pojawil na wersji 4.1 szczegolnie na plancie zc kiedy to cbal
            ' nie wystepuja pobierajac z mgo tylko sama pustke :D
            ' zobaczymy czy ten dodatkowy warunek rozwala ten problem
            If rng.offset(-4, -2) = "" Then
                rng.offset(-4, -2) = "0"
            End If
            
            
            If Not IsError(ir) Then
                ' iteracja po kolei kolumny formul
                If fst_time Then
                    Do
                    
                        If the_layout = STD Then
                            ir.NumberFormat = "0_ ;[Red]-0 "
                        Else
                        
                            ir.NumberFormat = "0_ ;[Red]-0 "
                        
                            If CLng(ir) >= 0 Then
                                If CLng(ir) < CLng(ir.offset(-4, 1)) Then
                                    ir.Font.Color = RGB(0, 255, 0)
                                ElseIf (CLng(ir) > CLng(ir.offset(-4, 1)) + CLng(ir.offset(-4, 2)) + CLng(ir.offset(-4, 3))) And (the_layout = GREEN_AND_BLUE) Then
                                    ir.Font.Color = RGB(0, 0, 255)
                                Else
                                    ir.Font.Color = RGB(0, 0, 0)
                                End If
                                
                            End If
                        End If
                        Set ir = ir.offset(0, 1)
                    Loop While ir <> ""
                Else
                
                
                    If the_layout <> STD Then
                        Do
                    
                        
                            ir.NumberFormat = "0_ ;[Red]-0 "
                        
                            If CLng(ir) >= 0 Then
                                If CLng(ir) < CLng(ir.offset(-4, 1)) Then
                                    ir.Font.Color = RGB(0, 255, 0)
                                ElseIf (CLng(ir) > CLng(ir.offset(-4, 1)) + CLng(ir.offset(-4, 2)) + CLng(ir.offset(-4, 3))) And (the_layout = GREEN_AND_BLUE) Then
                                    ir.Font.Color = RGB(0, 0, 255)
                                Else
                                    ir.Font.Color = RGB(0, 0, 0)
                                End If
                                
                            End If
                        
                            Set ir = ir.offset(0, 1)
                        
                        Loop While ir <> ""
                    End If
                    
                End If
            End If
            
        End If
        
        Set rng = rng.offset(1, 0)
        If rng = "" Then
            Set rng = rng.offset(1, 0)
            Set rng = rng.offset(1, 0)
        End If
    Loop While rng <> ""
End Sub


Public Sub show_me()
    
    ' j15 domyslnie szara ramka
    ' j16 ciemny niebieski
    ' j17 jasny niebieski - razem te dwa tworza szachownice danych
    ' j18 kolor pod ebal rozowy domyslnie
    ' j19 bialy z niczym - prawdopodobnie w ogole nie bedzie go w implementacji
    Dim j15 As Range, j16 As Range, j17 As Range, j18 As Range, j19 As Range
    Set j15 = Sheets("Parts").Range("j15")
    Set j16 = Sheets("Parts").Range("j16")
    Set j17 = Sheets("Parts").Range("j17")
    Set j18 = Sheets("Parts").Range("j18")
    Set j19 = Sheets("Parts").Range("j19")
    
    
    ' grey
    switch_kolor Range("J7"), j15
    switch_kolor Range("J6:U6"), j15
    switch_kolor Range("Q7:Q10"), j15
    
    ' rozowy
    switch_kolor Range("Q11:U11"), j18
    
    'dark
    
    ' pasek req
    switch_kolor Range("R7:U7"), j16
    
    switch_kolor Range("K7"), j16
    switch_kolor Range("K9"), j16
    switch_kolor Range("K11"), j16
    
    switch_kolor Range("L8:M8"), j16
    switch_kolor Range("L10:M10"), j16
    
    switch_kolor Range("N7:O7"), j16
    switch_kolor Range("N9:O9"), j16
    switch_kolor Range("N11:O11"), j16
    
    switch_kolor Range("p7:p11"), j16
    
    
    'bright
    switch_kolor Range("K8"), j17
    switch_kolor Range("K10"), j17
    
    switch_kolor Range("L7:M7"), j17
    switch_kolor Range("L9:M9"), j17
    switch_kolor Range("L11:M11"), j17
    
    switch_kolor Range("N8:O8"), j17
    switch_kolor Range("N10:O10"), j17
End Sub

Private Sub switch_kolor(ByRef ch As Range, ByRef you As Range)
    ch.Interior.Color = you.Interior.Color
End Sub




