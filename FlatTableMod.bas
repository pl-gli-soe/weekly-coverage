Attribute VB_Name = "FlatTableMod"
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


' Hey this is a test!
Private Sub getTransitFromCmntTest()

    Dim k As Komentarz, coll As Collection
    Set k = New Komentarz
    
    Set coll = Nothing
    Set coll = New Collection
    Set coll = k.pobierz_dane(ActiveCell)
    
    ' Debug.Print coll.Count
    
    Set k = Nothing
    Set coll = Nothing
End Sub


' connect with the ribbon
Public Sub create_flat_table(ictrl As IRibbonControl)

    ThisWorkbook.Sheets("register").Range("togglehandler") = 3
    getTransitFromComment
    
End Sub

Private Function check_if_already_exists() As Boolean
    Dim s As Worksheet
    For Each s In ThisWorkbook.Sheets
        If check_labels(s) Then
            check_if_already_exists = True
            Exit Function
        End If
    Next s
End Function


' generate a proxy - czyli tak wlasciwie stworz flat table
Public Sub getTransitFromComment()


    
    Application.EnableEvents = False
    
    Dim source_sheet As Worksheet
    Set source_sheet = ThisWorkbook.ActiveSheet
    
    
    ' kolekcja wszystkich asnow z calego arkusza
    ' potem stworzymy swoista liste i na jej podstawie damy rade zrobic pivota
    ' little coll to wciaz tylko kolekcja czystego obiektu typu asn co za tym idzie nie ma danych zwiazanych
    ' ze szczegolami bardziej zwiazanymi z sama czescia
    ' dopiero tutaj na wysokosci dodawania elementow do kolekcji coll
    ' mamy mozliowosc rozszerzenia jej komponenetow
    
    
    ' ogolnie implementacja jest dosyc plaska poniewaz nie uwzglednia zadnych klas
    ' z jednej strony bardzo slabo
    ' ale z drugiej gdy bylem w ZA i od czasu do czasu popelnialem babola momentalnie widzialem gdzie on sie znajduje
    ' a nie tak jak mam to w ff gdzie wszystko zatrzymuje sie na wywolaniu pierwszej instacji DailyItem i plakac mi sie chce
    ' gdyz zanim znajde odp na moje pyt musze przejsc przez 4 poziomy obiektow zagniezdzonych pelno metod side'owych
    ' a tutaj wszystko jak na dloni - bo plaskie
    ' wiec wiadomo - szybko i sprawnie przyszlo mi to ostrojstwo napisac
    ' jednak koniec koncow jest trudne do pilnowania
    ' a sama ta metoda jest dosyc rozleglym spaghetti :(
    ' wiec zeby nieco rozjasnic dzialanie
    ' po krotce wyjasnie o co tutaj kaman!
    '
    '
    ' 1
    ' loop_to_fill_td -> glowny motyw (Set little_coll = k.pobierz_dane(r))
    ' gdzie jest to pobieranie danych z komentarza i wstawianie do obiektu o nazwie little comment :D
    ' w srodku tej metody znajduje sie petla ktora przechodzi przez wszystkie dane coverage stad taka zwiezlosc kodu
    ' maly paradoks zwiezlosc na niezwiezlym kodzie, ktory jest w formie spaghetti :D
    '
    ' 2 teraz mozna dodac nowy arkusz gdy little coll jest zapelniony danymi
    ' oraz nadanie mus nazwy z prefixem FLAT
    '
    '
    ' 3
    ' fill labels to dosyc prosta metoda ktora nie robi wlasciwie nic innego tylko wypisuje po kolei nazwy kolumn
    ' jednak z racji tego ze sporo miejsca zajmuje wydzielilem ja jako osobna metode co by ladniej i przejrzysciej
    ' wygladala implementacja getTransitFromComment
    '
    '
    ' 4
    ' fill_sheet_from_coll_el
    ' dosyc prosta i przejsrzysta petla
    ' ktora iteruje zawartosc malej kolekcji i zwraca odpowiednio jej elementy w odpowiednie kolumny flat table
    ' tutaj wazna sprawa ze sa to tylko elementy statyczne - czyli nie uwzgleniamy wszystkie formul i dynamicznych wartosci
    '
    
    ' 4.1
    ' prepare_list_of_containers_in_register -> czyli kopiujemy dane containers
    ' takie rozwiazanie powstalo wczoraj tj poczatek lipca 2014
    ' z racji tego ze ta plaska implemetacja wisi juz miesiaca a nagle mi sie zachcialo dodowac scenariusze z usuwaniem poszczegolnych
    ' kontenerow z racji bytu dla coverage musialem wsadzic nieco wczesniej liste unikalnych kontenerow do workshsetu register
    ' ponieaz na dzien dobry przeliczenia wstepna wyrzucaly blad na pozostalych formulach na flat table
    ' i jesli zrobic to inaczej bylo by bardzo na okolo
    ' zatem stad te dziwne 4.1
    ' najlepsze jest to ze ten sub jest w ogole poza implementacja tego modulu i trzeba go szuakc w CatchCONTModule :P
    ' i co jeszcze lepiej ten sub teoretycznie powinien byc private poniewaz rowniez jest komponentem dla urochomienia formualrza wyrzucania niechcianych
    ' kontenerow w symualcji lancucha dostaw z perspektywy pokrycia na rep coveage.
    ' =====================================================================================
    ' mamy 8 lipca i przestalo mi sie podobac to roziwazanie poniewaz zbyt mocno ingeruje na zewnatrz w arkusz register
    ' ogolnie masakra zbyt wiele do udzwigniecia i algortym przestaje byc ladny - lepiej bedzie gdy wrzuce po prostu nowa kolumne
    ' ogolnie pewnie zrzuce za kolumne URGENCY dodatkowy parametr sterujacy TOGGLE :)
    '' =====================================================================================
    ' koniec koncow nawet nie wiem czy chce dalej utrzymywac ta czesc kodu przy zyciu - poniewaz jesli dam dodatkowa kolumne
    ' ona naturalnie sobie poradzi i wejdzie w tlo w flat table dokladnie tak samo jak reszta poniewaz jest to statyczna wartosc bez historii a sterowania i trzeba inaczej
    ' zrealizowac
    '' =====================================================================================
    '
    '
    ' 5
    ' fill_coverage_with_formulas -> args -> sh, komorka_kowerydza, rep_sh
    ' gdzie:
    ' sh - to jest arkusz flat table
    ' rep_sh to coverage worksheet
    ' to jest fajna metoda ponieaz ingeruje w uklad samego coverage a nie flat table
    ' jest to przyczynek do zagnienia crossed formulas :)
    ' formula rowniez byla na poczatku bardzo toporna to znaczy jesli nie chcacy usunalem flat table
    ' polaczenia formul ginely i coverage stawal sie bezuzyteczny z duza iloscia b³êdów typu #ARG
    ' jednak kolejne odslony tej metody pozwolily sprytnie pobudzac do zycia coverage file
    ' any time tj. nawet jesli FLAT table zostal zniszczony a od ktorego podczas tworzenia coverage
    ' jest uzalezniony bezpowrotnie formulami odniesienia wartosci ASNow
    ' przelicza po prostu wartosci z komentarza nadajac z powrotem uzytecznosci coverage worksheet
    '
    '
    '
    '
    ' 6
    ' copy_s_to_t_as_values sh, source, place i put_formula_into_toggle_column sh, 4
    ' na przemien do przekopiowywania scenariuszy
    ' i najpierw copy_s_to_t_as_values
    ' te metoda zostaje uruchomiona kilka razy poniewaz
    ' bedzie kopiowac informacje z formuly w kolumnie S do miejsc od klkumny U czyli statyczne wartosci
    ' RUNOUT vs EDA CW jednak sama metoda jest bardzo ogolna i mozna ja wykorzystywac wlasciwie wszedzie
    ' jej poziom abstrakcji jest na bardzo wysokim poziomie co za tym idzie moze byc jedna z glownych metod
    ' takich jak chocby last row - proste kopiowanie z srouce to place calych kolumn jako wartosci
    '
    ' druga metoda czyli praca na toggle jest to podmiana kolejnych wartosci dla formuly znajdujacej sie w kolumnie toggle
    ' odpowiaada ona za to jak asny pracuja tj. jako aktywne czy tez nie
    ' odpowiednio zeruja wybrane elementy ukladanki flat table
    ' sam schemat wyglada nastepujaco:
    '
    ' a) copy_s_to_t_as_values -> kopiowanie wszystkich asnow -> kolumna U jako dane
    ' b) put_formula_into_toggle_column -> usuwamy asny ze statusem mniejszym niz 4
    ' c) copy_s_to_t_as_values -> kopiowanie asnow a raczej runout vs eda -> kolumna W jako dane czyli o dwa dalej
    ' d) put_formula_into... -> usuwamy asny ze statusem mniejszym niz 3
    ' e) copy... i analogicznie kopiujemy wartosci do kol V
    '
    ' no i teraz maly trick poniewaz sprawunia jest maly z eda - runout dla kolumny x czyli patrzac z perspektywy tylko i wylacznie DOH
    ' jest to robione na wysokosci fill_sheet_from_coll_el
    ' czyli jest punkt 4!
    ' o wiele wczesniej!
    ' niestety nazwa tej metody niefortunie nic o tym nie mowi i prawie sam bym tego nie zobaczyl
    
    ' sama idea robienia tego w ten sposob wynikalo raczej z wykorzystania okazji ze przez jakis czas w tamtym rejonie kodu
    ' toggle wszystkie sa na zero!
    ' stad tak a nie inaczej.
    '
    '
    '
    ' 7
    ' define_instance_on_supply_chain
    ' nareszcie ostatnia rzecz
    ' czyli podsumowanie wsyzstkich scenariuszy
    ' dwie dodatkowe kolumny INSTANCE definiuje zlamania supply chain
    '
    '
    '
    ' 8
    ' Private Sub fill_urgency(r As Range)
    ' URGENCY - kolumna jako labelka ktora w sumie odwzorowuje INSTANCE :D
    
    
    
    StatusBox.Show
    Application.ScreenUpdating = False
    StatusBox.ProgressBar.Value = 0
    StatusBox.ProgressBar.Max = 8
    
    Dim coll As Collection, k As Komentarz, little_coll As Collection, rep_sh As Worksheet
    Set k = New Komentarz
    Set coll = Nothing
    Set coll = New Collection

    Dim rng As Range
    Set rep_sh = ActiveSheet
    Set rng = ActiveSheet.Range("i2")
    
    If rng <> "Past due" Then
        MsgBox "to nie jest kowerydz!"
        Exit Sub
    End If
    
    Dim r As Range
    Dim i As ITransit
    Dim td As TransitDetails
    Set rng = Range(rng.offset(1, 0), rng.offset(4, 20))
    StatusBox.Description.Caption = "Filling collection"
    StatusBox.Repaint
    DoEvents
    
    loop_to_fill_td rep_sh, rng, little_coll, td, coll, k
    
    ProgressIncrease
    
    ' nowy arkusz i wyrzucenie danych
    ' sh to jest arkusz danych posrednich jak dobrze pojdzie potem bedzie mozna bedzie usuanac
    ' dodatkowo zmienna routing_cell ktora wskazuje na a2 arkusza z ktorego powstaje flat table
    Dim sh As Worksheet, routing_cell As Range
    Set routing_cell = Range("b1") ' the now symbol is most uniq
    Set sh = ThisWorkbook.Sheets.Add
    sh.Name = "FLAT_" & CStr(source_sheet.Name)
    
    StatusBox.Description.Caption = "Labels"
    StatusBox.Repaint
    DoEvents
    Set rng = sh.Range("a1")
    ' labels
    fill_labels rng, routing_cell
    
    ProgressIncrease
    
    Set rng = sh.Range("a2")
    
    StatusBox.Description.Caption = "Filling data"
    StatusBox.Repaint
    DoEvents
    Dim el As TransitDetails
    fill_sheet_from_coll_el el, coll, rng
    
    '4.1
    ' i think t should be obsolete
    ' no i  leave it like it is - so in REM
    ' this subroutine is outside of this module - is in catchCONTModule
    ' prepare_list_of_containers_in_register "init"
    Application.EnableEvents = False
    
    ProgressIncrease
    
    
    ' teraz zamaist trazytow dorzucamy formuly -
    ' gdzie o2 to pierwszy element tabeli tranzytow gdzie kolumna O -> przechowuje adr row
    ' wiersz w jakim sie znajduje wybrany transit oczywiscie musismy wyfiltrowac dane ktore nie spelniaja warunkow
    ' ta petla ma sluzyc wypelnieniu za pomoca formul arkusza coverage
    ' pobierac bedzie dane z arkusza proxy
    StatusBox.Description.Caption = "Crossed Formulas on Coverage"
    StatusBox.Repaint
    DoEvents
    
    
    Dim komorka_kowerydza As Range
    fill_coverage_with_formulas sh, komorka_kowerydza, rep_sh
    ProgressIncrease
    
    ' potem jak juz z powrotem nalozylismy wartosci powinno byc wszystko ok
    
    StatusBox.Description.Caption = "Formulas on Toggle"
    StatusBox.Repaint
    DoEvents
    
    Dim source As Range, place As Range
    
    
    ' narazie wszystkie toggle sa na 1
    Set source = Range(sh.Range("s2"), sh.Range("s2").End(xlDown))
    Set place = sh.Range("u2")
    copy_s_to_t_as_values sh, source, place
    ProgressIncrease
    
    
    StatusBox.Description.Caption = "Status 4"
    StatusBox.Repaint
    DoEvents
    
    put_formula_into_toggle_column sh, 4
    ' teraz znow kopiujemy dane statycznie, tyle tym razem na na wiekszym shortage
    Set source = Range(sh.Range("s2"), sh.Range("s2").End(xlDown))
    Set place = sh.Range("w2")
    copy_s_to_t_as_values sh, source, place
    ProgressIncrease
    
    
    StatusBox.Description.Caption = "Status 3"
    StatusBox.Repaint
    DoEvents
    
    ' tutaj pierwsza instancja toggle ze status 3 jako pierwsze asny ktore widac z punktu widzenia coverage'a
    put_formula_into_toggle_column sh, 3
    ' teraz znow kopiujemy dane statycznie, tyle tym razem na na wiekszym shortage
    Set source = Range(sh.Range("s2"), sh.Range("s2").End(xlDown))
    Set place = sh.Range("v2")
    copy_s_to_t_as_values sh, source, place
    
    ' ok mamy kopie trzech scenarioszy runout vs eda
    ' teraz wystarczy podliczyc ile razy wystpilo conajmniej zero aby moc zdefiniowac
    ' w ktorym miejscu lancuch zostanie przerwany (bedziemy okreslac instancje problemu)
    define_instance_on_supply_chain Range(sh.Range("y2"), sh.Range("y2").End(xlDown))
    
    ProgressIncrease
    StatusBox.Hide
    
    fill_urgency Range(sh.Range("y2"), sh.Range("y2").End(xlDown)).offset(0, 1)
    
    ' jesli tabela ta potem bedzie usuwana autofit mija sie z celem
    ' bo i przeciez czas zajmuje niepotrzebnie
    sh.Columns("A:AA").EntireColumn.AutoFit
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub define_instance_on_supply_chain(r As Range)

    Dim item As Range, i As Range
    For Each i In r
        Set item = Range(i.offset(0, -1), i.offset(0, -4))
        ' jedna iteracja
        i = Int(Application.WorksheetFunction.CountIf(item, "<1"))
    Next i
End Sub

Private Sub fill_urgency(r As Range)

    Dim i As Range
    For Each i In r
        i = ThisWorkbook.Sheets("register").Range("c" & CStr(i.offset(0, -1) + 1))
    Next i
End Sub

Private Sub put_formula_into_toggle_column(sh As Worksheet, st As Integer)


    ' zakladam ze wszystko dziala wiec zadnych dodatkowych ifow nie zamierzam tutaj pakowac
    ' zatem jesli cos sie wypnie to prawdopodobnie tutaj :)
    Dim item As Range
    Set item = sh.Range("t2") ' this is a begining for the toggle column
    Do
        
        
        item.Formula = "=0"
        ' ta formula jest dosyc sroga
        ' poniewaz mamy dwie opcje
        ' jedna z orem druga z andem
        ' no i teraz klocki poniewaz moj umysl swego czasu byl nastawiony na logike odwrotna
        ' tj. to co oznacze wyrzuc z symulacji
        ' a w sumie gosciu w zaragozie bedzie myslal dodaj cos co bede rozpatrywal
        ' i teraz ciekawy jestem jak to rozwiazac aby algorytm pozostal taki sam
        ' a uzytkownikowi dawal zludzenie ze nie odejmuje kolejnych kontenerow z symulacji a dodaje :P
        ' przez czas jakis uzywalem and dzieki temu warunek wyrzucania symulacji dzialal tylko wtedy kiedy obydwie flagi sie ze soba zgadzaly
        ' czyli ze nie ma tak latwo z modyfikowaniem symulacji
        ' odkad zrezygnowale z listy pomocniczej w rejestrze nadalo wydaje mi sie sensem bycia zmiany z powrotem na OR
        ' dzieki czemu elementy ustawione na zero sam nawzajem dla siebie przezroczyste :) - a czy to jest dobre rozwiazanie pokaze czas testowania
        item.Formula = ThisWorkbook.Sheets("register").Range("formulaifor").Formula
        tmp_str = CStr(item.Formula)
        tmp_str = Replace(tmp_str, "xx", CStr(item.offset(0, -6).Address) & ">=" & CStr("togglehandler"))
        tmp_str = Replace(tmp_str, "yy", "1")
        tmp_str = Replace(tmp_str, "zz", "0")
        
        
        ' toggle with additional statement thanks to OR formula
        ' obsolete registrowy
        'Dim beg_of_uniq_cont_list As Range
        'Set beg_of_uniq_cont_list = ThisWorkbook.Sheets("register").Range("h2")
        'Set beg_of_uniq_cont_list = Range(beg_of_uniq_cont_list, beg_of_uniq_cont_list.End(xlDown).offset(0, 2))
        
        ' tutaj trzeba nieco zmienic implementacje aby nie laczyla sie z rejestrem
        ' ============================================================================================================
        ' tmp_str = Replace(tmp_str, "qq", "catchcontainers(" & item.offset(0, -14).AddressLocal & ",register!" & beg_of_uniq_cont_list.Address & ")") ' ma byc przezroczyste dla false w przypadku or wystarczy ze caly czas ta czesc warunku bedzie falszywa
        tmp_str = Replace(tmp_str, "qq", CStr(item.offset(0, 7).Address))
        ' ============================================================================================================
        
        item.Formula = tmp_str
        ' item.Calculate
        
        Set item = item.offset(1, 0)
    Loop While item <> ""
    
    
    Set item = Range(sh.Range("t2"), sh.Range("t2").End(xlDown))
    item.offset(0, -16).Calculate
    item.offset(0, -1).Calculate
    item.Calculate
    
    
    DoEvents
End Sub

Private Sub copy_s_to_t_as_values(sh As Worksheet, source As Range, place As Range)


    ' copy with toggle on before I change it
    ' wczesniej byla tu masakra przeliczajaca milion razy poprawnie matematycznie
    ' ale jesli pojsc po rozum do glowy wystarczy skopiowac dane z kolumny jako wartosci i juz
    
    
    source.Copy
    place.PasteSpecial xlPasteValues
    DoEvents
    Application.CutCopyMode = False
    sh.Range("f1").Select
End Sub

Private Sub fill_coverage_with_formulas(sh As Worksheet, komorka_kowerydza As Range, rep_sh As Worksheet)

    Set r = sh.Range("a2")
    Do
        ' pierwszy warunek ktory okresla ze bawimy sie tylko transitami
        ' ktore mieszcza sie przedziale od 0 do 2
        Range(r, r.offset(0, 18)).Font.Bold = False
        
            
    
        ' dzieki prostej formule mozemy wyluskac to co potrzeba - tj. nazwe arkusza :D
        ' nazwa_arkusza = wyluskaj_nazwe_arkusza_z_formuly_adresu(CStr(r.offset(0, 3).Formula))
        ' przejdz do komorki w kowerydzu
        ' set komorka_kowerydza =
        ' ThisWorkbook.Sheets("register").Range("formulaif")
        Set komorka_kowerydza = rep_sh.Cells(Int(r.offset(0, 14)), Int(r.offset(0, 15)))
        If Not komorka_kowerydza.HasFormula Or Application.WorksheetFunction.IsError(komorka_kowerydza) Then
            komorka_kowerydza.Formula = "=" & sh.Name & "!" & r.offset(0, 11).Address
        Else
            komorka_kowerydza.Formula = komorka_kowerydza.Formula & "+" & sh.Name & "!" & r.offset(0, 11).Address
        End If
        
        If (Int(r.offset(0, 13)) = 2) Then
            Range(r, r.offset(0, 18)).Font.Bold = True
            ' Range(r, r.offset(0, 18)).Font.Color = RGB(220, 0, 0)
        End If
    
        r.offset(0, 11) = ""
        r.offset(0, 11).Formula = ThisWorkbook.Sheets("register").Range("formulaifor").Formula
        tmp_str = CStr(r.offset(0, 11).Formula)
        tmp_str = Replace(tmp_str, "xx", CStr(r.offset(0, 19).Address) & "=0")
        tmp_str = Replace(tmp_str, "yy", "")
        tmp_str = Replace(tmp_str, "zz", CStr(r.offset(0, 17).Address))
        
        ' tutaj dodatkowy or ktory ma sprawowac wladze nad kontenerami
        ' tu tak zostanie forver poniewaz to jest qty ktore jest dopiero zalezne od TOGGLE
        ' zostawie dla potomnych moze sie jeszcze przyda
        ' ale narazie niech bedzie na zero i niech nie przeszkadza
        tmp_str = Replace(tmp_str, "qq", "0") ' ma byc przezroczyste dla false w przypadku or wystarczy ze caly czas ta czesc warunku bedzie falszywa
        r.offset(0, 11).Formula = tmp_str
        
        Set r = r.offset(1, 0)
    Loop While r <> ""
End Sub


Private Sub fill_sheet_from_coll_el(el As TransitDetails, coll As Collection, rng As Range)

    For Each el In coll
        ' MsgBox el.firstRunout & " " & el.t.mName & " " & el.t.mQty
        rng = el.plt
        rng.offset(0, 1) = el.pn
        rng.offset(0, 2) = el.duns
        rng.offset(0, 3).Formula = el.firstRunout
        rng.offset(0, 4) = el.scac
        rng.offset(0, 5) = el.kontener
        rng.offset(0, 6) = el.t.mDeliveryDate
        rng.offset(0, 7) = el.t.mDeliveryTime
        rng.offset(0, 8) = el.t.mPickupDate
        rng.offset(0, 9) = el.t.mMODE
        rng.offset(0, 10) = el.t.mName
        rng.offset(0, 11) = el.t.mQty
        rng.offset(0, 12) = el.t.mRoute
        rng.offset(0, 13) = el.t.mST
        rng.offset(0, 14) = el.P.row
        rng.offset(0, 15) = el.P.col
        rng.offset(0, 16) = el.EDA_CW
        rng.offset(0, 17) = el.t.mQty
        rng.offset(0, 18).Formula = "=" & rng.offset(0, 3).Address & "-" & rng.offset(0, 16).Address
        rng.offset(0, 19) = 1
        
        If Application.WorksheetFunction.IsError(rng.offset(0, 3)) Then
            rng.offset(0, 20) = "need_to_redefine"
        Else
            rng.offset(0, 20) = rng.offset(0, 3) - rng.offset(0, 16)
        End If
        
        ' teraz bedzie male zamieszanie poniewaz odwrocilem w flat table dane
        ' toggle z eda vs runout
        
        rng.offset(0, 21) = 1
        rng.offset(0, 22) = 1
        ' doh -> gdzie doh -> czyli wsyzstkie asny usuwamy i sprawdzamy runout vs eda cw
        rng.offset(0, 23) = el.firstCW + Int(el.doh / 7) - rng.offset(0, 16)
        rng.offset(0, 24) = 0
        ' rng.offset(0, 25) tj. urgency column
        ' tutaj dla managera kontenerow dodatkowa kolumna na flat table
        ' na poczatku zero :D bo i tak daje status 3 scenario i zero dla tego scenariusza jest przezroczyste
        rng.offset(0, 26) = 0
        
        
        Set rng = rng.offset(1, 0)
    Next el
End Sub


Private Sub fill_labels(rng As Range, rng_from_coverage As Range)
    rng = "PLT " & CStr(rng_from_coverage)
    rng.offset(0, 1) = "PN"
    rng.offset(0, 2) = "DUNS"
    rng.offset(0, 3) = "FIRST RUNOUT"
    rng.offset(0, 4) = "SCAC"
    rng.offset(0, 5) = "CONTAINER"
    rng.offset(0, 6) = "EDA"
    rng.offset(0, 7) = "ETA"
    rng.offset(0, 8) = "SDATE"
    rng.offset(0, 9) = "MODE"
    rng.offset(0, 10) = "NAME"
    rng.offset(0, 11) = "QTY"
    rng.offset(0, 12) = "ROUTE"
    rng.offset(0, 13) = "ST"
    rng.offset(0, 14) = "ADR ROW"
    rng.offset(0, 15) = "ADR COL"
    rng.offset(0, 16) = "EDA CW"
    rng.offset(0, 17) = "REF QTY"
    rng.offset(0, 18) = "RUNOUT vs EDA formula"
    rng.offset(0, 19) = "TOGGLE"
    rng.offset(0, 20) = "RUNOUT vs EDA"
    rng.offset(0, 21) = "RUNOUT <ST3 vs EDA"
    rng.offset(0, 22) = "RUNOUT <ST4 vs EDA"
    rng.offset(0, 23) = "RUNOUT DOH vs EDA"
    rng.offset(0, 24) = "INSTANCE"
    rng.offset(0, 25) = "URGENCY"
    rng.offset(0, 26) = "CONT MANAGER"
End Sub


Private Sub loop_to_fill_td(rep_sh As Worksheet, rng As Range, little_coll As Collection, td As TransitDetails, coll As Collection, k As Komentarz)


    Dim r As Range
    Do
        ' rng.Select
        For Each r In rng
        
            Set little_coll = Nothing
            Set little_coll = New Collection
            Set little_coll = k.pobierz_dane(r)
            If Not little_coll Is Nothing Then
                For Each i In little_coll
                
                    Set td = Nothing
                    Set td = New TransitDetails
                    ' dodajemy element ITransit jako komponent dla TransitDetails
                    Set td.t = i
                    td.firstRunout = CStr("=" & CStr(rep_sh.Name) & "!" & rng.item(1, 1).offset(0, -2).Address)
                    td.duns = CStr(rng.item(1, 1).offset(2, -7))
                    td.plt = CStr(rng.item(1, 1).offset(0, -8))
                    td.pn = CStr(rng.item(1, 1).offset(0, -7))
                    td.scac = Left(i.mTRLR, 4)
                    If (Len(i.mTRLR) - 4) > 0 Then
                        td.kontener = Right(i.mTRLR, Len(i.mTRLR) - 4)
                    End If
                    If CStr(Trim(rng.item(1, 1).offset(1, -3))) = "" Then
                    
                        td.doh = 0
                    Else
                        td.doh = CDbl(rng.item(1, 1).offset(1, -3))
                    End If
                    
                    td.firstCW = CStr(Int(Year(Date))) & CStr(rng.item(1, 1).offset(-1, 1))
                    
                    td.fill_eda_cw
                    
                    ' td position
                    td.P.adr = r.AddressLocal
                    td.P.col = r.Column
                    td.P.row = r.row
                    
                    coll.Add td
                Next i
            End If
        Next r
        Set rng = rng.offset(7, 0)
    Loop While rng.item(1, 1).offset(-1, 0) <> ""
End Sub

Private Function wyluskaj_nazwe_arkusza_z_formuly_adresu(s As String)

    ' tymczasowa wartosc, ktora przechowuje ktory z kolei jest wykrzyknik
    tmp = Int(Application.WorksheetFunction.Find("!", CStr(r.offset(0, 3).Formula)))
    ' dzieki prostej formule mozemy wyluskac to co potrzeba - tj. nazwe arkusza :D
    nazwa_arkusza = Mid(CStr(r.offset(0, 3).Formula), 2, tmp - 1)
End Function
