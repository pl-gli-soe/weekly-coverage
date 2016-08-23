Attribute VB_Name = "CatchCONTModule"
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


Public Sub containers_from_pivot(ictrl As IRibbonControl)



    

    Application.EnableEvents = False
    ' jesli chodzi o scenario kontenerow z guzika ma byc przezroczysty tj. ze stare zmiany maja pozostac nienaruszone
    ' i jest to bardzo wygodne
    ' jednak jesli chodzi o przygotowanie scenariusza na podstawie selekcji chcemy miec full picture na podstawie tego co sobie sami wybralismy bez
    ' pozostawienia na tapecie czego kolwiek innego
    ' no i teraz najgorsze jeszcze przede mna poniewaz
    ' oprocz wybierania samych kontenerow
    ' nalezy rowniez uwzgleni te ktore sa z punktu widzenia wybieranych kontenerow przeszle
    ' najprosciej rzecz ujmujac trzeba sprawdzic ktore kontenery z naszego wyboru beda najszybciej i w ten sposob odwzorowac reszte :D
    ' oczywiscie nie jest to rozwiazanie idealne ale coz
    ' moze z czasem na cos fajnego wpadne


    If check_is_pivot(ActiveSheet) Then
    
    
        ' ThisWorkbook.Sheets("register").Range("togglehandler") = 7
        Dim r As Range, flat As Worksheet
        
        
        ' tutaj trzeba pamietac ze ten selection przetrzymuje surowe dane, ktore w zaden sposob
        ' nie zostaly sprawdzone przez algorytm
        ' co jest stosunkowo nie fajne i moze latwo wyrzucic blad wiec trzeba bedzie o tym rowniez pomyslec
        ' sprawdzenie jednak nie bedzie twarde tak mysle teraz (tj. 2014-07-22)
        ' czyli pierwszy dzien po urlopie - moge nie miec tego owego w glowie jeszcze
        ' sprawa bedzie duzo prostsza wrzuc co kolwiek do selekcji a potem dopaduj
        ' jesli beda bzdury po prostu nic nie zostanie wybrane
        ' czy bedzie mozna robic zbiory nadmiarowe i ktore przypadkiem moga trafic w liste badz w ogole
        ' a wiec looz
        ' zatem kontynuujac lista z selekcji wcale nie musi byc z pivota ale z dowolnego zrodla
        Set r = Selection
        
        Set flat = go_to_correspoding_flat_table(ActiveSheet)
        flat.Activate
        
        inner_start_form_remove_containers PIVOT_SCENARIO, r
        
        
        ' MsgBox "scenario based on selection from pivot is ready!"
    Else
        ' MsgBox "this selection is not from pivot table"
        ' te dane beda customowe wiec bedzie musial uzytkownik okreslic z jakiego flat table dane maja byc sprawdzane hmmm...
        ' na ten czas czekac bede na pozniejsze stworzenie algorytmu ktory edzie bardzo srogi i wiele warunkow bedzie mialo swoje rozwiazanie
        ' i to bardzo mi sie pododba :)
        '
        ' MsgBox "for custom data implementation is not ready yet"
        
        
        ' zatem jak to rozwiaze...
        ' problem zaczyna sie od tego ze implementacja zmian nazw po routingu jest wysoce niestabilna i bedzie trzeba ja prawdopodobnie w czasie pozniejszym poprawic
        ' mimo wszysko
        ' zaczynaja sie jaja szczegolnie gdy nazwy wlasne zawieraja nazwy kluczowe plus jesli zmieniam dane tak ze usuwam prefixy a w niektorych nie
        ' to wtedy rowniez wszystko sie miesza niesamowicie
        ' jednak jesli to pomijac wystarczy znalezc odpowiedni flat table aby wykonac w miare prosty sposob dane
        ' tak teraz pomyslalem ze jesli nie jestem na pivot sheet to moge najzwyczajniej w swiecie uruchomic managera kontenerow
        ' bez zadnych narzucen z wczesniejsza informacja na msgbox
        ' zatem:
        MsgBox "you are not in pivot table, click ok & start container manager without any additional actions!"
        inner_start_form_remove_containers GUZIK
        
    End If
    
    Application.EnableEvents = True
End Sub

Private Function check_is_pivot(ash As Worksheet) As Boolean

    ' te sprawdzenie jest bardzo prymitywne i czas pomyslec niedlugo o nieco bardziej finezyjnym rozwiazaniu
    If (ash.Name Like "PIVOTW_*") Or (ash.Name Like "PIVOTD_*") Then
        check_is_pivot = True
    Else
        check_is_pivot = False
    End If
End Function


Private Function go_to_correspoding_flat_table(sh As Worksheet) As Worksheet

    
    Dim s As Worksheet
    For Each s In ThisWorkbook.Sheets
        If check_labels(s) Then
            If "PLT " & CStr(sh.Range("a1")) = CStr(s.Range("a1")) Then
                Set go_to_correspoding_flat_table = s
                Exit Function
            End If
        End If
    Next s
End Function


Public Sub start_form_remove_containers(ictrl As IRibbonControl)

    inner_start_form_remove_containers GUZIK
End Sub


Private Sub inner_start_form_remove_containers(arg As starter_dla_listy_kontenerow_w_formularzu, Optional lista As Range)



    ' super wazne!
    ' arg to enum ktory przyjmuje dwie mozeliwosci
    ' uruchamianie z poziomu guzika wybierania kontenerow ktore nas interesuja
    ' 2 wybieramy za pomoca selekcji liste kontenerow a one nam sie ladnie podswietlaja jesli znajda dopasowanie
    ' argument lista -> uruchmiany w przypadku gdy enum jest wybrany na pivota wtedy pobiera selekcji i po kolei dopasowuje :D
    

    ' globalna funkcja sprawdzjaca czy jest na flat table
    If check_labels(ActiveSheet) Then
    
        Dim zmienna_bool As Boolean
        ' ten zasieg jest tymczasowym zasiegiem po to aby mozna bylo w prosty sposob odnalezc dany kontener w flat table
        ' poniewaz sortowanie odbywa sie poprzez worksheet register
        Dim tmp_range As Range
        ' tymczas dla pivot scenario
        ' drugi warunek if handler
        Dim tmp As Range
        Dim sc As Range
        Set sc = ThisWorkbook.Sheets("register").Range("q2")
        Range(sc, sc.offset(0, 1).End(xlDown).offset(0, -1).End(xlUp)).Clear
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        odlegosc_miedzy_cont_a_cont_manger_column = 21
    
    
        Dim zasieg As Range, r As Range
        Set zasieg = Range("A2:A" & CStr(Range("a2").End(xlDown).row)).offset(0, 5)
        zasieg.Copy sc
        Range(sc, sc.offset(0, 1).End(xlDown).offset(0, -1).End(xlUp)).Sort sc, xlAscending
        Set sc = Range(sc, sc.offset(0, 1).End(xlDown).offset(0, -1).End(xlUp))
        StatusBox.Show
        
        StatusBox.ProgressBar.Value = 0
        StatusBox.ProgressBar.Max = sc.Count
    
        RemoveCONT.ListBox1.Clear
        i = 0
        For Each r In sc
            If Not check_if_its_already_in_list(r) Then
                RemoveCONT.ListBox1.AddItem r
                
                
                If arg = GUZIK Then
                    ' help define which should be selected already
                    Set tmp_range = zasieg.Find(r)
                    If tmp_range.offset(0, odlegosc_miedzy_cont_a_cont_manger_column) = 1 Then
                        zmienna_bool = True
                        RemoveCONT.ListBox1.Selected(i) = zmienna_bool
                    Else
                        zmienna_bool = False
                        RemoveCONT.ListBox1.Selected(i) = zmienna_bool
                    End If
                    
                ElseIf arg = PIVOT_SCENARIO Then
                    
                    ' ten algorytm jest ok
                    ' ale nie daje kontenerow ktore sa z perspektywy wybranych przeszle
                    ' nalezy dodatkowo opracowac prosty algorytm ktory ujmuje je w pracy i wyniku raportu bardziej
                    ' prawdopodobny scenariusz :D
                    
                    ' ewentualna opcja jest ustawienie flag na scenariuszu z perspektywy statusu 3 czyli
                    ' zostawienia wszsytkich asnow ktore sa juz na ladzie
                    
                    ' do przedyskutowania z kolegami z ZA
                    
                    Set tmp = Nothing
                    Set tmp = lista.Find(CStr(RemoveCONT.ListBox1.List(i)))
                    If Not tmp Is Nothing Then
                        zmienna_bool = True
                    Else
                        zmienna_bool = False
                    End If
                    
                    RemoveCONT.ListBox1.Selected(i) = zmienna_bool
                End If
                
                i = i + 1
                
            End If
            StatusBox.Description.Caption = CStr(r)
            StatusBox.Repaint
            DoEvents
            
            
            
            ProgressIncrease
        Next r
        
        RemoveCONT.Show
        
        StatusBox.Hide
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        
        
        ' to tutaj jest bez sensu bo i tak zostanie wykonane nawet jesli w formie klikniesz na "x" co nie jest pozadane :D
        ' MsgBox "ready!"
        
    Else
        MsgBox "flat table must be active worksheet"
    End If
End Sub

Private Function check_if_its_already_in_list(r As Range) As Boolean

    check_if_its_already_in_list = False

    With RemoveCONT.ListBox1
        If .ListCount > 0 Then
            For x = 0 To .ListCount - 1
                If CStr(r) = CStr(.List(x)) Then
                    check_if_its_already_in_list = True
                    Exit Function
                End If
            Next x
        End If
    End With
End Function
