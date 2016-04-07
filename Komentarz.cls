VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Komentarz"
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


Private tresc As String
Private cmnt As Comment
Private it As ITransit

' 1. name
' 2. mode
' 3. status
' 4. eda
' 5. eta
' 6. sdate
' 7. trlr
' 8. qty
' 9. route


' zasada dzialania tej klasy jest daleka od idealu
' jednakze zaklada uruchomienie w miare przejrzystych procedur

' 1. stworzenie komentarza
'   ta metoda jest nawet swoiscie przezroczysta ponieaz nawet jesli komentarz juz istnieje
'   nie spowoduje to natychmiastowego zwieszenia systemu ale zgodnie z algorytmem
'   przypisze do pola juz istniejacego komentarza a nawet doda text z drugiego argumentu
'   zatem teoretycznie metoda ta moze sluzyc rownoczesnie jako add addtional comment
'   jednak unikamy tego typu wykorzystania

' 2. przechwyc komentarz
' ten element wykonuje tylko czesc dzialania metody stworzenia
' dodatkowo jesli proba przechwytu bedzie nieudana
' klasa wykasuje wszystkie swoje pola (tak jakby co)

' 3. dodaj asn
' dodaje tranzyt do komentarza
' jest to najbardziej wrazliwa metoda calej klasy ponieaz nie jest odporna na dewiacje
' takie jak brak komentarza
' poza tym dosyc surowo oczekuje ze parametrem jest abstrakcyjny typ ITransit a wiec uzaleznienie
' od zewnetrznej implementacji jeszcze sie zastanowie czy to aby napewno dobry pomysl

' 4. pobierz dane
' jako metoda ktora wyrzuca na swiat zewnetrzny kolekcje danych z komentarza
' zalozenie jest takie ze ta kolekcja zawiera zbior typow ITransit a wiec tam sam newralgiczny element co w metodzie 3. dodaj asn

' PS. jako niby 5. czy_istnieje
' to jest metoda ktora wystepowac bedzie nieregularnie na potrzeby sprawdzania wypadkowego czy w danej komorce wystepuje de facto komentarz (zwrot typu boolean)
'
'
'
'
' jednak tak popatrzylem na implementacje znajdujaca sie w CoverageModule
' i musze przyznac ze to musi sie skonczyc na jednej przejrzystej metodzie, ktora jest super transparent
' wszystkiego sama sie dowiadywac i odpowiednio podejmowac decyzje ooooo.
' nazwiemy ja: DODAJ i juz :D

Private Sub Class_Initialize()
    tresc = ""
    Set cmnt = Nothing
End Sub

Public Property Get x_tresc() As String
  x_tresc = tresc
End Property

Public Property Let x_tresc(arg As String)
   tresc = arg
End Property

Public Function czy_istnieje(Optional mr As Range) As Boolean
    
  czy_istnieje = False
  
  If Not mr.Comment Is Nothing Then
    czy_istnieje = True
  End If
End Function

Public Function przechwyc_komentarz(mr As Range) As Comment
    If Not mr.Comment Is Nothing Then
        Set cmnt = mr.Comment
        tresc = CStr(cmnt.Text)
        Set przechwyc_komentarz = cmnt
    Else
        Set cmnt = Nothing
        tresc = ""
        Set it = Nothing
    End If
End Function



Public Function policz_linie_komentarza(mr As Range) As Integer
    If czy_istnieje(mr) Then
    
        policz_linie_komentarza = Int(CountLines(mr.Comment.Text))
    Else
        policz_linie_komentarza = -1
    End If
End Function

Public Function pobierz_dane(mr As Range) As Collection

    Set pobierz_dane = Nothing
    

    ' ten element ma dac nam mozliwosc wyciagniecia powtornego danych z komentarza
    ' przeskakiwanie przez kolejne coverage (tj. pasek z jednym pn)
    ' mysle ze bedzie podobny implementacyjnie do copy critical  ktory tez przeciez musi wysluskac elementy
    ' z calego duzego coveraga
    ' kolejna sprawa jest czy na pewno chcemy aby metoda pobierz dane przechodzi³a przez caly zakres komorka po komorce
    ' moze dodatkowa wrzuce warunek if - ale wtedy bedzie nadmiarowa implementacja w samej klasie pobierania danych - tranzytow
    ' wiec z punktu widzenia przejrzystosci tj dobry  pomysl aby przechodzic przez wszystkie komorki ta metoda
    ' bedzie bardzo ladnie wygladac dana implementacja
    ' a tak czy siak musimy przjesc przez wszystkie komorki
    ' wiec im wiecej schowamy algorytmu w inne metody tym lepiej
    ' a co najwazniejsze ponizej i tak juz sprawdzamy czy komentarz istnieje
    ' mozemy co najwyzej dorzucic warunek dodatkowy sprawdzajacy czy dany komentarz na pewno jest komentarzem dla asna
    If Not mr.Comment Is Nothing And Cells(mr.row, 1) <> "PLT" Then
        tresc = mr.Comment.Text
        Set cmnt = mr.Comment
        
        Dim tmp_kolekcja As Collection, ia As ITransit ' ia is item_asn
        Set tmp_kolekcja = Nothing
        Set tmp_kolekcja = New Collection
        
        arr = Split(tresc, Chr(10))
        For x = LBound(arr) To UBound(arr)
        
            
            ' 1. name
            ' 2. mode
            ' 3. status
            ' 4. eda
            ' 5. eta
            ' 6. sdate
            ' 7. trlr
            ' 8. qty
            ' 9. route
            ' Debug.Print arr(x) ' OK
            ' to jest koniec kazdego asna
            
            ' to jest pierwszy element komentarza danego asna
            If arr(x) Like "1. NM:*" Then
                ' zatem skoro pierwszy mozna bez bolu zainicjowac nowy obiekt klasy asn (ITransit)
                Set ia = Nothing
                Set ia = New ASN
                ' ===========================
                ' NM
                ia.mName = CStr(Trim(Right(arr(x), Len(arr(x)) - 7)))
                ' MODE
                ia.mMODE = Trim(Right(arr(x + 1), Len(arr(x + 1)) - 9))
                ' ST
                ia.mST = Trim(Right(arr(x + 2), Len(arr(x + 2)) - 7))
                ' EDA
                tmp_d = Trim(Right(arr(x + 3), Len(arr(x + 3)) - 8))
                
                If tmp_d Like "??/??/????" Then
                    tmp_d = Right(tmp_d, 4) & "-" & Mid(tmp_d, 4, 2) & "-" & Left(tmp_d, 2)
                ElseIf tmp_d Like "??.??.????" Then
                    tmp_d = Right(tmp_d, 4) & "-" & Mid(tmp_d, 4, 2) & "-" & Left(tmp_d, 2)
                ElseIf tmp_d Like "????-??-??" Then
                    tmp_d = Left(tmp_d, 4) & "-" & Mid(tmp_d, 6, 2) & "-" & Right(tmp_d, 2)
                End If
                
                ' tmp_d = Format(tmp_d, "dd/mm/yyyy")
                ia.mDeliveryDate = CDate(Format(tmp_d, "yyyy-mm-dd"))
                ' ETA
                ia.mDeliveryTime = Trim(Right(arr(x + 4), Len(arr(x + 4)) - 8))
                ' SDATE
                tmp_d = Trim(Right(arr(x + 5), Len(arr(x + 5)) - 10))
                
                If tmp_d Like "??/??/????" Then
                    tmp_d = Right(tmp_d, 4) & "-" & Mid(tmp_d, 4, 2) & "-" & Left(tmp_d, 2)
                ElseIf tmp_d Like "??.??.????" Then
                    tmp_d = Right(tmp_d, 4) & "-" & Mid(tmp_d, 4, 2) & "-" & Left(tmp_d, 2)
                ElseIf tmp_d Like "????-??-??" Then
                    tmp_d = Left(tmp_d, 4) & "-" & Mid(tmp_d, 6, 2) & "-" & Right(tmp_d, 2)
                End If
                
                
                ia.mPickupDate = CDate(Format(tmp_d, "yyyy-mm-dd"))
                ' TRLR
                ia.mTRLR = Trim(Right(arr(x + 6), Len(arr(x + 6)) - 9))
                ' QTY
                ia.mQty = Trim(Right(arr(x + 7), Len(arr(x + 7)) - 8))
                ' ROUTE
                ia.mRoute = Trim(Right(arr(x + 8), Len(arr(x + 8)) - 10))
                ' ===========================
                x = x + 9
            End If
            If arr(x) Like "*============*" Then
                tmp_kolekcja.Add ia
            End If
        Next x
        
        Set pobierz_dane = tmp_kolekcja
        
    End If
    
    
    
End Function

Private Sub p_wypelnij(ByRef ia As ITransit, ByRef arr() As Variant, x As Variant)
    
    
End Sub

Public Function dodaj_asn(it As ITransit) As String
' 1. name
' 2. mode
' 3. status
' 4. eda
' 5. eta
' 6. sdate
' 7. trlr
' 8. qty
' 9. route
' 10. eda - sdate in days

    tmp_txt = "1. NM: " & CStr(it.mName) & Chr(10) & _
        "2. MODE: " & CStr(it.mMODE) & Chr(10) & _
        "3. ST: " & CStr(it.mST) & Chr(10) & _
        "4. EDA: " & CStr(Format(CDate(it.mDeliveryDate), "yyyy-mm-dd")) & Chr(10) & _
        "5. ETA: " & CStr(it.mDeliveryTime) & Chr(10) & _
        "6. SDATE: " & CStr(Format(CDate(it.mPickupDate), "yyyy-mm-dd")) & Chr(10) & _
        "7. TRLR: " & CStr(it.mTRLR) & Chr(10) & _
        "8. QTY: " & CStr(it.mQty) & Chr(10) & _
        "9. ROUTE: " & CStr(it.mRoute) & Chr(10) & _
        "10. EDA-SDATE: " & CStr(it.mDeliveryDate - it.mPickupDate) & Chr(10) & _
        "==============="
            
   
   
    dodaj_asn = tmp_txt
    
End Function

Public Function stworz_komentarz(mr As Range, Optional txt As String) As Boolean
    If czy_istnieje(mr) Then
        Set cmnt = mr.Comment
        tmp_width = mr.Comment.Shape.Width
        tmp_height = mr.Comment.Shape.Height
        s_new_cmnt = cmnt.Text & Chr(10) & CStr(txt)
        mr.AddComment (CStr(s_new_cmnt)) & Chr(10)
        Set cmnt = mr.Comment
        cmnt.Shape.Width = tmp_width
        cmnt.Shape.Height = tmp_height + G_COMMENT_HEIGHT
        stworz_komentarz = False
    Else
    
        mr.AddComment (CStr(txt)) & Chr(10)
        Set cmnt = mr.Comment
        cmnt.Shape.Width = G_COMMENT_WIDTH
        cmnt.Shape.Height = G_COMMENT_HEIGHT
        stworz_komentarz = True
    End If
End Function

Public Sub dodaj(mr As Range, it As ITransit)
    If czy_istnieje(mr) Then
        Set cmnt = przechwyc_komentarz(mr)
        s_new_cmnt = cmnt.Text & Chr(10) & dodaj_asn(it)
        tmp_width = mr.Comment.Shape.Width
        tmp_height = mr.Comment.Shape.Height
        mr.ClearComments
        mr.AddComment (CStr(s_new_cmnt)) & Chr(10)
        Set cmnt = mr.Comment
        cmnt.Shape.Width = tmp_width
        cmnt.Shape.Height = tmp_height + G_COMMENT_HEIGHT
    Else
        wynik = stworz_komentarz(mr, dodaj_asn(it))
        If Not wynik Then
            MsgBox "cos poszlo nie tak z dodawaniem komentarza do komorki na dzien dobry! (ten komentarz przy prawidlowej pracy w ogole nie powinien sie pojawic)"
        End If
    End If
End Sub

