Attribute VB_Name = "VersionModule"
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



'to simplify...'
''
'
'coverage 3.30 jako makro wejsciowe pierwsze kilka podwersji ma byc przetarciem
'nastepnie potrzebny bedzie fire flake jako generator czesci dla zaragozy robie poranny zrzut w poniedzialek -> tutaj na miejscu DOH 4 - 10
'2014-06-02 7:50 - 8:10
'
'no ale sie okazalo ze ten wybor Za okazal sie bardzo kiepski to jakis lokalny FUP z 2 tyg TTIMEm na czesciach i ogolnie nic ciekawego
'
'
'
'zatem poprawilem sie i wybralem okolo 100 czesci z osea w gliwicach tutaj dane okazaly sie duzo bardziej przejrzyste
'zadowolony rowniez jestem z szybkosci generowania raportu proxy - ultra szybszy niz ma to miejsce w FF i jego tam generowaniu raportu PUS
'mysle ze tam cos przedobrzylem ale nie chce mi sie teraz bardzo temu przygladac
'
'ogarniamy raczej tylko 1
'ewentualnie potem jeszcze pogadamy
'
'babki sa bardzo konkretne
'widac ze wiedza o co chodzi'
'
'
'przede wszystkim element oddzielenia logiki czesci faktycznie krytycznych ktorymi musi sie zajac FUP bezposrednio
'dopiero potem po odsianiu tych czesci mozemy sie zajac normalnie elementami ktore z punktu widzenia MGO
'
'dodatkowo pasowaloby nie zrobic statycznego zrzutu transportu bo to chyba raczej nic nie dane'
'
'powinno to wygladac na analize what if ktora sprawdza ze gdy zabraknie tego transportu jak zmieni sie first runout
'
'to naruszy odrobine czasowosc wykonania takiego projektu ale bedzie srogi jak juz na niego tak patrze.'
'
'
'
'
'!!!! wersja 3.35 zawieria kopie czystego raportu bez formul
'3.36 rowniez po skopiowaniu nawet nie jest to takie duze jakby sie wydawalo no i co wazne mozna robic coverage ktore sa naprawde duze
'
'
'4.02 jest przygotowana procedura uruchomienia pivot juz po iribbon control ale jeszcze bez guzika!
'
'
'4.04 wersja gotowa bez jeszcze guzikow scenariuszy
'  - zrbione kolory status 0 - 1 dla diff eda - sdate i real ttime from mgo
'  - dodatkowa linijka w komentarzu eda - sdate in days
'4.05 guziki scenariuszy
'
'
'4.09
'- scenario on containers
'- dynamic sh change names on routing pattern
'
'
'4.10 chce rowniez zwykle kopiowanie w grupie kopiowanie critical i kopi runout order
'to swoja droga ale chyba powiniene zrezygnowac z nadmiernego korzystania z arkusza rejestru i dodac nowa kolumne kontrolujaca managera kontenerow :)
'
'
'4.12 kontenery z pivota
'4.13
'
'    If inPLT = "ZC" Then
'        Sess0.screen.sendKeys ("<pf8>")
'        waitForMgo
'    End If
'
'    4.14 dzia³a ca³kiem dobrze oprócz ZC
'
'4.15 ma³a zmiana implementacji ktore bardziej wpasowuje sie w dane od ZC gdy na przyklad CBAL jest puste

' 4,17 wersja ostatnia jak zostala udostepniona przed shutdownem zaweira wlasciwie wszystko lacznie z rqm dla zc po wsadzenia flagi 2000
'
'
'
'
' ostatni zapis z notepada jest u gory
' ============================================================================' ============================================================================
' ============================================================================' ============================================================================
' ============================================================================' ============================================================================
' ============================================================================' ============================================================================
' .--. .-. .. --- MORSE PRIO
'
'
' ' version 4.18
' '
' to jest wersja ktora zawierac bedzie posortowana liste kontenerow w managerze kontenerow - SORT ON CONT MANAGER!!!!!
' plus bedzie to wersja z dobudowka pod rozkmine frau zanieskiej apropos czesci pod ewentualne help shipmenty i ucinananie nadmiarowych zamowien generowanych przez MGO.
' element ten opierac sie bedzie przede wszystkim o kopiowanie cov na mniejsze zakresy z automatu lub opierajac sie o custom liste :)
' '
' ewentualnie cov moze sam sie domyslec part numberow wspolnych ktore potencjanie maja szanse pracy wymiany asnow
'
' wciaz implementacja polega w glownej mierze na korzystaniu z flat table przy czym musi byc ona aktywna podczas tworzenia pivotow, czy tez uruchamiania container managera


' version 4.19
' wersja kolejna ktora uwzlgednia zmiany nazw duzo dokladniej! nie ma bledu jesli szybko bedziemy przesakakiwac pomiedzy akruszami

' version 4.20
' wersja ktora uwzglednia poprawke jesli chodzi o workbook open czyli uruchomienie procedury sprawdzenia pod zmiany w rejestrze ktory arkusz jest aktywny
' w wersji 4.19 zmienilem procedure dodajac jeden argument wiecej


' version 4.21 i 22
' wlasciwie powinno byc 4.21 ale juz z rozpedu dalem 4.22 i niech tak zostanie
' usunalem guziki z arkusza parts - niech juz nie irytuja swoim wadliwym dzialaniem :D


' versrion 4.23
' naprawia blad podwojnego zliczania CBAL :(

' version 4.24
' podczas grzebania w naprawianiu CBALa cos poszlo nie tak z auto component filling :P
' druga sprawa ze kolory z automatu nie chca sie zmieniac - najlepsz jest to ze moj wlasny warunek to blokuje :D


' version 4.25 change on format dates
' this was a huge one
' my polish yyyy-mm-dd
' spanish dd/mm/yyyy
' us mm/dd/yyyy

' version on 4.26 change on fst runout logic switching between years important



' version 4.31 jako ostatnia wersja no i troche dziury bylo
' no i teraz jakis problem sie pojawil z KB ni wiadomo czemu?

' version 4.32 zmiana w module Coverage w sub hit_F8_until_you_see_2000_on_sched_point
' poniewaz w 4.31 byl do loop wiec przynajmniej raz kliklo F8, a powinno najpierw sprawdzic czy juz mamy wartosc 2000
' a potem dopiero klikac na wariata
' dodatkowo dorzucona implementacja sprawdzania podwojnie nazewnictwa nowego akursza
' nadane doswiadczeniem ze moze sie zdarzyc ze usuwamy pojedynczy arkusz i robi sie gap
' a potem dodajemy post lub prefix nazwy i sie sypie poniewaz "name is already taken"
' dodatkowo jest implementacja kontrolowania nazewnictwa arkuszy - bardzo ladna z reszta implementacja oparta na rekurencji z potrojnym wykorzystaniem tego samego nazewnictwa
' zmiennej lokalnej nowej, starej jak i samej nazwy funkcji :D


' version 4.33
' it's all about KB & ZC and cliking F8 to have sched_point 2000


' version 4.34
' adjust to ted request to have at the end have ss on sea and A on air


' narazie tutja lece z plugin orange 0.4
' jesli wszytko pojdzie tak jak trzeba to zakoncze temat z version 4.35 - jednak narazie testy!

' ===================================================================================================================
' version 4.35
' dodanie orange feature
' dodatnie MISC & QHD

' ===================================================================================================================
'
' version 4.40
' Paula - kosmetyczne zmiany zwiazane z przeliczeniem akurszy, co by sie nie chwytal innych plikow
' gdy coverage jest akytwny
' kaizen praktyki studenckiej
' qhd od alc - KB
' TC module usuniety - przesuniecie konfiguracji do initForm
'
' ===================================================================================================================
' version 4.41
' przygotowanie pod kolejna podwersje
' przeliczanie akurszy gotowe
' add schedules z uwzglednieniem curr week zrobione
' qhd dla kb wciaz do zrobienia
' TC module delete - zrobione
' copy  frames - misc & qhd
' debug:
' pojawil sie blad w czesci orange
' moze sie okazac ze ktos zrobil literowke
' na przyklad w miesiacu wpisanym w komentarzu - data portu!
' stad musi byc zewnetrzna validacja wraz z funkcja z MGOMod
' gdzie jest MGOMonth to number!
' dodatkowo w select case znajduje sie else z minus jeden
' done


' ===================================================================================================================
' version 4.42
' to be:
' - paulina add schedules z aktualnym tydniem
' - kasia dreja qhd dla kb
' - blaszczyk idea
' - 3 ideas ewas (prognoza) - to jest masakra :D
' ===================================================================================================================


' ===================================================================================================================
' version 4.43
' to be:
' - ekran pod NSR issue - gotowe!
' - blaszczyk idea - calendar week 38 deadline to do it a mamy 10.09.2015 dalej nie zrobilem to be done on 4.44
' - prio idea container manager (add plt lists) - calendar week 39 deadline :)
' - juz wiem ze teraz robie proj dla to be done on 4.44
' fma i napewno nie zdaze narazie wdrozyc wersji 4.43 - tak potwierdzam dalej stoje z robota 12/10/2015 - wszsytko to
' be done w wersji 4.44

' teraz musze sie zajac hot issue with iso week num
' ===================================================================================================================

' ===================================================================================================================
' version 4.44
' to be:
' - ekran pod NSR issue - gotowe!
' - blaszczyk idea - calendar week 38 deadline to do it a mamy 10.09.2015 dalej nie zrobilem to be done on 4.44
' - prio idea container manager (add plt lists) - calendar week 39 deadline :)
' - juz wiem ze teraz robie proj dla to be done on 4.x - nawet nie wiem do konca ile to sie jeszcze zmieni mini jobow
' fma i napewno nie zdaze narazie wdrozyc wersji 4.44 - tak potwierdzam dalej stoje z robota 12/10/2015 - wszsytko to
' be done w wersji 4.45

' zle liczyl dla daily rqms - trzeba bylo zmieniac logike sprawdzania danych i ewentualnego przesuwania lat
' ===================================================================================================================


' ===================================================================================================================
' version 4.45
' to be:
' - ekran pod NSR issue - gotowe!
' - blaszczyk idea - calendar week 38 deadline to do it a mamy 10.09.2015 dalej nie zrobilem to be done on 4.44
' - prio idea container manager (add plt lists) - calendar week 39 deadline :)
' - juz wiem ze teraz robie proj dla to be done on 4.x - nawet nie wiem do konca ile to sie jeszcze zmieni mini jobow
' fma i napewno nie zdaze narazie wdrozyc wersji 4.44 - tak potwierdzam dalej stoje z robota 12/10/2015 - wszsytko to
' be done w wersji 4.45

' teraz dodatek w postaci usuwania wybranych asnow - dodatkowy guzik - removie asn0
' zmiany wlasciwie tylko w scenario module


' version 4.46 asn offset

' ===================================================================================================================


' version 4.47
' - zly warunek brzegowy dla not yet received - nie uwzglednial zmiany roku i wyrzucal stare not yet received

' ===================================================================================================================


' ===================================================================================================================


' version 4.48
' - dalej jest cos nie tak tym razem dla ASN w terminach od 1st do 4th of january
' - offset colors in parts sheet control
' - check if time frame in cov can be dynamic
'

' ===================================================================================================================


' ===================================================================================================================


' version 4.49 dev
' - dalej jest cos nie tak tym razem dla ASN w terminach od 1st do 4th of january
' - offset colors in parts sheet control - tego nie chce juz tego zmieniac
' - check if time frame in cov can be dynamic - to jest raczej trudne do zrealizowania
' - plt z mozliwoscia GME
' - podswietlenie dla Pauliny backlogu

' ===================================================================================================================

' ===================================================================================================================


' version 4.5x dev
' - slepa uliczka niezaleznego MGO

' ===================================================================================================================




' ===================================================================================================================


' version 4.6x dev
' - problem w PO z IP - zatem, aby uniknac wszelkiego rodzaju IPkow w ogole wrzucamy je wszystie do jednego wora
' - druga opcja to checkbox dzieki ktoremu zaciagamy CBAL z POF dzieki czemu widzimy ile jest w ASLu
' - poprawione ramki czarne na dole pojawiajace sie bez PNu
' - formatowanie pod CBALu
' - 4.62 weryfikacja poprawnosci bankow srednia powyzej kreski ttime'u - paleta kolorow
' - kreski ttimeu zmieniaja sie w zaleznosci od zmiany godzinowego TTIME zmienia ramke i przeliczone dni
' od 4.63:
' - kolorowanie banku ma byc opcjonalne
' - update kolorow na bankach
' - jesli rqms avr = 0 to wtedy nie kolorujemy bankow w ogole
' - dwa dodatkowe guziki poniewaz ustawianie formatowanie jest bezwlasne i pozostaje ustawienie rowniez po wygenerowaniu
' raportu weekly

' ===================================================================================================================


' ===================================================================================================================


' version 4.64 dev
' - misc z weekly i daily screen - mozliwosc usuniecia

' ===================================================================================================================

' ===================================================================================================================


' version 4.65 dev
' - poprawka na logice bankow i kreski ttime'u
' - poprawka Pauliny pod nieprawidlowe przeliczania rodzaju cov
' (po usuwaniu std zamienial sie w green lub blue and green jesli takowy byl aktywny)

' ===================================================================================================================


' version 4.66 dev
' - dodanie opcji clear dla listy wsadowej czesci - Paulina
' ===================================================================================================================


' version 4.67 dev
' - fix na czesci implementacji zwiazanej z wyodrebnianiem z zk7pwrqm itemow MISC*OTHR*
' - brakowalo warunku if dla danych ktore pomimo zarzadzenia ze chce widziec misc nie pokazuje nic zamiast "odjac zero"
' ===================================================================================================================



' version 4.68 dev
' - fix na transportach IP - oprocz asnow z mode T mamy rowniez od teraz akceptowane czesci z brakiem mode at all
' klasa TworzenieTranzytow linia 69.
' ===================================================================================================================
