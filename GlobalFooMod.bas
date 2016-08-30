Attribute VB_Name = "GlobalFooMod"
' funkcja wykorzystywana przez klase Pivot oraz przez modul catchCONT
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


Public Function check_labels(ByRef Sh As Worksheet) As Boolean

    check_labels = False

    ' obsolete
    ' If Sh.Name Like "FLAT*" Then
    On Error Resume Next
        If Sh.Range("a1").Value Like "PLT *" Then
            If Sh.Range("b1").Value = "PN" Then
                If Sh.Range("c1").Value = "DUNS" Then
                    If Sh.Range("d1").Value = "FIRST RUNOUT" Then
                        check_labels = True
                        Exit Function
                    End If
                End If
            End If
        End If
    ' End If
    
End Function


Public Sub sprawdz_czy_nazwa_zostala_zmieniona(s As Worksheet)

    Set zmiana_nazw_powiazanych_arkuszy = Nothing
    Set zmiana_nazw_powiazanych_arkuszy = New EventZmienNazwyPowiazanych
    zmiana_nazw_powiazanych_arkuszy.zdarzenie.sprawdzCzyNazwaArkuszaZostalaZmieniona s
End Sub
