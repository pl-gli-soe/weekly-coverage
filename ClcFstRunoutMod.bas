Attribute VB_Name = "ClcFstRunoutMod"
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


Public Function calcFirstRunOut(r As Range)


    rok = Int(Year(Date)) * 100
    calcFirstRunOut = ""

    Dim rr As Range
    Set rr = r.offset(-5, 0)
    If rr.item(1) < rr.item(r.Count) Then
        For Each i In r
            If i < 0 Then
                calcFirstRunOut = rok + i.offset(-5, 0)
                Exit Function
            End If
        Next i
    Else
        ' tutaj dodatkowo dochodzi opcja ze mamy przejscie przez nowy rok i mamy zalamanie ciaglosci danych jesli
        ' chodzi tylko i wylacznie o czysty CW
        ' zatem musi algorytm w szybki i prosty sposob umiec to rozpoznac
        For Each i In r
            If i < 0 Then
                If i.offset(-5, 0) >= rr.item(1) Then
                    calcFirstRunOut = rok + i.offset(-5, 0)
                    Exit Function
                Else
                    rok = rok + 100
                    calcFirstRunOut = rok + i.offset(-5, 0)
                    Exit Function
                End If
            End If
        Next i
    End If

    If rr.item(1) < rr.item(r.Count) Then
        calcFirstRunOut = rok + r.item(r.Count).offset(-5, 0)
    Else
        rok = rok + 100
        calcFirstRunOut = rok + r.item(r.Count).offset(-5, 0)
    End If
    
    
End Function
