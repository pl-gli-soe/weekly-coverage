VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PivotHandler"
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


Private pivot_cache As PivotCache
Private pivot_table As PivotTable
Private source_range As Range
Private pivot_sh As Worksheet
Private slajser_cache1 As SlicerCache
Private slajser_cache2 As SlicerCache


Public Sub init(e As pivot_layout)
    
    
    Set pivot_cache = ThisWorkbook.PivotCaches.Create(xlDatabase, source_range)
    Dim Sh As Worksheet
    Set Sh = ThisWorkbook.ActiveSheet
    
    If e = EDA Then
        tmp_str = Replace(Sh.Name, "FLAT", "PIVOTD")
    ElseIf e = EDACW Then
        tmp_str = Replace(Sh.Name, "FLAT", "PIVOTW")
    End If
    Set pivot_sh = ThisWorkbook.Sheets.Add
    pivot_sh.Name = tmp_str
    tmp_routing = Sh.Range("a1")
    pivot_sh.Range("a1") = Replace(tmp_routing, "PLT ", "")
    
    Set pivot_table = pivot_sh.PivotTables.Add(PivotCache:=pivot_cache, TableDestination:=Range("B4"))
    
End Sub

Public Sub config_pivot()


    Dim pivot_field As PivotField
    
    With pivot_table
        .PivotFields("CONTAINER").Orientation = xlRowField
        .PivotFields("CONTAINER").Position = 1
        
        .PivotFields("PN").Orientation = xlRowField
        .PivotFields("PN").Position = 2
        
        
        
        .PivotFields("FIRST RUNOUT").Orientation = xlColumnField
        .PivotFields("FIRST RUNOUT").Position = 1
        
        
        Set pivot_field = .PivotFields("PN")
        pivot_field.Orientation = xlDataField
        pivot_field.Function = xlCount
        
        .PivotFields("CONTAINER").ShowDetail = False
        
        
        .PivotFields("MODE").Orientation = xlPageField
        .PivotFields("MODE").Position = 1
        .PivotFields("ST").Orientation = xlPageField
        .PivotFields("ST").Position = 2
        
        
        .PivotFields("ST").ClearAllFilters
        On Error Resume Next
        .PivotFields("ST").CurrentPage = "2"
        
        With .PivotFields("MODE")
            .ClearAllFilters
            On Error Resume Next
            .PivotItems("AC").Visible = False
            On Error Resume Next
            .PivotItems("AE").Visible = False
            On Error Resume Next
            .PivotItems("M").Visible = False
            
            
            .EnableMultiplePageItems = True
            .CurrentPage = "(All)"
            
            ' psuedo code
            ' .CurrentPage = "SS", "VE"
        End With
        
        .ColumnGrand = False
        .RowGrand = False
        
        .TableStyle2 = "PivotStyleMedium15"
    End With
End Sub

Public Sub add_slicers(e As pivot_layout)

    Set slajser_cache1 = Nothing
    Set slajser_cache2 = Nothing
    
    ' tutaj ino troche sliska sprawa z tymi on error resume next
    ' zabopiega oczywiscie pojawieniu sie bledu jakiegokolwiek zwiazanego z slicerami
    ' i jedyne co sie wtedy stanie to brak na pivocie naszych slicerow
    ' hmmmm ale na dluzsza mete trzeba bedzie o tym jeszcze pomyslec jak ewentualne inne bledy zalatwic
    ' bo jednak jest to nieeleganckie rozwiazanie
    ' 2014-07-01
    If e = EDACW Then
        On Error Resume Next
        Set slajser_cache1 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "EDA CW")
        slajser_cache1.Slicers.Add pivot_sh, , "EDA CW" & CStr(pivot_sh.Name), "EDA CW Slicer", 100, 400, 150, 200
        
        On Error Resume Next
        Set slajser_cache2 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "URGENCY")
        slajser_cache2.Slicers.Add pivot_sh, , "URGENCY" & CStr(pivot_sh.Name), "URGENCY Slicer", 140, 500, 500, 200
        
    ElseIf e = EDA Then
        On Error Resume Next
        Set slajser_cache1 = ThisWorkbook.SlicerCaches.Add2(pivot_table, "EDA")
        slajser_cache1.Slicers.Add pivot_sh, , "EDA" & CStr(pivot_sh.Name), "EDA Date Slicer", 100, 400, 150, 200
    End If
    
    
    
End Sub




Public Sub if_flat_table_prepare_source_range()

    ' metoda prosta i klarowna
    ' aktywnym sheetem musi byc sheet flat table
    ' metoda sprawdzjaca czy faktycznie tak jest
    ' to check_labels
    
    Dim Sh As Worksheet
    Set Sh = ActiveSheet
    
    If check_labels(Sh) Then
        ' dziala zatem krotko zrob nowy pivot
        prepare_source_range Sh
    Else
        MsgBox "only flat table can be source for PRIO pivot table"
    End If

    
End Sub

Private Sub prepare_source_range(Sh As Worksheet)
    
    Set source_range = Nothing
    
    Dim r As Range
    Set r = Sh.Range("a1")
    Set r = Range(r, r.End(xlDown))
    Set r = Range(r, r.End(xlToRight))
    
    Set source_range = r
End Sub

