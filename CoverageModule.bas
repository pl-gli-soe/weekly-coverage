Attribute VB_Name = "CoverageModule"
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
'



Function OstatniaNiedziela(Data)
    OstatniaNiedziela = Data - Weekday(Data) + 1
End Function

Public Function IsoWeekNumber(InDate As Date) As Integer
    Dim d As Long
    d = DateSerial(Year(InDate - Weekday(InDate - 1) + 4), 1, 3)
    IsoWeekNumber = Int((InDate - d + Weekday(d) + 5) / 7)
End Function

Sub MakeCoverage(rrow, inPLT, inPRT, Comment, Component, kolor)
    Dim commentHeader As String
    Dim alc_sheet As Worksheet
    mgoInit
    
    Set s = this_workbook.ActiveSheet
    
    CoverCols = 20
    
    s.Cells(rrow, 1).Value = "PLT"
    s.Cells(rrow, 9).Value = "Past due"
    s.Cells(rrow, 7).Value = "First runout"
    ' tutaj specjalne miesjce pod ruchomy first runout
    ' i chyba powinno to byc opcjonalne :D
    ' ========================================================
    s.Cells(rrow, 7).Value = "First runout"
    s.Cells(rrow + 1, 7).FormulaR1C1 = "=calcFirstRunOut(R[4]C[3]:R[4]C[22])"
    s.Cells(rrow + 1, 7).Font.Bold = True
    ' ========================================================
    
    's.Cells(rrow + 2, 7).FormulaR1C1 = "=R[-1]C"
    's.Cells(rrow + 3, 7).FormulaR1C1 = "=R[-1]C"
    's.Cells(rrow + 4, 7).FormulaR1C1 = "=R[-1]C"
    's.Cells(rrow + 5, 7).FormulaR1C1 = "=R[-1]C"
    
    
    s.Cells(rrow, 8).Value = Component
    
    s.Cells(rrow + 1, 1).Value = inPLT
    s.Cells(rrow + 1, 2).Value = inPRT
    s.Cells(rrow + 4, 2).Value = Comment
    s.Cells(rrow + 4, 2).Interior.Color = kolor
    
    s.Cells(rrow + 1, 8).Value = "req"
    s.Cells(rrow + 2, 8).Value = "Vessel"
    s.Cells(rrow + 3, 8).Value = "AIR"
    s.Cells(rrow + 4, 8).Value = "manual"
    s.Cells(rrow + 5, 8).Value = "coverage"
    
    s.Cells(rrow + 1, 3).Value = "BANK"
    s.Cells(rrow + 2, 3).Value = "MODE"
    s.Cells(rrow + 3, 3).Value = "PCS TO GO"
    s.Cells(rrow + 4, 5).Value = "TTIME"
    s.Cells(rrow + 5, 3).Value = "MISC"
    
    s.Cells(rrow + 1, 5).Value = "CBAL"
    s.Cells(rrow + 2, 5).Value = "DOH"
    s.Cells(rrow + 3, 5).Value = "QHD"
    s.Cells(rrow + 4, 3).Value = "STD PACK"
    s.Cells(rrow + 5, 5).Value = "BACKLOG"
    
    

    Dim cwr As Range
    ' Najpierw wpisz daty
    For x = 0 To CoverCols - 1
        s.Cells(rrow, x + 10).Value = IsoWeekNumber(Round(OstatniaNiedziela(Date)) + x * 7 + 1)
        ' DodajKomentarz rrow, x + 10, Format(Round(OstatniaNiedziela(Date)) + x * 7 + 1, "dd-mm-yyyy") + " - " + Format(Round(OstatniaNiedziela(Date)) + x * 7 + 7, "dd-mm-yyyy"), 150
        
        Set cwr = s.Cells(rrow, x + 10)
        cwr.AddComment CStr(Format(Round(OstatniaNiedziela(Date)) + x * 7 + 1, "dd-mm-yyyy") + " - " + Format(Round(OstatniaNiedziela(Date)) + x * 7 + 7, "dd-mm-yyyy"))
        cwr.Comment.Shape.Width = 150
        cwr.Comment.Shape.Height = 30
    Next x
    ' teraz formu³y
    s.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C"
    
    If Comment = "ALC" Then
        s.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C+RC[-6]"
    End If
    
    For x = 1 To CoverCols - 1
        s.Cells(rrow + 5, 10 + x).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C"
    Next x
    
    ' weekly reqs
    StatusBox.Description.Caption = "Weekly reqms " + CStr(inPLT) + " " + CStr(inPRT)
    
    
    If Comment = "ALC" Then
        Set alc_sheet = Nothing
        On Error Resume Next
        Set alc_sheet = Workbooks("ALC.xlsx").Sheets("Sheet1")
        If Not alc_sheet Is Nothing Then
            s.Cells(rrow + 5, 4).Value = take_data_from_alc(alc_sheet, CStr(inPLT), CStr(inPRT), Int(1))
            s.Cells(rrow + 3, 6).Value = take_data_from_alc(alc_sheet, CStr(inPLT), CStr(inPRT), Int(2))
        Else
            MsgBox "PLEASE OPEN ALC.xslx"
        End If
    End If
    
    If Component = "C" Then
        Sess0.Screen.sendKeys ("<Clear>")
        waitForMgo
        Sess0.Screen.sendKeys ("MS7P3100 <Enter>")
        waitForMgo
        ProgressIncrease
        Sess0.Screen.putString inPLT, 1, 21
        Sess0.Screen.putString inPRT, 6, 13
        
        Sess0.Screen.sendKeys ("<Enter>")
        waitForMgo
        If (inPLT <> "KB") And (inPLT <> "ZC") Then
            Dept = Sess0.Screen.getString(6, 57, 4)
            Sess0.Screen.sendKeys ("<Clear>")
            waitForMgo
            Sess0.Screen.sendKeys ("MS8P1600 <Enter>")
            waitForMgo
            Sess0.Screen.putString inPLT, 1, 21
            Sess0.Screen.putString inPRT, 4, 8
            Sess0.Screen.putString Dept, 4, 50
            Sess0.Screen.sendKeys ("<Enter>")
            waitForMgo
            actcol = 11
            ' pierwsze reqs, od 2giego tygodnia
            For y = 11 To 19
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(y, 14, 10))
                actcol = actcol + 1
            Next y
            For y = 10 To 19
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(y, 35, 10))
                actcol = actcol + 1
            Next y
       
        
        ElseIf (inPLT = "ZC") Then
    
            hit_F8_until_you_see_2000_on_sched_point
        
            Sess0.Screen.sendKeys ("<pf6>")
            waitForMgo
            Sess0.Screen.sendKeys ("<pf5>")
            waitForMgo
            
            actcol = 10
            For y = 8 To 20
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(y, 51, 9))
                actcol = actcol + 1
            Next y
            
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
            For y = 8 To 14
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(y, 51, 9))
                actcol = actcol + 1
            Next y
            
         ElseIf (inPLT = "KB") Then
         
            put_zero_on_BLD_SEQ

            Sess0.Screen.sendKeys ("<pf6>")
            waitForMgo
            Sess0.Screen.sendKeys ("<pf5>")
            waitForMgo
            
            actcol = 10
            For y = 8 To 20
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(y, 51, 9))
                actcol = actcol + 1
            Next y
            
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
            For y = 8 To 14
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(y, 51, 9))
                actcol = actcol + 1
            Next y
        End If
    Else
        Sess0.Screen.sendKeys ("<Clear>")
        waitForMgo
        Sess0.Screen.sendKeys ("ZK7PWRQM <Enter>")
        waitForMgo
          
        ProgressIncrease
            
        Sess0.Screen.putString "  ", 3, 8
        Sess0.Screen.putString "        ", 4, 11
        Sess0.Screen.putString "    ", 5, 10
        
        Sess0.Screen.putString inPLT, 3, 8
        Sess0.Screen.putString inPRT, 4, 11
        
        Sess0.Screen.sendKeys ("<Enter>")
        waitForMgo
        
        
        miscLine = -1
        If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("MISC") = 0 Then
            For y = 10 To 21
                
                If Trim(Sess0.Screen.getString(y, 2, 9)) Like "*MISC*OTHR*" Then
                    miscLine = y
                End If
            Next y
        End If
        
        actcol = 11
        ' pierwsze reqs, od 2giego tygodnia
        For y = 1 To 4
            If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("MISC") = 1 Then
                s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(9, 22 + 8 * y, 8))
            Else
            
                If miscLine <> -1 Then
                    s.Cells(rrow + 1, actcol).Value = _
                        Val(Sess0.Screen.getString(9, 22 + 8 * y, 8)) - Val(Sess0.Screen.getString(miscLine, 22 + 8 * y, 8))
                Else
                    s.Cells(rrow + 1, actcol).Value = _
                        Val(Sess0.Screen.getString(9, 22 + 8 * y, 8))
                End If
            End If
            actcol = actcol + 1
        Next y
        
        ' i pozosta³e reqs
        Do
            Sess0.Screen.sendKeys ("<Pf11>")
            waitForMgo
            
            For y = 0 To 4
                
                If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("MISC") = 1 Then
                    s.Cells(rrow + 1, actcol).Value = Val(Sess0.Screen.getString(9, 22 + 8 * y, 8))
                Else
                    If miscLine <> -1 Then
                        s.Cells(rrow + 1, actcol).Value = _
                            Val(Sess0.Screen.getString(9, 22 + 8 * y, 8)) - Val(Sess0.Screen.getString(miscLine, 22 + 8 * y, 8))
                    Else
                        s.Cells(rrow + 1, actcol).Value = _
                            Val(Sess0.Screen.getString(9, 22 + 8 * y, 8))
                    End If
                End If
                actcol = actcol + 1
                If actcol > 9 + CoverCols Then Exit Do
                
            Next y
        Loop Until False
    
        ' zsumowanie daily reqs
        StatusBox.Description.Caption = "Daily rqm " + CStr(inPLT) + " " + CStr(inPRT)
        LimitDate = OstatniaNiedziela(Date) + 7   ' kolejna niedziela
        Sess0.Screen.sendKeys ("<Clear>")
        waitForMgo
        Sess0.Screen.sendKeys ("ZK7PDRQM <Enter>")
        waitForMgo
          
        ProgressIncrease
            
        Sess0.Screen.putString "  ", 3, 8
        Sess0.Screen.putString "        ", 4, 8
        Sess0.Screen.putString "    ", 5, 10
        
        Sess0.Screen.putString inPLT, 3, 8
        Sess0.Screen.putString inPRT, 4, 8
            
        Sess0.Screen.sendKeys ("<Enter>")
        waitForMgo
        
            
        DayReqs = 0
        miscValue = 0
        
        While (Sess0.Screen.getString(22, 2, 5) = "R6693")
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
        Wend
            
        ' ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("MISC") = 0
        miscLine = -1
        PltTotLine = 9
        For q = 9 To 20
        
            If Sess0.Screen.getString(q, 2, 9) Like "*MISC*" Then
                miscLine = q
            End If
            
            
            If Sess0.Screen.getString(q, 2, 9) = "PLT TOTAL" Then
                PltTotLine = q
            End If
        Next q
            
        For y = 0 To 4
            MDay = Val(Sess0.Screen.getString(8, 24 + 8 * y, 2))
            MMonth = MgoMonth(Sess0.Screen.getString(8, 27 + 8 * y, 2))
            If Month(Now) = 12 And MMonth = 1 Then
                MYear = Year(Now) + 1
            Else
                MYear = Year(Now)
            End If
            mDate = DateSerial(MYear, MMonth, MDay)
            'MsgBox (MDate)
            If mDate <= LimitDate Then ' sumuj do niedzieli w³¹cznie
                DayReqs = DayReqs + Val(Sess0.Screen.getString(PltTotLine, 22 + 8 * y, 8))
                'MsgBox (Val(Sess0.Screen.GetString(PltTotLine, 22 + 8 * y, 8)))
                'MsgBox (DayReqs)
                
                If miscLine <> -1 Then
                    miscValue = miscValue + Val(Sess0.Screen.getString(miscLine, 22 + 8 * y, 8))
                End If
            End If
        Next y
            
        ' quick & dirty hack
        ' powtorzenie powy¿szych pêtli dla kolejnego ekranu - przydatne przy pracuj¹cych sobotach
        
        Sess0.Screen.sendKeys ("<Enter>")
        waitForMgo
        Sess0.Screen.sendKeys ("<Pf11>")
        waitForMgo
        
        While (Sess0.Screen.getString(22, 2, 5) = "R6693")
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
        Wend
        

        miscLine = -1
        PltTotLine = 9
        For q = 9 To 20
        
        
            If Sess0.Screen.getString(q, 2, 9) Like "*MISC*OTHR*" Then
                miscLine = q
            End If
        
            If Sess0.Screen.getString(q, 2, 9) = "PLT TOTAL" Then
                PltTotLine = q
            End If
        Next q
            
        For y = 0 To 4
            MDay = Val(Sess0.Screen.getString(8, 24 + 8 * y, 2))
            MMonth = MgoMonth(Sess0.Screen.getString(8, 27 + 8 * y, 2))
            If Month(Now) > MMonth Then
                MYear = Year(Now) + 1
            Else
                MYear = Year(Now)
            End If
            mDate = DateSerial(MYear, MMonth, MDay)
            'MsgBox (MDate)
            If mDate <= LimitDate Then ' sumuj do niedzieli w³¹cznie
                DayReqs = DayReqs + Val(Sess0.Screen.getString(PltTotLine, 22 + 8 * y, 8))
                'MsgBox (Val(Sess0.Screen.GetString(PltTotLine, 22 + 8 * y, 8)))
                'MsgBox (DayReqs)
                
                If miscLine <> -1 Then
                    miscValue = miscValue + Val(Sess0.Screen.getString(miscLine, 22 + 8 * y, 8))
                End If
            End If
        Next y
        
        If ThisWorkbook.Sheets(COV.REGISTER_SH_NM).Range("MISC") = 0 Then
        
            s.Cells(rrow + 1, 10).Value = DayReqs - miscValue
        Else
            s.Cells(rrow + 1, 10).Value = DayReqs
        End If
            
    End If
            
    ' wypiswanie begining balance
    StatusBox.Description.Caption = "Beg balance " + CStr(inPLT) + " " + CStr(inPRT)
    
    cbal = ""
    asl = ""
    vcbal = 0
    vasl = 0
    
    
    If InitForm.CheckBoxCBALFromPOF.Value = True Then
    
    
        Sess0.Screen.sendKeys ("<Clear>")
        waitForMgo
        Sess0.Screen.sendKeys ("ms9pof00 <Enter>")
        waitForMgo
        
        Sess0.Screen.putString "6", 2, 18
        
        Sess0.Screen.putString "  ", 3, 13
        Sess0.Screen.putString "        ", 3, 21
        Sess0.Screen.putString inPLT, 3, 13
        Sess0.Screen.putString inPRT, 3, 21
        
        Sess0.Screen.sendKeys ("<Enter>")
        waitForMgo
        
        If Sess0.Screen.getString(5, 38, 1) = "-" Then
            cbal = "-" + Trim(Sess0.Screen.getString(5, 25, 13))
        Else
            cbal = Trim(Sess0.Screen.getString(5, 25, 13))
        End If
        
        If Sess0.Screen.getString(5, 51, 1) = "-" Then
            asl = "-" + Trim(Sess0.Screen.getString(5, 44, 7))
        Else
            asl = Trim(Sess0.Screen.getString(5, 44, 7))
        End If
        
        
        If cbal = "" Then
            cbal = "0"
        End If
        
        If asl = "" Then
            asl = "0"
        End If
        
        
        
        vcbal = Int(cbal)
        vasl = Int(asl)
        
        If inPLT = "ZC" Then
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
            
            
            If Sess0.Screen.getString(5, 51, 1) = "-" Then
                cbal = "-" + Sess0.Screen.getString(5, 25, 13)
            Else
                cbal = Sess0.Screen.getString(5, 25, 13)
            End If
            
            If Trim(cbal) = "" Then
                cbal = "0"
            End If
            
            
            ' wartosci sumowana miedzy jednym DGO a drugim
            vcbal = vcbal + Int(cbal)
            vasl = vasl + Int(asl)
        End If
        
    End If
    
    
    Sess0.Screen.sendKeys ("<Clear>")
    waitForMgo
    Sess0.Screen.sendKeys ("ms9pop00 <Enter>")
    waitForMgo
      
    Sess0.Screen.putString "6", 2, 18

    ProgressIncrease
    
    
        
    Sess0.Screen.putString "  ", 3, 13
    Sess0.Screen.putString "        ", 3, 22
    Sess0.Screen.putString inPLT, 3, 13
    Sess0.Screen.putString inPRT, 3, 22

    
    
    Sess0.Screen.sendKeys ("<Enter>")
    waitForMgo
    
    If InitForm.CheckBoxCBALFromPOF.Value = False Then
        If Sess0.Screen.getString(5, 30, 1) = "-" Then
            cbal = "-" + Sess0.Screen.getString(5, 22, 8)
        Else
            cbal = Sess0.Screen.getString(5, 22, 8)
        End If
    
    
        If Trim(cbal) = "" Then
            cbal = "0"
        End If
        
        If Trim(asl) = "" Then
            asl = "0"
        End If
        
        
        
        vcbal = Int(cbal)
        vasl = Int(asl)
    
        If inPLT = "ZC" Then
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
            
            
             If Sess0.Screen.getString(5, 30, 1) = "-" Then
                cbal = "-" + Sess0.Screen.getString(5, 22, 8)
            Else
                cbal = Sess0.Screen.getString(5, 22, 8)
            End If
            
            If Trim(cbal) = "" Then
                cbal = "0"
            End If
            
            
            ' wartosci sumowana miedzy jednym DGO a drugim
            vcbal = vcbal + Int(cbal)
            
            
        End If
    End If
    
   
    
    
    
    
    
    ' Debug.Assert CStr(inPRT) <> "13337293"
    
    If Sess0.Screen.getString(6, 73, 1) = "-" Then
        OS = "-" + Trim(Sess0.Screen.getString(6, 67, 6))
    Else
        OS = Trim(Sess0.Screen.getString(6, 67, 6))
    End If
            
    Rkod = Sess0.Screen.getString(23, 2, 5)
    If Rkod <> "     " And Rkod <> "I4767" Then Debug.Print CStr(PlantCode) + " " + CStr(PartNumber) + " " + Sess0.Screen.getString(23, 2, 80)
            
    ' wyjatkowo formatowanie damy od razu przy wpisie
    ' ==================================================
    s.Cells(rrow + 1, 6).NumberFormat = "0_ ;[Red]-0 "
    s.Cells(rrow + 1, 6).Font.Bold = True
    ' ==================================================
    s.Cells(rrow + 1, 6).Value = vcbal
    s.Cells(rrow + 2, 2).Value = Sess0.Screen.getString(7, 17, 15)
    s.Cells(rrow + 3, 2).Value = Sess0.Screen.getString(7, 7, 9)
    s.Cells(rrow + 2, 6).Value = Sess0.Screen.getString(4, 43, 3)
    
    
    ' let check if Global variable want it
    's.Cells(rrow + 5, 3).Value = "MISC"
    's.Cells(rrow + 3, 5).Value = "QHD"
    ' QHD i MISC
    
    ' logika z ALC (tylko dla KB)
    ' s.Cells(rrow + 5, 4).Value = take_data_from_alc(alc_sheet, CStr(inPLT), CStr(inPRT), Int(1))
    ' s.Cells(rrow + 3, 6).Value = take_data_from_alc(alc_sheet, CStr(inPLT), CStr(inPRT), Int(2))
    ' MISC
    ' tu jest dodawanie samej siebie poniewaz misc przechowuje rowniez wartosci z pliku alc!
    s.Cells(rrow + 5, 4).Value = s.Cells(rrow + 5, 4).Value + CLng(Trim(Sess0.Screen.getString(5, 73, 7)))
    
    ' QHD
    ' tu jest dodawanie samej siebie poniewaz qhd przechowuje rowniez wartosci z pliku alc!
    s.Cells(rrow + 3, 6).Value = s.Cells(rrow + 3, 6).Value + CLng(Trim(Sess0.Screen.getString(5, 48, 7)))
    
    If CBool(G_QHD_MISC_BACKLOG) Then
    
    
        Dim qim As Range
        ' misc
        If s.Cells(rrow + 5, 4).Value <> 0 Then
            Set qim = s.Cells(rrow + 5, 4)
            qim.Font.Color = RGB(200, 10, 20)
        End If
        
        ' and niezaleznie
        
        ' qhd
        If s.Cells(rrow + 3, 6).Value <> 0 Then
            Set qim = s.Cells(rrow + 3, 6)
            qim.Font.Color = RGB(200, 10, 20)
        End If
        
        
        If CStr(OS) <> "0" Then
            If CLng(OS) < 0 Then
                s.Cells(rrow + 5, 6).Font.Color = RGB(250, 190, 0)
            Else
                s.Cells(rrow + 5, 6).Font.Color = RGB(0, 176, 240)
            End If
        End If
        
    End If
    
    
    ' std pack zamieniony z ttime
    s.Cells(rrow + 4, 4).Value = Sess0.Screen.getString(10, 30, 9)
    
    s.Cells(rrow + 5, 6).Value = OS
    
    s.Cells(rrow + 1, 4).Value = Trim(Sess0.Screen.getString(6, 54, 7))
    s.Cells(rrow + 2, 4).Value = Trim(Sess0.Screen.getString(8, 21, 2))
    s.Cells(rrow + 3, 4).Value = Trim(Sess0.Screen.getString(6, 23, 7))
    ttime = Trim(Sess0.Screen.getString(8, 31, 8))
    s.Cells(rrow + 4, 6).Value = ttime
    s.Cells(rrow + 3, 7).Value = "in days"
    
    s.Cells(rrow + 4, 7).Value = Application.WorksheetFunction.Round((CDbl(Val(ttime) / 24#) / CDbl(ThisWorkbook.Sheets("register").Range("tc"))) * 7, 2)
    ' s.Cells(rrow + 5, 4).Value = Trim(sess0.Screen.GetString(5, 73, 7))
    
        
    
    ' zsumowanie daily reqs dla plantu PC
    If Component = "C" Then
        StatusBox.Description.Caption = "Daily rqm " + CStr(inPLT) + " " + CStr(inPRT)
        LimitDate = OstatniaNiedziela(Date) + 7   ' kolejna niedziela
        ProgressIncrease
        DayReqs = 0
            
        For y = 0 To 4
            MDay = Val(Sess0.Screen.getString(8 + y, 42, 2))
            MMonth = MgoMonth(Sess0.Screen.getString(8 + y, 40, 2))
            MYear = Year(Now)
            mDate = DateSerial(MYear, MMonth, MDay)
            If mDate <= LimitDate Then ' sumuj do niedzieli w³¹cznie
                DayReqs = DayReqs + Val(Sess0.Screen.getString(8 + y, 45, 10))
            End If
        Next y
        Sess0.Screen.sendKeys ("<Pf11>")
        waitForMgo
        For y = 0 To 4
            MDay = Val(Sess0.Screen.getString(8 + y, 42, 2))
            MMonth = MgoMonth(Sess0.Screen.getString(8 + y, 40, 2))
            MYear = Year(Now)
            mDate = DateSerial(MYear, MMonth, MDay)
            If mDate <= LimitDate Then ' sumuj do niedzieli w³¹cznie
                DayReqs = DayReqs + Val(Sess0.Screen.getString(8 + y, 45, 10))
            End If
        Next y
        s.Cells(rrow + 1, 10).Value = DayReqs
    End If
    
    
    ' wypiswanie pickupów w tranzycie
    ' element gruntownie do przerobienia w forme OOP
    ' 2014-05-19
    ' ===========================================================================================================================
    ' ===========================================================================================================================
    '
    ' z racji kluczowej informacji o first runout nie jest mozliwe aby uruchomic pobieranie na osobny arkusz
    ' od razu asnow z ms9po400 chyba ze zrobimy to stepami ale nie bedzie to zbyt wygodne
    ' -----------------------------------------------
    ' 1. pierwsza opcja to gruntowne przerobienie implementacji tworzenia opisow
    ' pod koementarze pod asny jaki i elementy niezgodne z statusami jak i modami w komentarzu
    ' - rozwiazanie przodujace bo moze w koncu nada sie zrobic jakies pozadne komentarze na coverage
    ' - z drugiej strony dobrze by bylo nadac rowniez odrebne numerowania makra - staje sie rozbudowane i zeby sie nie skonczylo tak
    ' jak jest aktualnie z fire flakeiem gdzie kupa rzeczy najzwyczjaniej w swiecie spowalnia dzialanie
    ' zatem Mateuszku uwazaj :)
    '
    ' 2. druga opcja jest bardziej liniwa ale pozostawi caly format danych na layoutcie tak samo jak jest dotychczas
    ' i z punktu widzenia userow
    ' oprocz dodatkowej opcji w formularzu wejsciowym reszta jest dokladnie taka sama
    ' moze bedzie nawet mozna bylo dodac nieco dynamiki jesli chodzi o first run out poniewaz
    ' tak za kazdym razem makro musi podliczyc kiedy tam czerwony kolor sie pojawi
    ' nalezy pamietac ze to rozwiazanie musi zostac wykonane w dwoch krokach z elementami ktore potrafia przetrzymac dane gdy juz z ms9po400 zejdziemy
    ' zrobimy layout i dopiero wtedy bedziemy mogli wszystkie dane nalozyc na siebie poniewaz dezajn coverage'a weekly traktuje pierwszy czerwony kolor nie jako dane
    ' ale jako czesc layoutu, ktora nie jest istotna z punktu widzenia implementacji
    ' -----------------------------------------------
    
    
        
    Dim imgo As MGO, tt As TworzenieTranzytow, kolekcja As Collection, ch As Komentarz 'ch : comment handler
    Set imgo = New MGO
    Set tt = New TworzenieTranzytow
    Set ch = New Komentarz
    Set kolekcja = Nothing
    Set kolekcja = New Collection
    tt.on_ms9po400 CStr(inPLT), CStr(inPRT), imgo, kolekcja
    ProgressIncrease
    
    For x = 9 To 9 + CoverCols
        s.Cells(rrow + 2, x).Value = 0
        s.Cells(rrow + 3, x).Value = 0
    Next x
    
    ' one_time = True
    Dim item As ITransit
    Dim range_pod_kolor As Range
    
    ' zanim wypelnimy ASNami
    ' jesli chcemy dorzucic ASLa
    If InitForm.CheckBoxCBALFromPOF.Value Then
    
        ofst = 2
        s.Cells(rrow + ofst, 9).Value = s.Cells(rrow + ofst, 9).Value + Int(vasl)
        ch.dodaj_raw_txt s.Cells(rrow + ofst, 9), "ASL: " & CStr(vasl)
    End If
    
    
    For Each item In kolekcja
        ' Debug.Print Application.WorksheetFunction.WeekNum(item.mDeliveryDate)
        
        'If one_time Then
        '    item.mDeliveryDate = CDate("2016-01-02")
        '    item.mPickupDate = CDate("2015-11-05")
        '    item.mMODE = "O"
        '    one_time = False
        'End If
        
        
        ' tutaj zalatwiam 9 kolumne past due
        If (item.mDeliveryDate = item.mPickupDate) Or item.mIsIP Then
            ofst = 4
            s.Cells(rrow + ofst, 9).Value = s.Cells(rrow + ofst, 9).Value + Int(item.mQty)
            s.Cells(rrow + ofst, 9).Interior.Color = RGB(200, 250, 200)
            ch.dodaj s.Cells(rrow + ofst, 9), item
        Else
        
            dd_cw = Int(Application.WorksheetFunction.IsoWeekNum(item.mDeliveryDate))
            dd_month = Month(item.mDeliveryDate)
            now_cw = Int(Application.WorksheetFunction.IsoWeekNum(Now))
            
            If Int(dd_cw) = 53 And Int(dd_month) = 1 Then
                dd_year = Int(Year(item.mDeliveryDate)) - 1
            Else
                dd_year = Int(Year(item.mDeliveryDate))
            End If
            now_year = Int(Year(Now))
        
            If ((dd_cw < now_cw) And (dd_year = now_year)) Or (dd_year < now_year) Then
            
            
                If CStr(ThisWorkbook.Sheets("register").Range("air")) Like "*" & CStr(item.mMODE) & "*" Then
                    ofst = 3
                    s.Cells(rrow + ofst, 9).Value = s.Cells(rrow + ofst, 9).Value + Int(item.mQty)
                    s.Cells(rrow + ofst, 9).Interior.Color = RGB(200, 250, 200)
                    s.Cells(rrow + ofst, 9).Font.Italic = True
                    ch.dodaj s.Cells(rrow + ofst, 9), item
                    
                ElseIf CStr(ThisWorkbook.Sheets("register").Range("sea")) Like "*" & CStr(item.mMODE) & "*" Then
                    ofst = 2
                    s.Cells(rrow + ofst, 9).Value = s.Cells(rrow + ofst, 9).Value + Int(item.mQty)
                    ch.dodaj s.Cells(rrow + ofst, 9), item
                    
                    sprawdz_edaminussdate item, s.Cells(rrow + ofst, 9), s.Cells(rrow + 4, 7)
                    sprawdz_delay_at_port item, s.Cells(rrow + ofst, 9)
                    
                Else
                    ofst = 2
                    s.Cells(rrow + ofst, 9).Value = s.Cells(rrow + ofst, 9).Value + Int(item.mQty)
                    Set range_pod_kolor = s.Cells(rrow + ofst, 9)
                    range_pod_kolor.Interior.Color = RGB(200, 250, 200)
                    ch.dodaj s.Cells(rrow + ofst, 9), item
                    
                    sprawdz_edaminussdate item, s.Cells(rrow + ofst, 9), s.Cells(rrow + 4, 7)
                    ' tylko sea
                    ' sprawdz_delay_at_port item, s.Cells(rrow + ofst, 9)
                    
                End If

            ' tutaj zalatwiam reszte kolumn dopasowanych do coverage i jego przyszlosci najblzszej
            Else
                For x = 10 To 10 + CoverCols
                    cw = Application.WorksheetFunction.IsoWeekNum(item.mDeliveryDate)
                    
                    ' dodatkowy warunek gdyz asny znikaja na cw 53
                    If Int(cw) = 53 Then
                        cw = 1
                    End If
                        If Int(cw) = Int(s.Cells(rrow, x)) Then
                            If CStr(ThisWorkbook.Sheets("register").Range("air")) Like "*" & CStr(item.mMODE) & "*" Then
                                ofst = 3
                                s.Cells(rrow + ofst, x).Value = s.Cells(rrow + ofst, x).Value + Int(item.mQty)
                                s.Cells(rrow + ofst, x).Font.Italic = True
                                ch.dodaj s.Cells(rrow + ofst, x), item
                                
                            ElseIf CStr(ThisWorkbook.Sheets("register").Range("sea")) Like "*" & CStr(item.mMODE) & "*" Then
                                ofst = 2
                                s.Cells(rrow + ofst, x).Value = s.Cells(rrow + ofst, x).Value + Int(item.mQty)
                                ch.dodaj s.Cells(rrow + ofst, x), item
                                
                                sprawdz_edaminussdate item, s.Cells(rrow + ofst, x), s.Cells(rrow + 4, 7)
                                sprawdz_delay_at_port item, s.Cells(rrow + ofst, x)
                                
                            Else
                                ofst = 4
                                s.Cells(rrow + ofst, x).Value = s.Cells(rrow + ofst, x).Value + Int(item.mQty)
                                ch.dodaj s.Cells(rrow + ofst, x), item
                                Set range_pod_kolor = s.Cells(rrow + ofst, x)
                                range_pod_kolor.Interior.Color = RGB(200, 250, 200)
                                
                                sprawdz_edaminussdate item, s.Cells(rrow + ofst, x), s.Cells(rrow + 4, 7)
                                ' tylko sea
                                ' sprawdz_delay_at_port item, s.Cells(rrow + ofst, x)
                                
                            End If
                        End If
                Next x
            End If
        End If
    Next item
    Set ch = Nothing
    ProgressIncrease
            
End Sub

Private Sub sprawdz_edaminussdate(item As ITransit, r As Range, ttimeindays As Range)
    If G_ASN01 Then
    
    If IsNumeric(item.mST) Then
        If Int(item.mST) < 2 Then
            If Math.Abs((item.mDeliveryDate - item.mPickupDate) - (ttimeindays)) > Int(ThisWorkbook.Sheets("register").Range("edaminussdate")) Then
            
                r.Font.Bold = True
                If (item.mDeliveryDate - item.mPickupDate) - (ttimeindays) > 0 Then
                    r.Font.Color = RGB(220, 34, 53)
                Else
                    r.Font.Color = RGB(46, 34, 213)
                End If
                
                
            Else
                r.Font.Bold = False
                r.Font.Color = RGB(0, 0, 0)
            End If
        End If
        
    End If
    
    End If
End Sub

Private Sub sprawdz_delay_at_port(item As ITransit, r As Range)
    If G_DELAY_FLAG And (G_DELAY_FLAG_BOUNDARY > 0) Then
    
        If CLng(item.mArrivalAtThePort) <> 0 Then
            If "23456" Like "*" & CStr(item.mST) & "*" Then
                If (item.mDeliveryDate - item.mArrivalAtThePort) <= G_DELAY_FLAG_BOUNDARY Then
                    
                    r.Font.Bold = True
                    r.Font.Color = RGB(240, 120, 20)
                    
                    
                Else
                    r.Font.Bold = False
                    r.Font.Color = RGB(0, 0, 0)
                End If
            End If
        End If
    
    End If
End Sub

Private Sub wez_wszystkie_total_rqm(ByRef Sess0, ByRef rrow, ByRef CoverCols)

    
    ' potrzebna jest petlaze zmienna pomocnicza wlasciwie dobrze by bylo napisac kod od nowa
    ' czyli za kazdym razem zaczynamy screen od samego poczatku inaczej gubia sie dane na screenie MGO
    
    na_ktorym_jestem_screenie = 0
    
    Do
    
        ' czesci bez danych
        If (Sess0.Screen.getString(22, 2, 5) Like "*R6017*") Or (Sess0.Screen.getString(22, 2, 5) Like "*R6117*") Then
            Exit Do
        End If
        
        ' tutaj oklipowane restrykcjami od góry i do³u jeœli chodzi o prace w tej petli
        
        
        If na_ktorym_jestem_screenie > 0 Then
        
            Sess0.Screen.sendKeys ("<Enter>")
            waitForMgo
        
            For x = 1 To na_ktorym_jestem_screenie
                
                Sess0.Screen.sendKeys ("<pf11>")
                waitForMgo
            Next x
        End If
        
    
        ' tutaj iteracja piekna do samego do³u ekranu
        While (Sess0.Screen.getString(22, 2, 5) = "R6693") Or ((na_ktorym_jestem_screenie) = 3 And (Sess0.Screen.getString(22, 2, 5) = "R6102"))
            Sess0.Screen.sendKeys ("<pf8>")
            waitForMgo
        Wend
        
        PltTotLine = 9
        For q = 9 To 20
            If Sess0.Screen.getString(q, 2, 9) = "PLT TOTAL" Then
                PltTotLine = q
                Exit For
            End If
        Next q
        
        
        For y = 0 To 4
            MDay = Val(Sess0.Screen.getString(8, 24 + 8 * y, 2))
            MMonth = MgoMonth(Sess0.Screen.getString(8, 27 + 8 * y, 2))
            If ((Month(Now) = 12) Or (Month(Now) = 11) Or (Month(Now) = 10)) And MMonth = 1 Then
                MYear = Year(Now) + 1
            Else
                MYear = Year(Now)
            End If
            mDate = DateSerial(MYear, MMonth, MDay)
            For q = 0 To CoverCols
                If mDate = this_workbook.ActiveSheet.Cells(rrow, 10 + q).Value Then
                    this_workbook.ActiveSheet.Cells(rrow + 1, 10 + q).Value = Val(Sess0.Screen.getString(PltTotLine, 22 + 8 * y, 8))
                    this_workbook.ActiveSheet.Cells(rrow + 2, 10 + q).Value = 0
                    Exit For
                End If
            Next q
        Next y
        
        If (Sess0.Screen.getString(22, 2, 5) Like "*R6120*") Or (Sess0.Screen.getString(22, 2, 5) Like "*R6123*") Or (Sess0.Screen.getString(22, 2, 5) Like "*R6017*") Then
            Exit Do
        End If
        
        na_ktorym_jestem_screenie = na_ktorym_jestem_screenie + 1
    Loop While na_ktorym_jestem_screenie < 4
End Sub

Sub MakeFullCoverage(Optional the_layout As layout_type, Optional asn01 As Boolean, Optional cpcritical As Boolean, Optional delay_flag_value As Boolean, Optional qhd_misc_flag As Boolean)


    G_ASN01 = asn01
    G_CP_CRITICAL = cpcritical
    G_DELAY_FLAG = delay_flag_value
    G_DELAY_FLAG_BOUNDARY = Int(InitForm.TextBoxDelayFlagBoundary.Value)
    G_QHD_MISC_BACKLOG = qhd_misc_flag

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = False

    Dim act_sh As Worksheet
    Dim plants(20) As String
    RowStart = 2
    OstatniWiersz = Sheets("Parts").Range("A10000").End(xlUp).row
    
    mgoInit
    StatusBox.Show
    Application.ScreenUpdating = False
    StatusBox.ProgressBar.Value = 0
    StatusBox.ProgressBar.Max = 4 * (OstatniWiersz - 1)
    Sheets.Add After:=Sheets(Sheets.Count)
    Set act_sh = ActiveSheet
    Set s = act_sh
    Set this_workbook = ThisWorkbook
    
    If the_layout = STD Then
    
        
        act_sh.Name = check_if_this_name_is_already_taken(CStr("simple_" & act_sh.Name))
    ElseIf the_layout = GREEN Then
        
        
        act_sh.Name = check_if_this_name_is_already_taken(CStr(act_sh.Name = "green_" & act_sh.Name))
    ElseIf the_layout = GREEN_AND_BLUE Then
    
        
        act_sh.Name = check_if_this_name_is_already_taken(CStr("green_and_blue_" & act_sh.Name))
    End If
    
    
    ActiveWindow.Zoom = 90
    ActiveWindow.DisplayGridlines = False
    
    Dim uchwyt_linii As Range
    Dim tmp_sess0 As Object
    
    For x = 2 To OstatniWiersz
    
        mgoInit
        Set tmp_sess0 = Sess0
    
        Set uchwyt_linii = this_workbook.Sheets("Parts").Cells(x, 2)
        If Not uchwyt_linii.EntireRow.Hidden Then
    
            If UCase(this_workbook.Sheets("Parts").Cells(x, 1).Value) = "GME" Then
                ' scan GME to get part numbers
                QQ = 0
                inPRT = this_workbook.Sheets("Parts").Cells(x, 2).Value
                StatusBox.Description.Caption = "Scan GME for " + CStr(inPRT)
                ProgressIncrease
                tmp_sess0.Screen.sendKeys "<Clear>"
                waitForMgo
                tmp_sess0.Screen.sendKeys "ms9pop00 <Enter>"
                waitForMgo
                tmp_sess0.Screen.putString "6", 2, 18
                tmp_sess0.Screen.putString "GME", 3, 5
                tmp_sess0.Screen.putString inPRT, 3, 22
                tmp_sess0.Screen.sendKeys "<Enter>"
                waitForMgo
                Do
                    plants(QQ) = tmp_sess0.Screen.getString(4, 13, 2)
                    If plants(QQ) <> "  " Then
                        QQ = QQ + 1
                    End If
                    tmp_sess0.Screen.sendKeys ("<Pf8>")
                    waitForMgo
                Loop Until tmp_sess0.Screen.getString(23, 2, 5) = "I4265"
                StatusBox.ProgressBar.Max = StatusBox.ProgressBar.Max + 3 * (QQ - 1)
                For Z = 0 To QQ - 1
                    this_workbook.ActiveSheet.Cells(RowStart, 1).Select
                    FormatCoverRecord this_workbook.ActiveSheet.Cells(RowStart, 1), the_layout
                    If Arkusz1.CheckBox1.Value Then Autoformatowanie
                    
                    MakeCoverage RowStart, plants(Z), inPRT, this_workbook.Sheets("Parts").Cells(x, 3).Value, _
                        this_workbook.Sheets("Parts").Cells(x, 4).Value, this_workbook.Sheets("Parts").Cells(x, 3).Interior.Color
                        
                    kreska_ttime_u New TheLayout, this_workbook.ActiveSheet.Cells(RowStart, 1)
                        
                    RowStart = RowStart + 7
                    
                    
                Next Z
            
            Else
                MakeCoverage RowStart, this_workbook.Sheets("Parts").Cells(x, 1).Value, _
                    this_workbook.Sheets("Parts").Cells(x, 2).Value, this_workbook.Sheets("Parts").Cells(x, 3).Value, _
                    this_workbook.Sheets("Parts").Cells(x, 4).Value, this_workbook.Sheets("Parts").Cells(x, 3).Interior.Color
                    
                
                
                this_workbook.ActiveSheet.Cells(RowStart, 1).Select
                FormatCoverRecord this_workbook.ActiveSheet.Cells(RowStart, 1), the_layout
                
                kreska_ttime_u New TheLayout, this_workbook.ActiveSheet.Cells(RowStart, 1)
                
                If Arkusz1.CheckBox1.Value Then Autoformatowanie
                RowStart = RowStart + 7
            End If
        End If
    Next x
    Columns("A:AC").EntireColumn.AutoFit
    'Columns("C:D").Select
    'Selection.EntireColumn.Hidden = True
    'Columns("G:G").Select
    'Selection.EntireColumn.Hidden = True
    StatusBox.Hide
    Range("A1").Select
    Selection = the_layout
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.Calculate
    przelicz_arkusz Int(Range("a1")), True
End Sub

Sub DodajKomentarz(rrow, RCol, Komentarz, Szerokosc, Optional CHeader As String)
    If this_workbook.ActiveSheet.Cells(rrow, RCol).Comment Is Nothing Then
        this_workbook.ActiveSheet.Cells(rrow, RCol).AddComment
        If CHeader > "" Then
            this_workbook.ActiveSheet.Cells(rrow, RCol).Comment.Text Text:=CHeader + Chr(10)
        End If
    End If
    
    this_workbook.ActiveSheet.Cells(rrow, RCol).Comment.Text Text:=this_workbook.ActiveSheet.Cells(rrow, RCol).Comment.Text + Komentarz + Chr(10)
    this_workbook.ActiveSheet.Cells(rrow, RCol).Comment.Shape.Width = Szerokosc
    Tr = this_workbook.ActiveSheet.Cells(rrow, RCol).Comment.Text
    this_workbook.ActiveSheet.Cells(rrow, RCol).Comment.Shape.Height = CountLines(Tr) * 12
End Sub

Function CountLines(txt)
    Lr = Split(txt, Chr(10))
    CountLines = UBound(Lr)
End Function

Private Function take_data_from_alc(alc_sheet As Worksheet, plt As String, part As String, ile_offsetu As Integer) As Range

    ' ile offsetu to parametr okreslajacy z jakiej kolumny sciagamy dane
    ' ile_offsetu = 1 == Frei
    ' ile_offsetu = 2 == Gesperrt

    Set take_data_from_alc = Nothing
    If Len(part) = 8 Then
        Set take_data_from_alc = alc_sheet.Columns("B:B").Find(part)
    End If


poczatek:
    If Len(plt) = 2 Then
        If plt <> CStr(take_data_from_alc.offset(0, -1)) Then
            Set take_data_from_alc = alc_sheet.Columns("B:B").Find(part, take_data_from_alc)
            GoTo poczatek
        End If
    End If
    
    Set take_data_from_alc = take_data_from_alc.offset(0, Int(ile_offsetu))
End Function


Private Function hit_F8_until_you_see_2000_on_sched_point()
    While CStr(Sess0.Screen.getString(19, 20, 4)) <> "2000"
        Sess0.Screen.sendKeys ("<pf8>")
        waitForMgo
    Wend
End Function

Private Sub put_zero_on_BLD_SEQ()
    
    Sess0.Screen.putString "    ", 6, 46
    Sess0.Screen.putString "0", 6, 46
    Sess0.Screen.sendKeys ("<Enter>")
    waitForMgo
End Sub


Private Function check_if_this_name_is_already_taken(s As String, Optional inc As Integer)
    
    
    If inc = 0 Then
        ' inc = 0
        check_if_this_name_is_already_taken = s
    Else
        check_if_this_name_is_already_taken = "sh_" & CStr(inc)
    End If
    
    Dim n As String
    For Each wsh In ThisWorkbook.Sheets
        If CStr(wsh.Name) = CStr(check_if_this_name_is_already_taken) Then
            
            ' jesli wpadnie cos w ten warunek to znaczy ze juz taka nazwa jest i trzeba nieco zmienic zawartosc nazwy
            check_if_this_name_is_already_taken = check_if_this_name_is_already_taken(CStr(check_if_this_name_is_already_taken), inc + 1)
        End If
    Next wsh
End Function


