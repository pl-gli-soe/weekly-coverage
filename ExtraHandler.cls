VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtraHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private s As ExtraSession
Private ss As ExtraSessions
Private sys As ExtraSystem
Private scr As ExtraScreen

Private plt As String
Private pn As String


Private Sub Class_Initialize()
    'Global Sessions As Object
    'Global System As Object
    'Global Sess0 As Object
    
    'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    'Sub mgoInit()
    mgoInit
    'Set System = CreateObject("EXTRA.System")   ' Gets the system object
    'Set Sessions = System.Sessions
    'Set Sess0 = System.ActiveSession

    Set sys = System
    Set ss = Sessions
    
    Set s = Nothing
    On Error Resume Next
    Set s = Sess0
    Set scr = s.Screen
    
    If s Is Nothing Then
        MsgBox "extra isn't active!"
    End If
    
    s.Activate
    
    If ThisWorkbook.FullName = ActiveWorkbook.FullName Then
        If check_if_active_sheet_is_rep() Then
            pn = take_selected_("PN")
            plt = take_selected_("PLT")
            
            
            If pn <> "" And plt <> "" Then
                ' scr.ClearScreen
                scr.sendKeys ("<Clear>")
                waitForMgo
                scr.sendKeys ("ms9pop00 <Enter>")
                waitForMgo
                
                scr.sendKeys "6"
                scr.sendKeys ("<Tab>")
                scr.sendKeys ("<Tab>")
                scr.sendKeys CStr(plt)
                scr.sendKeys CStr(pn)
                
                scr.sendKeys ("<Enter>")
                waitForMgo
            End If
        End If
    End If

End Sub

Private Sub Class_Terminate()
    Set sys = Nothing
    Set ss = Nothing
    Set s = Nothing
    Set scr = Nothing
End Sub


Public Sub run_macro(arg As String)
     
     
     
    ' teraz tutaj przyda sie implementacja ktora ogarnie
    ' w jaki sposob jestem w stanie uruchomic makra z poziomu extry
    Dim strProgramName As String, strProgramNamex86 As String
    Dim strArgument As String
    
    
    strProgramName = "C:\Program Files\Attachmate\EXTRA!\Ebrun.exe"
    strProgramNamex86 = "C:\Program Files (x86)\Attachmate\EXTRA!\Ebrun.exe"
    strArgument = "C:\mgo_scr\" & CStr(arg) & ".ebm"
    
    On Error Resume Next
    Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
    
    On Error Resume Next
    Call Shell("""" & strProgramNamex86 & """ """ & strArgument & """", vbNormalFocus)
End Sub



Private Function take_selected_(arg As String) As String
    take_selected_ = ""
    
    
    Dim rng As Range
    Set rng = Selection.item(1)
    
    ktory_pn = CLng(Application.WorksheetFunction.RoundUp(rng.row / 8, 0))
    
    
    
    If CStr(arg) = "PN" Then
    
        take_selected_ = CStr(Cells(3 + (7 * (ktory_pn - 1)), 2))
    ElseIf CStr(arg) = "PLT" Then
    
        take_selected_ = CStr(Cells(3 + (7 * (ktory_pn - 1)), 1))
    End If
End Function

Private Function check_if_active_sheet_is_rep() As Boolean
    check_if_active_sheet_is_rep = False
    
    If Range("G2") = "First runout" Then
        check_if_active_sheet_is_rep = True
    End If
End Function

