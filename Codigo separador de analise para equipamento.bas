Attribute VB_Name = "Módulo12"
Sub Varian()

Dim range1, cell As Range
Dim aba_idendificacao, aba_agilent, aba_varian As String
Dim contador, total, z As Integer

Application.ScreenUpdating = False
aba_idendificacao = "IDENTIFICAÇÃO DE AMOSTRAS"
aba_agilent = "ANALISE_MERC_ AGILENT"
aba_varian = "ANALISE_MERC_VARIAN"

Sheets(aba_varian).Activate
Range("B10:B26").Select
Selection.ClearContents
Set range1 = Range("B10", Range("B10").End(xlDown))
Sheets(aba_idendificacao).Activate
Range("A1").Select



For Each cell In range1
        
    Sheets(aba_idendificacao).Activate
    Cells(27, 7).FormulaR1C1 = "=COUNTIF(R[-16]C[4]:R[0]C[4],""Varian"")"
    z = CInt(Range("G27").Value)
    

    If contador < z Then
    

    Pesquisa = Cells.Find(What:="Varian", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
        contador = contador + 1
        cell.Offset(0, 0).Value = Selection.Offset(0, -10).Value
        
    Else
        Sheets(aba_idendificacao).Activate
        Cells(27, 7) = " "
        Sheets(aba_varian).Activate
        End
        
    End If

    
Next


Sheets(aba_varian).Activate


   
End Sub

Sub Agilent()

Dim range1, cell As Range
Dim aba_idendificacao, aba_agilent, aba_varian As String
Dim contador, total, z As Integer

Application.ScreenUpdating = False
aba_idendificacao = "IDENTIFICAÇÃO DE AMOSTRAS"
aba_agilent = "ANALISE_MERC_ AGILENT"
aba_varian = "ANALISE_MERC_VARIAN"

Sheets(aba_agilent).Activate
Range("B10:B26").Select
Selection.ClearContents
Set range1 = Range("B10", Range("B10").End(xlDown))
Sheets(aba_idendificacao).Activate
Range("A1").Select



For Each cell In range1
        
    Sheets(aba_idendificacao).Activate
    Cells(27, 7).FormulaR1C1 = "=COUNTIF(R[-16]C[4]:R[0]C[4],""Agilent"")"
    z = CInt(Range("G27").Value)
    

    If contador < z Then
    

    Pesquisa = Cells.Find(What:="Agilent", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
        contador = contador + 1
        cell.Offset(0, 0).Value = Selection.Offset(0, -10).Value
        
    Else
        Sheets(aba_idendificacao).Activate
        Cells(27, 7) = " "
        Sheets(aba_agilent).Activate
       
        End
        
    End If

    
Next


Sheets(aba_agilent).Activate


   
End Sub






