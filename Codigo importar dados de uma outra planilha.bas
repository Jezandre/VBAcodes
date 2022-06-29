Attribute VB_Name = "Módulo1"
Sub Importar_dados()

On Error GoTo Erro

Application.ScreenUpdating = False

Dim Guia As Object
Dim Planilha As Workbook
Dim EnderecoPlan As String
Dim Coluna As Double, Linha As Double, ColDestino As Double
Dim ColInicial As Double, ColFinal As Double, LinOrigem As Double

mes = ActiveSheet.Cells(7, 2).Value

EnderecoPlan = Application.GetOpenFilename(filefilter:="file, *.xls*")


If EnderecoPlan <> Empty And EnderecoPlan <> "Falso" Then
    Set Planilha = Application.Workbooks.Open(EnderecoPlan)
    Else
    Application.ScreenUpdating = True
    Exit Sub
End If




Set Guia = Planilha.Worksheets(mes)

Windows(Planilha.Name).Visible = False

Coluna = 1
Linha = 10


Inicio:
Do
    Linha = Linha + 1
    If Guia.Cells(Linha, Coluna).Value <> Empty Then
        LinOrigem = Linha
        ColInicial = Coluna
        
        Do
            Coluna = Coluna + 1
            Loop Until Guia.Cells(Linha, Coluna).Value = Empty
            ColFinal = Coluna - 1
        Exit Do
    End If
    
    If Coluna = 100 Then
        MsgBox "Não encontrado cabeçalho!", vbExclamation, "IMPORTAR"
        Exit Sub
    End If
    

Loop Until Linha = 100

If LinOrigem = Empty Then
    Coluna = Coluna + 1
    Linha = 1
    GoTo Inicio:
    
End If

Coluna = ColInicial
ColDestino = 1
Linha = WorksheetFunction.CountA(ActiveSheet.Range("A:A")) + 8


With ActiveSheet
    Do
    
        LinOrigem = LinOrigem + 1
        If Guia.Cells(LinOrigem, 3).Value = "" Then
              GoTo Fim:
            End If
        For Coluna = ColInicial To ColFinal
            .Cells(Linha, ColDestino).Value = Guia.Cells(LinOrigem, Coluna).Value
            ColDestino = ColDestino + 1

              
        Next Coluna
        
        ColDestino = 1
        Linha = Linha + 1
                
    Loop Until Guia.Cells(LinOrigem, ColInicial).Value = Empty
        
End With

Fim:
Windows(Planilha.Name).Visible = True
Application.DisplayAlerts = False
Windows(Planilha.Name).Close
Application.DisplayAlerts = True


Set Planilha = Nothing
Set Guia = Nothing

Application.ScreenUpdating = True
MsgBox "Os dados de saída de " & EnderecoPlan & " foram importados com sucesso!"

Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "IMPORTAR"

End Sub

