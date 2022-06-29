Attribute VB_Name = "Módulo12"
Sub Calibracao_01()


Dim Linha As Integer
Dim Coluna As Integer
Dim etiqueta As Integer

Hoje = Date
Linha = 2
Coluna = 1
etiqueta = 0

Application.ScreenUpdating = False

Worksheets("Calibrações").Activate
Worksheets("Etiquetas").Range("A2:AZ100").ClearContents
Worksheets("Etiquetas").Range("A2:AZ100").ClearFormats
For Each celula In Worksheets("Calibrações").Range("I6:I7000")
      If celula <> "-" Then
          If celula - Hoje >= 0 Then
                
                Worksheets("Etiquetas").Cells(Linha, Coluna).Value = "Identificação: " & celula.Offset(0, -6).Value
                Worksheets("Etiquetas").Cells(Linha, Coluna).Rows.RowHeight = 17
                Worksheets("Etiquetas").Cells(Linha, Coluna).Columns.ColumnWidth = 28
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Characters(1, 14).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Interior
                    .Pattern = xlSolid
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0.499984740745262
                    .PatternTintAndShade = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Value = "Nome: " & celula.Offset(0, -5).Value
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).VerticalAlignment = xlCenter
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Rows.RowHeight = 28
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).WrapText = True
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Characters(1, 5).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Value = "Última Calibração: " & celula.Offset(0, -2).Value
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Characters(1, 17).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Value = "Próxima Calibração: " & celula.Value
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Characters(1, 18).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
    
                Linha = Linha + 5
                etiqueta = etiqueta + 1
    
        
          End If
      End If

If etiqueta = 8 Then
    etiqueta = 0
    Linha = 2
    Coluna = Coluna + 2
    Worksheets("Etiquetas").Cells(Linha, Coluna - 1).Columns.ColumnWidth = 1
   
End If


Worksheets("Etiquetas").Activate
    
Next
End Sub

Sub Calibracao_02()


Dim Linha As Integer
Dim Coluna As Integer
Dim etiqueta As Integer

Hoje = Date
Linha = 2
Coluna = 1
etiqueta = 0

Application.ScreenUpdating = False

Worksheets("Calibrações").Activate
Worksheets("Etiquetas").Range("A2:AZ100").ClearContents
Worksheets("Etiquetas").Range("A2:AZ100").ClearFormats
For Each celula In Worksheets("Calibrações").Range("M6:M7000")
      If celula <> "-" Then
          If celula - Hoje >= 0 Then
                
                Worksheets("Etiquetas").Cells(Linha, Coluna).Value = "Identificação: " & celula.Offset(0, -10).Value
                Worksheets("Etiquetas").Cells(Linha, Coluna).Rows.RowHeight = 17
                Worksheets("Etiquetas").Cells(Linha, Coluna).Columns.ColumnWidth = 28
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Characters(1, 14).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Interior
                    .Pattern = xlSolid
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0.499984740745262
                    .PatternTintAndShade = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Value = "Nome: " & celula.Offset(0, -9).Value
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).VerticalAlignment = xlCenter
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Rows.RowHeight = 28
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).WrapText = True
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Characters(1, 5).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Value = "Ultima Calibração: " & celula.Offset(0, -6).Value
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Characters(1, 17).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Value = "Próxima Calibração: " & celula.Value
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Characters(1, 18).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
    
                Linha = Linha + 5
                etiqueta = etiqueta + 1
    
        
          End If
      End If

If etiqueta = 8 Then
    etiqueta = 0
    Linha = 2
    Coluna = Coluna + 2
    Worksheets("Etiquetas").Cells(Linha, Coluna - 1).Columns.ColumnWidth = 1
   
End If

Worksheets("Etiquetas").Activate
    
Next
End Sub

Sub Calibracao_03()


Dim Linha As Integer
Dim Coluna As Integer
Dim etiqueta As Integer

Hoje = Date
Linha = 2
Coluna = 1
etiqueta = 0

Application.ScreenUpdating = False

Worksheets("Calibrações").Activate
Worksheets("Etiquetas").Range("A2:AZ100").ClearContents
Worksheets("Etiquetas").Range("A2:AZ100").ClearFormats
For Each celula In Worksheets("Calibrações").Range("Q6:Q7000")
      If celula <> "-" Then
          If celula - Hoje >= 0 Then
                
                Worksheets("Etiquetas").Cells(Linha, Coluna).Value = "Identificação: " & celula.Offset(0, -14).Value
                Worksheets("Etiquetas").Cells(Linha, Coluna).Rows.RowHeight = 17
                Worksheets("Etiquetas").Cells(Linha, Coluna).Columns.ColumnWidth = 28
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Characters(1, 14).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Interior
                    .Pattern = xlSolid
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0.499984740745262
                    .PatternTintAndShade = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Value = "Nome: " & celula.Offset(0, -13).Value
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).VerticalAlignment = xlCenter
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Rows.RowHeight = 28
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).WrapText = True
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Characters(1, 5).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Value = "Ultima Calibração: " & celula.Offset(0, -10).Value
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Characters(1, 17).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Value = "Próxima Calibração: " & celula.Value
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Characters(1, 18).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
    
                Linha = Linha + 5
                etiqueta = etiqueta + 1
    
        
          End If
      End If

If etiqueta = 8 Then
    etiqueta = 0
    Linha = 2
    Coluna = Coluna + 2
    Worksheets("Etiquetas").Cells(Linha, Coluna - 1).Columns.ColumnWidth = 1
   
End If

Worksheets("Etiquetas").Activate
    
Next
End Sub

Sub Calibracao_04()


Dim Linha As Integer
Dim Coluna As Integer
Dim etiqueta As Integer

Hoje = Date
Linha = 2
Coluna = 1
etiqueta = 0

Application.ScreenUpdating = False

Worksheets("Calibrações").Activate
Worksheets("Etiquetas").Range("A2:AZ100").ClearContents
Worksheets("Etiquetas").Range("A2:AZ100").ClearFormats
For Each celula In Worksheets("Calibrações").Range("U6:U7000")
      If celula <> "-" Then
          If celula - Hoje >= 0 Then
                
                Worksheets("Etiquetas").Cells(Linha, Coluna).Value = "Identificação: " & celula.Offset(0, -18).Value
                Worksheets("Etiquetas").Cells(Linha, Coluna).Rows.RowHeight = 17
                Worksheets("Etiquetas").Cells(Linha, Coluna).Columns.ColumnWidth = 28
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Characters(1, 14).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                With Worksheets("Etiquetas").Cells(Linha, Coluna).Interior
                    .Pattern = xlSolid
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0.499984740745262
                    .PatternTintAndShade = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Value = "Nome: " & celula.Offset(0, -17).Value
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).VerticalAlignment = xlCenter
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Rows.RowHeight = 28
                Worksheets("Etiquetas").Cells(Linha + 1, Coluna).WrapText = True
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Characters(1, 5).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 1, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Value = "Ultima Calibração: " & celula.Offset(0, -14).Value
                Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Characters(1, 17).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 2, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
                
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Value = "Próxima Calibração: " & celula.Value
                Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Rows.RowHeight = 17
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Characters(1, 18).Font
                    .Name = "Calibri"
                    .FontStyle = "Negrito"
                End With
                With Worksheets("Etiquetas").Cells(Linha + 3, Coluna).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = 0
                End With
                
    
                Linha = Linha + 5
                etiqueta = etiqueta + 1
    
        
          End If
      End If

If etiqueta = 8 Then
    etiqueta = 0
    Linha = 2
    Coluna = Coluna + 2
    Worksheets("Etiquetas").Cells(Linha, Coluna - 1).Columns.ColumnWidth = 1
   
End If

Worksheets("Etiquetas").Activate
    
Next
End Sub

Sub Imprimir_etiquetas()

     Dim Lr As Long
     With ActiveSheet
   Lr = .Cells(Rows.Count, 1).End(xlUp).Row     'define a ultima linha nao vazia
   Lr = IIf(Lr < 5, 5, Lr)
      .PageSetup.PrintArea = ("A2:Z" & Lr)
     End With
    Application.Dialogs(xlDialogPrintPreview).Show
End Sub
