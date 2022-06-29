VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCadastro 
   Caption         =   "Cadastro de Equipamentos"
   ClientHeight    =   9300.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870.001
   OleObjectBlob   =   "Codigos de formulario de preenchimento de banco de dados.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 
Dim ctrl As Control


  
    If TextBox1.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label1.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If TextBox2.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label2.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If TextBox3.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label3.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If TextBox4.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label4.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If TextBox7.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label16.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If ComboBox1.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label5.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    If ComboBox3.Text = "" Then
        MsgBox "Preenchimento Obrigatório do campo " & Label21.Caption, vbExclamation, "AVISO"
        TextBox1.SetFocus
        Exit Sub
    End If
    
    lsInserirTextBox frmCadastro, "Dados", 1
    
    lsLimparTextBox frmCadastro
    
    TextBox1.SetFocus
    
    MsgBox "Produto Cadastrado com sucesso", vbOKOnly, "Aviso"
    

    
End Sub

Private Sub lsInserir(ByRef lTextBox As Variant, ByVal lSheet As String, ByVal lColunaCodigo As Long, ByVal lUltimaLinha As Long)
    If (TypeOf lTextBox Is MSForms.TextBox) Or (TypeOf lTextBox Is MSForms.ComboBox) Then
        Sheets(lSheet).Range(lTextBox.Tag & lUltimaLinha).Value = lTextBox.Text
    Else
        If TypeOf lTextBox Is MSForms.OptionButton Then
            If lTextBox.Value = True Then
                Sheets(lSheet).Range(lTextBox.Tag & lUltimaLinha).Value = lTextBox.Caption
            End If
        End If
    End If
End Sub

Public Function lsInserirTextBox(formulario As UserForm, ByVal lSheet As String, ByVal lColunaCodigo As Long)
    Dim controle            As Control
    Dim lUltimaLinhaAtiva   As Long
    
    lUltimaLinhaAtiva = Worksheets(lSheet).Cells(Worksheets(lSheet).Rows.Count, lColunaCodigo).End(xlUp).Row + 1
    
    For Each controle In formulario.Controls
        lsInserir controle, lSheet, lColunaCodigo, lUltimaLinhaAtiva
    Next
End Function

Public Function lsLimparTextBox(formulario As UserForm)
    Dim controle            As Control
    
    For Each controle In formulario.Controls
        If TypeOf controle Is MSForms.TextBox Then
            controle.Text = ""
        End If
        If TypeOf controle Is MSForms.ComboBox Then
            controle.Text = ""
        End If
    Next
End Function

Private Sub CommandButton2_Click()
    lsLimparTextBox frmCadastro
    
    TextBox1.SetFocus
End Sub

Private Sub CommandButton3_Click()

    frmCadastro.Hide
    
End Sub


Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

letras = Len(Range("C2")) + 5

TextBox1.MaxLength = letras + 3

Select Case KeyAscii

Case 8
Case 13: SendKeys "{TAB}"

Case 48 To 57



If TextBox1.SelStart = 0 Then TextBox1.SelText = Range("C2").Value & " "
If TextBox1.SelStart = letras Then TextBox1.SelText = "."

Case Else: KeyAscii = 0
End Select

End Sub

Private Sub TextBox13_Change()

End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

TextBox6.MaxLength = 10

Select Case KeyAscii

Case 8
Case 13: SendKeys "{TAB}"

Case 48 To 57

If TextBox6.SelStart = 2 Then TextBox6.SelText = "/"
If TextBox6.SelStart = 5 Then TextBox6.SelText = "/"
Case Else: KeyAscii = 0
End Select

End Sub
Private Sub TextBox6_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox6) And TextBox6 <> "" Then

        MsgBox "data inválida"
        
        TextBox6 = ""

        Cancel = True

    End If

End Sub



Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    TextBox8.MaxLength = 10

    Select Case KeyAscii

        Case 8
        Case 13: SendKeys "{TAB}"

        Case 48 To 57

        If TextBox8.SelStart = 2 Then TextBox8.SelText = "/"
        If TextBox8.SelStart = 5 Then TextBox8.SelText = "/"
        Case Else: KeyAscii = 0
    End Select

End Sub
Private Sub TextBox8_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox8) And TextBox8 <> "" Then

        MsgBox "data inválida"

        TextBox8 = ""

        Cancel = True

    End If

End Sub


Private Sub TextBox9_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    TextBox9.MaxLength = 10

    Select Case KeyAscii

        Case 8
        Case 13: SendKeys "{TAB}"

        Case 48 To 57

        If TextBox9.SelStart = 2 Then TextBox9.SelText = "/"
        If TextBox9.SelStart = 5 Then TextBox9.SelText = "/"
        Case Else: KeyAscii = 0
    End Select

End Sub
Private Sub TextBox9_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox9) And TextBox9 <> "" Then

        MsgBox "data inválida"

        TextBox9 = ""

        Cancel = True

    End If

End Sub

