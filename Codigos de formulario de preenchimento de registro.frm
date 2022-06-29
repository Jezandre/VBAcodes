VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calibracao 
   Caption         =   "Registo de Calibração"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "Codigos de formulario de preenchimento de registro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calibracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_click()

    flinha = Cells.Find(ComboBox1).Select
    TextBox1.Value = ActiveCell.Offset(0, 16).Value
    TextBox2.Value = ActiveCell.Offset(0, 20).Value
    TextBox3.Value = ActiveCell.Offset(0, 24).Value
    TextBox4.Value = ActiveCell.Offset(0, 28).Value
    TextBox5.Value = ActiveCell.Offset(0, 17).Value
    TextBox6.Value = ActiveCell.Offset(0, 21).Value
    TextBox7.Value = ActiveCell.Offset(0, 25).Value
    TextBox8.Value = ActiveCell.Offset(0, 29).Value
    TextBox9.Value = ActiveCell.Offset(0, 15).Value
    TextBox10.Value = ActiveCell.Offset(0, 19).Value
    TextBox11.Value = ActiveCell.Offset(0, 23).Value
    TextBox12.Value = ActiveCell.Offset(0, 27).Value

End Sub




Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox1) And TextBox1 <> "" Then

        MsgBox "data inválida"

        TextBox1 = ""

        Cancel = True

    End If

End Sub

Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox2) And TextBox2 <> "" Then

        MsgBox "data inválida"

        TextBox2 = ""

        Cancel = True

    End If

End Sub

Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox3) And TextBox3 <> "" Then

        MsgBox "data inválida"

        TextBox3 = ""

        Cancel = True

    End If

End Sub

Private Sub TextBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsDate(TextBox4) And TextBox4 <> "" Then

        MsgBox "data inválida"

        TextBox4 = ""

        Cancel = True

    End If

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    TextBox1.MaxLength = 10

    Select Case KeyAscii

        Case 8
        Case 13: SendKeys "{TAB}"

        Case 48 To 57

        If TextBox1.SelStart = 2 Then TextBox1.SelText = "/"
        If TextBox1.SelStart = 5 Then TextBox1.SelText = "/"
        Case Else: KeyAscii = 0
    End Select

End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    TextBox2.MaxLength = 10

    Select Case KeyAscii

        Case 8
        Case 13: SendKeys "{TAB}"

        Case 48 To 57

        If TextBox2.SelStart = 2 Then TextBox2.SelText = "/"
        If TextBox2.SelStart = 5 Then TextBox2.SelText = "/"
        Case Else: KeyAscii = 0
    End Select

End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    TextBox3.MaxLength = 10

    Select Case KeyAscii

        Case 8
        Case 13: SendKeys "{TAB}"

        Case 48 To 57

        If TextBox3.SelStart = 2 Then TextBox3.SelText = "/"
        If TextBox3.SelStart = 5 Then TextBox3.SelText = "/"
        Case Else: KeyAscii = 0
    End Select

End Sub
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    TextBox4.MaxLength = 10

    Select Case KeyAscii

        Case 8
        Case 13: SendKeys "{TAB}"

        Case 48 To 57

        If TextBox4.SelStart = 2 Then TextBox4.SelText = "/"
        If TextBox4.SelStart = 5 Then TextBox4.SelText = "/"
        Case Else: KeyAscii = 0
    End Select

End Sub

Private Sub CommandButton1_Click()

If ComboBox1.MatchFound = False Then
    GoTo LastLine2
End If

If ComboBox1.Text = "" Then
    GoTo LastLine2
End If


'Textbox Grandeza de calibração'

    If TextBox9 <> "" Then
        On Error GoTo LastLine2
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 15).Value = TextBox9

    Else
        GoTo LastLine
    End If
    
    If TextBox10 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 19).Value = TextBox10

    Else
        GoTo LastLine
    End If
    
    If TextBox11 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 23).Value = TextBox11

    Else
        GoTo LastLine
    End If
    
    If TextBox12 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 27).Value = TextBox12

    Else
        GoTo LastLine
    End If
    

'Text box datas'

    If TextBox1 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 16).Value = TextBox1
        ActiveCell.Offset(0, 16).Select
    Else
        GoTo LastLine
    End If
    
    If TextBox2 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 20).Value = TextBox2
    Else
        GoTo LastLine
    End If
    
    If TextBox3 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 24).Value = TextBox3
    Else
        GoTo LastLine
    End If
    
    If TextBox4 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 28).Value = TextBox4
        ActiveCell.Offset(0, 28).Select
    Else
        GoTo LastLine
    End If
    
    
'Textbox prazo'

    If TextBox5 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 17).Value = TextBox5

    Else
        GoTo LastLine
    End If
    
    If TextBox6 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 21).Value = TextBox6

    Else
        GoTo LastLine
    End If
    
    If TextBox7 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 25).Value = TextBox7

    Else
        GoTo LastLine
    End If
    
    If TextBox8 <> "" Then
        Cells.Find(ComboBox1).Select
        ActiveCell.Offset(0, 29).Value = TextBox8

    Else
        GoTo LastLine
    End If
    

LastLine:
    MsgBox "Data de Calibração registrada com sucesso", vbApplicationModal
    GoTo LastLine3

LastLine2:
        MsgBox "Insira uma ID Válida", vbExclamation, "Atenção"

LastLine3:

End Sub

Private Sub CommandButton2_Click()

    Calibracao.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Me.ComboBox1.RowSource = "Dados!A4:A3000"
    
End Sub
