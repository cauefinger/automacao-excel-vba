Option Explicit

Const ABA_ORIGEM = "Pessoas"
Const ABA_DESTINO = "Prazos"
Const PRAZO_DIAS = 7
Const DIAS_ALERTA = 2

Const COL_NOME = 0
Const COL_CARGO = 1
Const COL_FATO = 2
Const COL_DATA = 3

Const COL_PRAZO = 4
Const COL_SITUACAO = 5

Sub OlaMacro()
    Dim doc As Object
    Dim linha As Long

    doc = ThisComponent
    linha = doc.CurrentSelection.getRangeAddress().StartRow

    MsgBox "Voce selecionou a linha " & (linha + 1), 0, "Ola, macro!"
End Sub

Sub EnviarParaPrazo()
    Dim doc As Object, wsA As Object, wsB As Object
    Dim celData As Object
    Dim linhaA As Long, linhaB As Long
    Dim dataFato As Double
    Dim fmt As Long

    doc = ThisComponent
    wsA = doc.Sheets.getByName(ABA_ORIGEM)
    wsB = doc.Sheets.getByName(ABA_DESTINO)

    If doc.CurrentController.ActiveSheet.Name <> ABA_ORIGEM Then
        MsgBox "Va para a aba '" & ABA_ORIGEM & "' e selecione a linha.", 48
        Exit Sub
    End If

    linhaA = doc.CurrentSelection.getRangeAddress().StartRow
    If linhaA < 1 Then
        MsgBox "Selecione uma linha com dados (nao o cabecalho).", 48
        Exit Sub
    End If

    celData = wsA.getCellByPosition(COL_DATA, linhaA)
    If celData.Type <> com.sun.star.table.CellContentType.VALUE Then
        MsgBox "A data esta vazia ou invalida na linha " & (linhaA + 1) & ".", 48
        Exit Sub
    End If
    dataFato = celData.Value

    If JaEnviado(wsB, wsA.getCellByPosition(COL_NOME, linhaA).String, dataFato) Then
        MsgBox "Esta ocorrencia ja foi enviada para '" & ABA_DESTINO & "'.", 64
        Exit Sub
    End If

    linhaB = PrimeiraLinhaVazia(wsB)

    wsB.getCellByPosition(0, linhaB).String = wsA.getCellByPosition(COL_NOME, linhaA).String
    wsB.getCellByPosition(1, linhaB).String = wsA.getCellByPosition(COL_CARGO, linhaA).String
    wsB.getCellByPosition(2, linhaB).String = wsA.getCellByPosition(COL_FATO, linhaA).String
    wsB.getCellByPosition(3, linhaB).Value = dataFato
    wsB.getCellByPosition(4, linhaB).Value = dataFato + PRAZO_DIAS

    fmt = FormatoData(doc)
    wsB.getCellByPosition(3, linhaB).NumberFormat = fmt
    wsB.getCellByPosition(4, linhaB).NumberFormat = fmt

    AtualizarSituacao
    doc.CurrentController.setActiveSheet(wsB)
    doc.CurrentController.select(wsB.getCellByPosition(0, linhaB))
End Sub

Function JaEnviado(ws As Object, nome As String, dt As Double) As Boolean
    Dim i As Long
    i = 1
    Do While ws.getCellByPosition(0, i).String <> ""
        If ws.getCellByPosition(0, i).String = nome _
           And ws.getCellByPosition(3, i).Value = dt Then
            JaEnviado = True
            Exit Function
        End If
        i = i + 1
    Loop
    JaEnviado = False
End Function

Function PrimeiraLinhaVazia(ws As Object) As Long
    Dim i As Long
    i = 1
    Do While ws.getCellByPosition(0, i).String <> ""
        i = i + 1
    Loop
    PrimeiraLinhaVazia = i
End Function

Function FormatoData(doc As Object) As Long
    Dim loc As New com.sun.star.lang.Locale
    Dim chave As Long
    loc.Language = "pt"
    loc.Country = "BR"
    chave = doc.NumberFormats.queryKey("DD/MM/YYYY", loc, False)
    If chave = -1 Then
        chave = doc.NumberFormats.addNew("DD/MM/YYYY", loc)
    End If
    FormatoData = chave
End Function

Sub AtualizarSituacao()
    Dim ws As Object, c As Object
    Dim i As Long, ultima As Long
    Dim hoje As Double, faltam As Double

    ws = ThisComponent.Sheets.getByName(ABA_DESTINO)
    hoje = CLng(Date)
    ultima = PrimeiraLinhaVazia(ws) - 1

    For i = 1 To ultima
        faltam = ws.getCellByPosition(COL_PRAZO, i).Value - hoje
        c = ws.getCellByPosition(COL_SITUACAO, i)

        If faltam < 0 Then
            c.String = "Vencido"
            c.CellBackColor = RGB(244, 150, 150)
        ElseIf faltam <= DIAS_ALERTA Then
            c.String = "Vence em " & faltam & " dia(s)"
            c.CellBackColor = RGB(255, 235, 130)
        Else
            c.String = "No prazo"
            c.CellBackColor = RGB(170, 220, 170)
        End If
    Next i
End Sub
