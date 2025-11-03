# Jogo-de-Cartas-21-em-VBA
Este repositório contém o código fonte de um jogo de cartas chamado "21", desenvolvido por mim e meus amigos como parte de um projeto da aula de Programação em Microinformática, sob a orientação do Professor Avelino Palma Pimenta Junior na Escola Estadual Fatec de Araraquara. Agradecemos por sua avaliação que resultou em uma nota 9.0!
Sobre o Projeto

Option Explicit

Dim cartas(1 To 14) As String
Dim proxColunaP1 As Integer
Dim proxColunaP2 As Integer
' Função para converter carta em valor numérico
Function ValorCarta(carta As String) As Integer
    Select Case carta
        Case "J", "Q", "K"
            ValorCarta = 10
        Case "A"
            ValorCarta = 11
        Case Else
            ValorCarta = CInt(carta)
    End Select
End Function

' Cadastro dos jogadores
Sub jogadores()
    Dim player_1 As String, player_2 As String
    
    player_1 = InputBox("Digite o nome do primeiro jogador:")
    player_2 = InputBox("Digite o nome do segundo jogador:")
    
    Range("B2").Value = player_1
    Range("B3").Value = player_2
End Sub

' Sorteio inicial de 2 cartas
Sub deck()
    Randomize
    
    Dim inicio As Integer
    For inicio = 1 To 10
        cartas(inicio) = CStr(inicio)
    Next inicio
    
    cartas(11) = "A"
    cartas(12) = "Q"
    cartas(13) = "J"
    cartas(14) = "K"
    
    ' Limpa mesa
    Range("C2:Z3").ClearContents
    Range("E2:E3").ClearContents
    
    ' Sorteia cartas iniciais
    Range("C2").Value = cartas(Int(Rnd * 14) + 1)
    Range("D2").Value = cartas(Int(Rnd * 14) + 1)
    Range("C3").Value = cartas(Int(Rnd * 14) + 1)
    Range("D3").Value = cartas(Int(Rnd * 14) + 1)
    
    ' Calcula pontos iniciais
    Range("E2").Value = ValorCarta(Range("C2").Value) + ValorCarta(Range("D2").Value)
    Range("E3").Value = ValorCarta(Range("C3").Value) + ValorCarta(Range("D3").Value)
    
    ' Próximas colunas livres para cada jogador
    proxColunaP1 = 5 ' começa em "F", já que C e D estão ocupadas
    proxColunaP2 = 5
End Sub

' Jogador pede carta
Sub Hit()
    Dim jogador As String
    jogador = InputBox("Digite 1 para " & Range("B2").Value & " ou 2 para " & Range("B3").Value)
    
    Dim novaCarta As String
    
    novaCarta = cartas(Int(Rnd * 14) + 1)
    
    If jogador = "1" Then
        Cells(2, proxColunaP1 + 1).Value = novaCarta
        Range("E2").Value = Range("E2").Value + ValorCarta(novaCarta)
        proxColunaP1 = proxColunaP1 + 1
        If Range("E2").Value > 21 Then MsgBox Range("B2").Value & " estourou!"
    ElseIf jogador = "2" Then
        Cells(3, proxColunaP2 + 1).Value = novaCarta
        Range("E3").Value = Range("E3").Value + ValorCarta(novaCarta)
        proxColunaP2 = proxColunaP2 + 1
        If Range("E3").Value > 21 Then MsgBox Range("B3").Value & " estourou!"
    End If
End Sub

' Verifica vencedor
Sub VerificarVencedor()
    Dim p1 As Integer, p2 As Integer
    p1 = Range("E2").Value
    p2 = Range("E3").Value
    
    If p1 > 21 And p2 > 21 Then
        MsgBox "Os dois estouraram. Ninguém ganhou!"
    ElseIf p1 > 21 Then
        MsgBox Range("B3").Value & " venceu!"
    ElseIf p2 > 21 Then
        MsgBox Range("B2").Value & " venceu!"
    ElseIf p1 = p2 Then
        MsgBox "Empate!"
    ElseIf p1 > p2 Then
        MsgBox Range("B2").Value & " venceu!"
    Else
        MsgBox Range("B3").Value & " venceu!"
    End If
    
End Sub



