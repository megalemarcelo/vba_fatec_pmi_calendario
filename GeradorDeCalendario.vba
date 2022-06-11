Attribute VB_Name = "GeradorDeCalendario"
'Autor: MARCELO ACERBI MEGALE
'Linguagem: VBA (Visual Basic for Appplications)
'Data: 27/04/2021
'Programa: Gerador de Calendario
'Descricao: O programa recebe, via Input Box, valores para mes e ano. Com base nesses dados, gera um calendario.

Sub Calendario()

'Declaracao de variaveis
Dim mes, ano, semana1, semanan, dia1, dian, linha, coluna, opcao As Integer
Dim titulo As String

'Configuracao da variavel da opcao para iniciar o programa
opcao = 1

'Inicio do loop para opcao 1
Do While (opcao = 1)

'InputBox do Menu para receber opcao
opcao = InputBox("Digite uma opcao: " & Chr(13) & "1) Gerar calendario" & Chr(13) & "2) Encerrar", "Calendario")
    
    'Mensagem de encerramento do programa
    If opcao = 2 Then
        MsgBox "Programa finalizado", , "Calendario"
        Else
        
        'Limpar variaveis mes e ano
        mes = 0
        ano = 0
        
            'InputBox para receber a entrada do mes
            Do While (mes > 12 Or mes < 1)
                mes = InputBox("Digite o mes (de 1 a 12):", "Mes")
            Loop
            
            'InputBox para receber a entrada do ano
            Do While (ano < 1)
                ano = InputBox("Digite o ano:", "Ano")
            Loop
            
            'Select Case para o titulo do calendario (concatenacao de string)
            Select Case mes
            Case 1
            titulo = "JANEIRO - " & ano
            Case 2
            titulo = "FEVEREIRO - " & ano
            Case 3
            titulo = "MAR�O - " & ano
            Case 4
            titulo = "ABRIL - " & ano
            Case 5
            titulo = "MAIO - " & ano
            Case 6
            titulo = "JUNHO - " & ano
            Case 7
            titulo = "JULHO - " & ano
            Case 8
            titulo = "AGOSTO - " & ano
            Case 9
            titulo = "SETEMBRO - " & ano
            Case 10
            titulo = "OUTUBRO - " & ano
            Case 11
            titulo = "NOVEMBRO - " & ano
            Case 12
            titulo = "DEZEMBRO - " & ano
            End Select
            
            'Escreve o titulo do calendario
            Cells(14, 2).Value = titulo
            
            'Calculo dos valores a serem utilizados para gerar o calendario
            dia1 = 1
            dian = Day(DateSerial(ano, mes + 1, dia1) - 1)
            semana1 = WorksheetFunction.WeekNum(DateSerial(ano, mes, dia1))
            semanan = WorksheetFunction.WeekNum(DateSerial(ano, mes + 1, dia1) - 1)
            linha = 16
            
            'Limpar as celulas de B16 a H21
            range("B16:H21").ClearContents
            
            'Inicio do loop para preencher as celulas dos dias no calendario
                Do While (semana1 <= semanan)
                    coluna = WorksheetFunction.Weekday(DateSerial(ano, mes, dia1))
                    coluna = coluna + 1
                    
                    'Se o ultimo dia for maior ou igual ao dia corrente, somar 1 e continuar para a proxima celula
                    If dian >= dia1 Then Cells(linha, coluna).Value = dia1
                    dia1 = dia1 + 1
                    
                    'Se a coluna for 8 (ultima coluna), entao adicionar mais uma semana e usar a proxima linha
                    If (coluna = 8) Then
                        semana1 = semana1 + 1
                        linha = linha + 1
                    End If
                Loop
                
                'Configuracao da opcao para continuar o loop
                opcao = 1
    End If
    
Loop

End Sub
Sub Limpar()
    
    'Limpar celulas dos dias no calendario
    range("B16:H21").Select
    Selection.ClearContents
    
    'Limpar titulo do calendario
    range("B14").Select
    Selection.ClearContents
    
End Sub
