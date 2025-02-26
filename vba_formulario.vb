Sub Cadastrar()
' Autor: Adryan Santos
' Versão: 2.0


    Dim wsFormulario As Worksheet ' Planilha de origem
    Dim wsBase As Worksheet       ' Planilha de destino
    Dim destinoRow As Long        ' Linha onde os dados serão colados

    ' Variáveis para armazenar as informações do formulário
    Dim dataHoraCadastro As String
    Dim nomeAtendente As String
    Dim enderecoInvalido As String
    Dim codigoH7 As String, codigoL7 As String, codigoAC7 As String
    Dim codigoC10 As String, codigoC12 As String, concatenacaocod As String
    Dim codigoK15 As String, codigoAF15 As String, codigoAK18 As String
    Dim codigoK18 As String, codigoX18 As String, codigoN29 As String
    Dim codigoS22 As String, codigoAM22 As String, codigoCad As String
    Dim codigoN25 As String, codigoN27 As String
    Dim codigoC31 As String, codigoB2 As String

    ' Variáveis para os CheckBoxes
    Dim chk1Atendimento As Boolean, chk2Atendimento As Boolean, chk3Atendimento As Boolean
    Dim chkRetorno As Boolean, chkReceptivo As Boolean

    Dim resultadoColunaV As String, resultadoColunaK As String
    Dim ultimoCodigo As String, novoCodigo As String
    Dim ultimaLinha As Long, numero As Long

    ' Define as planilhas
    Set wsFormulario = ThisWorkbook.Sheets("FORMULÁRIO")

    ' Verifica o status dos CheckBoxes de atendimento e outros
    chk1Atendimento = wsFormulario.Shapes("Check Box 1").ControlFormat.Value = 1
    chk2Atendimento = wsFormulario.Shapes("Check Box 2").ControlFormat.Value = 1
    chk3Atendimento = wsFormulario.Shapes("Check Box 3").ControlFormat.Value = 1
    chkReceptivo = wsFormulario.Shapes("Check Box 26").ControlFormat.Value = 1
    chkRetorno = wsFormulario.Shapes("Check Box 33").ControlFormat.Value = 1

  ' Verifica se múltiplos atendimentos foram selecionados
    Dim atendimentosSelecionados As Integer
    atendimentosSelecionados = 0
    
    If chk1Atendimento Then atendimentosSelecionados = atendimentosSelecionados + 1
    If chk2Atendimento Then atendimentosSelecionados = atendimentosSelecionados + 1
    If chk3Atendimento Then atendimentosSelecionados = atendimentosSelecionados + 1
    If chkReceptivo Then atendimentosSelecionados = atendimentosSelecionados + 1
    If chkRetorno Then atendimentosSelecionados = atendimentosSelecionados + 1
    
    ' Verifica se há mais de um atendimento selecionado
    If atendimentosSelecionados > 1 Then
        MsgBox "Você só pode selecionar um tipo de atendimento por vez. Por favor, revise a seleção.", vbExclamation, "Erro de Seleção"
        Exit Sub
    End If

    ' Define a planilha de destino com base no CheckBox "Receptivo"
    If chkReceptivo Then
        Set wsBase = ThisWorkbook.Sheets("BASE RECEPTIVO")
    Else
        Set wsBase = ThisWorkbook.Sheets("BASE ATENDIMENTO")
    End If

    ' Validação inicial
    If Not chkReceptivo Then
        If wsFormulario.Range("H7").Value = "" Or wsFormulario.Range("O12").Value = "" Or wsFormulario.Range("K18").Value = "" Or wsFormulario.Range("AM22").Value = "" Or wsFormulario.Range("N25").Value = "" Or wsFormulario.Range("N27").Value = "" Or wsFormulario.Range("N29").Value = "" Or wsFormulario.Range("C31").Value = "" Or wsFormulario.Range("AL29").Value = "" Or wsFormulario.Range("AT29").Value = "" Or wsFormulario.Range("Y18").Value = "" Or wsFormulario.Range("C31").Value = "" Then
            MsgBox "Inserir Todas as Informações Corretas", vbExclamation, "Erro de Cadastro"
            Exit Sub
        End If
    End If
    
        
        
    ' Verifica qual atendimento está sendo cadastrado
    Dim atendimentoAtual As String
    If chk1Atendimento Then
        atendimentoAtual = "1º Atendimento"
    ElseIf chk2Atendimento Then
        atendimentoAtual = "2º Atendimento"
    ElseIf chk3Atendimento Then
        atendimentoAtual = "3º Atendimento"
    ElseIf chkReceptivo Then
        atendimentoAtual = "Receptivo"
    ElseIf chkRetorno Then
        atendimentoAtual = "Retorno"
    Else
        MsgBox "Selecione um atendimento válido.", vbExclamation, "Erro"
        Exit Sub
    End If



' Verifica se o "3º Atendimento" foi assinalado e se há "aguardando retorno" no formulário
If chk3Atendimento Then
    Dim celulasAguardandoRetorno As Boolean
    celulasAguardandoRetorno = False

    ' Verifica se alguma das células específicas contém a expressão "Aguardando Retorno"
    If LCase(wsFormulario.Range("N25").Value) Like "*aguardando retorno*" Or _
       InStr(1, LCase(wsFormulario.Range("O12").Value), "aguardando retorno") > 0 Or _
       InStr(1, LCase(wsFormulario.Range("K18").Value), "aguardando retorno") > 0 Or _
       InStr(1, LCase(wsFormulario.Range("AN25").Value), "aguardando retorno") > 0 Or _
       InStr(1, LCase(wsFormulario.Range("N29").Value), "aguardando retorno") > 0 Then
        celulasAguardandoRetorno = True
    End If

    ' Se encontrar "aguardando retorno", impede o registro
    If celulasAguardandoRetorno Then
        MsgBox "O 3º Atendimento não pode ser 'aguardando retorno'. Corrija as informações antes de prosseguir.", vbExclamation, "Erro de Validação"
        Exit Sub
    End If
End If

    ' Verifica se já existe o atendimento 1, 2 ou 3 para o parceiro na ordem correta
    Dim parceiroID As String
    Dim atendimentoExistente As Range
    Dim atendimentosRegistrados As String
    Dim primeiraOcorrencia As Range

    ' Obtém o ID do parceiro atual
    parceiroID = Trim(wsFormulario.Range("H7").Value)

    ' Verifica na coluna E (ID Parceiro)
    Set atendimentoExistente = wsBase.Columns("E").Find(What:=parceiroID, LookIn:=xlValues, LookAt:=xlWhole)

    ' Se nenhum registro do parceiro for encontrado
    If atendimentoExistente Is Nothing Then
        ' Permite o cadastro apenas se for o 1º Atendimento, Receptivo ou Retorno
        If chk1Atendimento Or chkReceptivo Or chkRetorno Then
            ' Prossegue normalmente
        Else
            MsgBox "Não é possível cadastrar o atendimento sem que o 1º Atendimento seja registrado antes.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        End If
    Else
   ' Concatena os registros de atendimento de todas as ocorrências do parceiro (Somente Números)
Set primeiraOcorrencia = atendimentoExistente
atendimentosRegistrados = ""  ' Inicializa a variável
    
' Percorre todas as ocorrências do ID do parceiro
Do
    ' Extrai somente números dos atendimentos (remove palavras como "º Atendimento")
    Dim atendimentoNumero As String
    atendimentoNumero = atendimentoExistente.Offset(0, 18).Value
    atendimentoNumero = Replace(Replace(Replace(Replace(atendimentoNumero, "º", ""), "Atendimento", ""), " ", ""), "Retorno", "4")
    
    ' Concatena apenas os números em uma string simples (ex: "123")
    atendimentosRegistrados = atendimentosRegistrados & atendimentoNumero
    
    ' Busca a próxima ocorrência
    Set atendimentoExistente = wsBase.Columns("E").FindNext(after:=atendimentoExistente)
Loop While Not atendimentoExistente Is Nothing And atendimentoExistente.Address <> primeiraOcorrencia.Address

' Remove espaços extras e limpa qualquer formatação desnecessária
atendimentosRegistrados = Trim(atendimentosRegistrados)

    
    ' Remove espaços extras e limpa qualquer formatação desnecessária
    atendimentosRegistrados = Trim(atendimentosRegistrados)
    
   ' **Nova Verificação:** Permitir "Retorno" mesmo após os 3 atendimentos
If InStr(1, atendimentosRegistrados, "1") > 0 And _
   InStr(1, atendimentosRegistrados, "2") > 0 And _
   InStr(1, atendimentosRegistrados, "3") > 0 Then
    
    ' Permitir o cadastro se for um retorno
    If chkRetorno Then
        GoTo ContinuarCadastro ' Permite o "Retorno" e segue
    Else
        MsgBox "Já existem um 1º atendimento, um 2º atendimento e um 3º atendimento para este parceiro, favor acionar a gestão.", vbExclamation, "Ação Necessária"
        Exit Sub
    End If
End If

' **Rótulo de Continuação (Para Retorno):**
ContinuarCadastro:

End If

' **Nova Verificação usando apenas números**
Dim possui1 As Boolean, possui2 As Boolean, possui3 As Boolean
possui1 = InStr(1, atendimentosRegistrados, "1") > 0
possui2 = InStr(1, atendimentosRegistrados, "2") > 0
possui3 = InStr(1, atendimentosRegistrados, "3") > 0

' **Verificação 1º Atendimento**
If chk1Atendimento And possui1 Then
    MsgBox "Este parceiro já possui o 1º Atendimento registrado. Não é possível cadastrar um novo 1º Atendimento.", vbExclamation, "Erro de Cadastro"
    Exit Sub
End If

' **Verificação 2º Atendimento**
If chk2Atendimento Then
    If Not possui1 Then
        MsgBox "O 1º Atendimento deve ser registrado antes do 2º Atendimento.", vbExclamation, "Erro de Cadastro"
        Exit Sub
    ElseIf possui2 Then
        MsgBox "Este parceiro já possui o 2º Atendimento registrado. Não é possível cadastrar novamente.", vbExclamation, "Erro de Cadastro"
        Exit Sub
    End If
End If

' **Verificação 3º Atendimento**
If chk3Atendimento Then
    If Not possui1 Then
        MsgBox "O 1º Atendimento deve ser registrado antes do 3º Atendimento.", vbExclamation, "Erro de Cadastro"
        Exit Sub
    ElseIf Not possui2 Then
        MsgBox "O 2º Atendimento deve ser registrado antes do 3º Atendimento.", vbExclamation, "Erro de Cadastro"
        Exit Sub
    ElseIf possui3 Then
        MsgBox "Este parceiro já possui o 3º Atendimento registrado. Não é possível cadastrar novamente.", vbExclamation, "Erro de Cadastro"
        Exit Sub
    End If
End If





' **PASSO 2: Verificação específica por atendimento**
Select Case True
    Case chk1Atendimento
        If InStr(1, atendimentosRegistrados, "1") > 0 Then
            MsgBox "Este parceiro já possui o 1º Atendimento registrado. Não é possível cadastrar um novo 1º Atendimento.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        End If

    Case chk2Atendimento
        If InStr(1, atendimentosRegistrados, "1") = 0 Then
            MsgBox "O 1º Atendimento deve ser registrado antes do 2º Atendimento.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        ElseIf InStr(1, atendimentosRegistrados, "2") > 0 Then
            MsgBox "Este parceiro já possui o 2º Atendimento registrado. Não é possível cadastrar novamente.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        End If

    Case chk3Atendimento
        If InStr(1, atendimentosRegistrados, "1") = 0 Then
            MsgBox "O 1º Atendimento deve ser registrado antes do 3º Atendimento.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        ElseIf InStr(1, atendimentosRegistrados, "2") = 0 Then
            MsgBox "O 2º Atendimento deve ser registrado antes do 3º Atendimento.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        ElseIf InStr(1, atendimentosRegistrados, "3") > 0 Then
            MsgBox "Este parceiro já possui o 3º Atendimento registrado. Não é possível cadastrar novamente.", vbExclamation, "Erro de Cadastro"
            Exit Sub
        End If
End Select


    ' Define a linha onde os dados serão colados (inserir no topo)
    With wsBase.Rows(5)
        .Insert Shift:=xlUp, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
    destinoRow = 5

    ' O restante do código permanece como está...
    ' Armazena os valores do formulário em variáveis e preenche a planilha de destino
    ' (...)


' Armazena os valores do formulário em variáveis
Dim dataAtual As Date
dataAtual = DateSerial(Year(Now), Month(Now), Day(Now)) + TimeValue(Now)
dataHoraCadastro = Format(dataAtual, "dd/mm/yyyy hh:mm")



celulaqtratou = "Setor"
nomeAtendente = "Usuário"
codigoH7 = wsFormulario.Range("H7").Value ' ID Parceiro
codigoL7 = wsFormulario.Range("L7").Value ' Razão Social
codigoAC7 = wsFormulario.Range("AC7").Value ' Nome Fantasia
codigoC10 = wsFormulario.Range("C10").Value ' Endereço
codigoAV10 = wsFormulario.Range("AV10").Value ' uf parceiro
codigoO12 = wsFormulario.Range("O12").Value ' Endereço Correto?!
codigoK15 = wsFormulario.Range("K15").Value ' Telefone 1
codigoAF15 = wsFormulario.Range("AF15").Value ' Telefone 2
codigoK18 = wsFormulario.Range("K18").Value ' Contato Atualizado?!
codigoY18 = wsFormulario.Range("Y18").Value ' Obs Contato
codigoS22 = wsFormulario.Range("S22").Value ' Estoque Sistemico
codigoAM22 = wsFormulario.Range("AM22").Value ' Estoque Validado
codigoN25 = wsFormulario.Range("N25").Value ' Status do Atendimento
codigoN27 = wsFormulario.Range("N27").Value ' Protocolo Atendimento
codigoAN25 = wsFormulario.Range("AN25").Value ' Envio Decl
codigoN29 = wsFormulario.Range("N29").Value ' Demanda
codigoC31 = wsFormulario.Range("C31").Value ' Observação do Atendimento
codigoAL29 = wsFormulario.Range("AL29").Value ' Valor Acordado?
codigoAT29 = wsFormulario.Range("AT29").Value ' Quanto?
codigoBL32 = wsFormulario.Range("BL32").Value ' Codigo 1
codigoCad = Day(wsFormulario.Range("B2").Value) & Month(wsFormulario.Range("B2").Value) & Year(wsFormulario.Range("B2").Value) & "-" & wsFormulario.Range("H7").Value ' Codigo 2
codigoB2 = wsFormulario.Range("B2").Value ' Maior Data de Atendimento
concatenacaocod = wsFormulario.Range("H7").Value & "-" & atendimentoAtual

' Verifica o status dos CheckBoxes de atendimento e outros (Form Controls)
chk1Atendimento = wsFormulario.Shapes("Check Box 1").ControlFormat.Value = 1
chk2Atendimento = wsFormulario.Shapes("Check Box 2").ControlFormat.Value = 1
chk3Atendimento = wsFormulario.Shapes("Check Box 3").ControlFormat.Value = 1
chkReceptivo = wsFormulario.Shapes("Check Box 26").ControlFormat.Value = 1
chkRetorno = wsFormulario.Shapes("Check Box 33").ControlFormat.Value = 1

' Preenche os dados na planilha de destino
wsBase.Cells(destinoRow, "B").Value = dataAtual
wsBase.Cells(destinoRow, "B").NumberFormat = "dd/mm/yyyy hh:mm"


wsBase.Cells(destinoRow, "C").Value = celulaqtratou
wsBase.Cells(destinoRow, "D").Value = nomeAtendente
wsBase.Cells(destinoRow, "E").Value = codigoH7 ' Id Parceiro
wsBase.Cells(destinoRow, "F").Value = codigoL7 ' Razão Social
wsBase.Cells(destinoRow, "G").Value = codigoAC7 ' Nome Fantasia
wsBase.Cells(destinoRow, "H").Value = codigoAV10 ' UF Parceiro
wsBase.Cells(destinoRow, "I").Value = codigoC10 ' Endereço
wsBase.Cells(destinoRow, "J").Value = codigoO12 ' Endereço Atualizado?!
wsBase.Cells(destinoRow, "K").Value = codigoK15 ' Telefone 1
wsBase.Cells(destinoRow, "L").Value = codigoAF15 ' Telefone 2
wsBase.Cells(destinoRow, "N").Value = codigoK18 ' Contato Atualizado
wsBase.Cells(destinoRow, "O").Value = codigoY18 ' Observação Contato
wsBase.Cells(destinoRow, "P").Value = codigoS22 ' Estoque Sistêmico
wsBase.Cells(destinoRow, "Q").Value = codigoAM22 ' Estoque Validado
wsBase.Cells(destinoRow, "R").Value = codigoN25 ' Status do Atendimento
wsBase.Cells(destinoRow, "S").Value = codigoAN25 ' Envio Decl
wsBase.Cells(destinoRow, "T").Value = codigoN27 ' Protocolo de Atendimento
wsBase.Cells(destinoRow, "U").Value = codigoN29 ' Demanda
wsBase.Cells(destinoRow, "V").Value = codigoC31 ' Observação Atendimento
wsBase.Cells(destinoRow, "X").Value = codigoBL32 ' Código sequencial
wsBase.Cells(destinoRow, "Y").Value = codigoAL29 ' Valor Acordado?
wsBase.Cells(destinoRow, "Z").Value = codigoAT29 ' Quanto?
wsBase.Cells(destinoRow, "AB").Value = concatenacaocod
Dim dataUltimoAtendimento As Date
dataUltimoAtendimento = DateSerial(Year(codigoB2), Month(codigoB2), Day(codigoB2))
wsBase.Cells(destinoRow, "AC").Value = dataUltimoAtendimento
wsBase.Cells(destinoRow, "AC").NumberFormat = "dd/mm/yyyy"

wsBase.Cells(destinoRow, "AD").Value = codigoCad ' Codigo ELI




    ' Preenche a coluna W com os CheckBoxes de Atendimento
    resultadoColunaW = ""
    If chk1Atendimento Then resultadoColunaW = resultadoColunaW & "1º Atendimento"
    If chk2Atendimento Then resultadoColunaW = resultadoColunaW & "2º Atendimento"
    If chk3Atendimento Then resultadoColunaW = resultadoColunaW & "3º Atendimento"
    If chkRetorno Then resultadoColunaW = resultadoColunaW & "Retorno"
    If resultadoColunaW <> "" Then
        resultadoColunaW = Left(resultadoColunaW, Len(resultadoColunaW))
        wsBase.Cells(destinoRow, "W").Value = resultadoColunaW
    End If
    
    ' Mensagem de sucesso
    MsgBox "Atendimento realizado com sucesso!", vbInformation, "Cadastro Concluído"

End Sub

